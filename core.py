"""core.py - logika inti tanpa UI (bisa diuji tanpa Streamlit/Gemini).

Alur:
1. read_any / detect_columns      -> baca file SAP (xlsx/csv) dan tebak kolomnya
2. clean_text                     -> teks yang MASUK ke template (perbaiki karakter rusak, spasi)
3. match_text                     -> teks INTERNAL untuk pencocokan (sinonim); tidak masuk template
4. classify_rows                  -> label contoh -> aturan -> AI (opsional) -> REVIEW
5. validate_template1 / build_template1_bytes -> validasi & file Template 1

Perbaikan kestabilan AI:
- Retry bertingkat untuk error sementara (503, 429, 5xx, timeout, koneksi).
- Error permanen tidak diulang berkali-kali.
- Kegagalan satu batch tidak menghentikan batch berikutnya.
- Baris yang tidak mendapat jawaban AI tetap masuk REVIEW dengan catatan yang jelas.
"""
from __future__ import annotations

import io
import random
import re
import time

import openpyxl
import pandas as pd


# ----------------------------------------------------------------------------
# 1. Membaca file
# ----------------------------------------------------------------------------
def read_any(name: str, raw: bytes) -> pd.DataFrame:
    """Baca xlsx atau csv. CSV dicoba UTF-8 dulu, lalu cp1252."""
    if name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1252"):
            try:
                return pd.read_csv(
                    io.BytesIO(raw), dtype=str, encoding=enc,
                    keep_default_na=False
                )
            except UnicodeDecodeError:
                continue
        return pd.read_csv(
            io.BytesIO(raw), dtype=str, encoding="latin1",
            keep_default_na=False
        )
    return pd.read_excel(io.BytesIO(raw), dtype=str).fillna("")


def detect_columns(cols) -> dict:
    """Tebak kolom kode / short / long / uom dari nama header ekspor SAP."""
    low = {c: str(c).lower() for c in cols}

    def pick(pred):
        for c, lc in low.items():
            if pred(lc):
                return c
        return None

    short = pick(lambda x: "description" in x or "short" in x)
    long_ = pick(lambda x: "po text" in x or "long" in x)
    uom = pick(lambda x: "unit" in x or "uom" in x)
    code = pick(
        lambda x: ("material" in x or "code" in x)
        and "description" not in x
    )

    guess = [code, short, long_, uom]
    fallback = list(cols)[:4]
    return {
        "code": code or (fallback[0] if fallback else None),
        "short": short or (fallback[1] if len(fallback) > 1 else (fallback[0] if fallback else None)),
        "long": long_ or (fallback[2] if len(fallback) > 2 else (fallback[0] if fallback else None)),
        "uom": uom or (fallback[3] if len(fallback) > 3 else (fallback[0] if fallback else None)),
        "_guessed": all(guess),
    }


def load_master(df_raw: pd.DataFrame):
    """Material Standard -> (df_nm: pasangan noun|modifier, df_ms: semua atribut)."""
    df = df_raw.fillna("").copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    for c in ("NOUN", "MODIFIER", "MANDATORY", "ATTRIBUTE"):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.upper()

    if "SEQ" in df.columns:
        df["SEQ"] = df["SEQ"].astype(str).str.strip().astype(int)

    required = {"NOUN", "MODIFIER"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Kolom wajib tidak ditemukan di Material Standard: "
            + ", ".join(sorted(missing))
        )

    df_nm = (
        df[["NOUN", "MODIFIER"]]
        .drop_duplicates()
        .sort_values(["NOUN", "MODIFIER"])
        .reset_index(drop=True)
    )
    return df_nm, df


# ----------------------------------------------------------------------------
# 2. Pembersihan teks (ini yang masuk ke template)
# ----------------------------------------------------------------------------
# Urutan penting: "Â°" harus diperbaiki sebelum "Â" dibuang.
CHAR_FIXES = [
    ("Â°", "°"),
    ("GÇ¥", '"'),
    ("GÇô", "-"),
    ("GÇ£", '"'),
    ("Â", ""),
    ("”", '"'), ("“", '"'),
    ("’", "'"), ("‘", "'"),
    ("–", "-"), ("—", "-"),
    ("\xa0", " "),
]
ASCII_SYMBOLS = [
    ("®", ""), ("™", ""), ("•", "-"),
    ("±", "+/-"), ("³", "3")
]


def clean_text(s, ascii_symbols: bool = False) -> str:
    s = "" if s is None else str(s)

    for a, b in CHAR_FIXES:
        s = s.replace(a, b)

    # 5½ -> 5-1/2, lalu ½ -> 1/2
    for ch, frac in (("½", "1/2"), ("¼", "1/4"), ("¾", "3/4")):
        s = re.sub(rf"(\d){ch}", rf"\1-{frac}", s)
        s = s.replace(ch, frac)

    if ascii_symbols:
        for a, b in ASCII_SYMBOLS:
            s = s.replace(a, b)

    s = re.sub(r"[\t\r\n]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


# ----------------------------------------------------------------------------
# 3. Teks internal untuk pencocokan (kamus sinonim).
#    Hasil fungsi ini TIDAK dimasukkan ke template.
# ----------------------------------------------------------------------------
SYNONYMS = [
    (r"\bPUP ?JOINTS?\b", "PUP JOINT"),
    (r"\bTBG\b", "TUBING"),
    (r"\bCSG\b", "CASING"),
    (r"\bX-?OVER\b|\bCROSS ?OVER\b", "CROSSOVER"),
    (r"\bHAND ?GLOVES?\b", "HAND GLOVE"),
    (r"\b(?:V-?BELTS?|BELT-?V)\b", "BELT V"),
    (r"\bELEC(?:TRIC)?\b", "ELECTRICAL"),
    (r"\bPNEU\b", "PNEUMATIC"),
    (r"\bTRUST\b", "THRUST"),
]


def match_text(s) -> str:
    t = str(s or "").upper()
    for pat, rep in SYNONYMS:
        t = re.sub(pat, rep, t)
    return t


# ----------------------------------------------------------------------------
# 4. Klasifikasi
# ----------------------------------------------------------------------------
def build_context(
    df_nm: pd.DataFrame,
    labeled: pd.DataFrame | None = None
) -> dict:
    """Siapkan struktur pencarian dari master dan file contoh berlabel.

    `labeled` = Template 2 berisi kolom:
    PSC Material Code, Noun, Modifier, Material - Long Description.
    """
    ctx = {
        "valid": set(zip(df_nm["NOUN"], df_nm["MODIFIER"])),
        "nouns": set(df_nm["NOUN"]),
        "mods_by_noun": df_nm.groupby("NOUN")["MODIFIER"].apply(list).to_dict(),
        "pairs_text": "\n".join(
            f"{n} | {m}" for n, m in zip(df_nm["NOUN"], df_nm["MODIFIER"])
        ),
        "labels": {},
        "index": [],
    }

    if labeled is not None and not labeled.empty:
        lab = labeled.fillna("")
        required = {
            "PSC Material Code", "Noun", "Modifier",
            "Material - Long Description"
        }
        missing = required - set(lab.columns)
        if missing:
            raise ValueError(
                "Kolom wajib Template 2 tidak ditemukan: "
                + ", ".join(sorted(missing))
            )

        first = lab[
            lab["Material - Long Description"].astype(str).str.strip() != ""
        ]

        for _, r in first.iterrows():
            code = str(r["PSC Material Code"]).strip()
            noun = str(r["Noun"]).strip().upper()
            mod = str(r["Modifier"]).strip().upper()
            text = str(r["Material - Long Description"])
            ctx["labels"][code] = (noun, mod)
            ctx["index"].append((_tokens(text), text, noun, mod))

    return ctx


def _tokens(text: str) -> frozenset:
    return frozenset(re.findall(r"[A-Z]{3,}", match_text(text)))


def similar_examples(
    text: str,
    ctx: dict,
    k: int = 3,
    exclude_text: str | None = None
) -> list:
    """k contoh berlabel paling mirip (Jaccard pada kata) untuk few-shot AI."""
    q = _tokens(text)
    if not q or not ctx["index"]:
        return []

    scored = []
    for toks, t, noun, mod in ctx["index"]:
        if exclude_text is not None and t == exclude_text:
            continue
        inter = len(q & toks)
        if inter:
            scored.append((inter / len(q | toks), t, noun, mod))

    scored.sort(key=lambda x: -x[0])
    return [(t, n, m) for _, t, n, m in scored[:k]]


def _first_token(t: str) -> str:
    return re.split(r"[,;]", t)[0].strip()


def rule_classify(short: str, long: str, ctx: dict):
    """Aturan berpresisi tinggi. Return tuple atau None.

    Hasil uji pada 526 material berlabel: sekitar 81% benar untuk sekitar
    35% data. Sisanya sengaja diserahkan ke AI/review.
    """
    s, l = match_text(short), match_text(long)
    comb = f"{s} , {l}"
    valid = ctx["valid"]
    mods_by_noun = ctx["mods_by_noun"]

    # a) CROSSOVER sebagai kata pertama
    if (
        re.match(r"^CROSSOVER\b", _first_token(l))
        or re.match(r"^CROSSOVER\b", _first_token(s))
    ):
        if ("CROSS-OVER", "--") in valid:
            return (
                "CROSS-OVER", "--", 90, "ATURAN",
                "Kata pertama CROSSOVER/X-OVER"
            )

    # b) PUP JOINT: modifier ditentukan kata CASING / TUBING
    if re.search(r"\bPUP\b", comb) and re.search(
        r"\bPUP JOINT\b|\bPUP\b.*\b(CASING|TUBING)\b|"
        r"\b(CASING|TUBING)\b.*\bPUP\b",
        comb
    ):
        mod = (
            "TUBING"
            if re.search(r"\bTUBING\b", comb)
            and not re.search(r"\bCASING\b", comb)
            else "CASING"
        )
        if ("PUP JOINT", mod) in valid:
            return (
                "PUP JOINT", mod, 78, "ATURAN",
                f"PUP JOINT + kata {mod}"
            )

    # c) Kata pertama = noun master dan modifier resmi tertulis di teks
    for tok in (_first_token(l), _first_token(s)):
        if tok in mods_by_noun:
            for mod in sorted(mods_by_noun[tok], key=len, reverse=True):
                if mod != "--" and re.search(
                    rf"\b{re.escape(mod)}\b", comb
                ):
                    return (
                        tok, mod, 80, "ATURAN",
                        f"Noun + modifier tertulis ({tok} | {mod})"
                    )

    # Noun dengan satu modifier saja sengaja tidak diterima otomatis.
    return None


# ---- AI (Gemini) ------------------------------------------------------------
class AIFatalError(Exception):
    """Error AI yang tidak akan sembuh dengan mengulang (kunci, model, kuota)."""


def is_fatal_error(error: Exception) -> bool:
    """Kunci salah, model tidak ada/sudah dimatikan, izin ditolak, kuota harian habis."""
    status = getattr(error, "status_code", None) or getattr(error, "code", None)
    try:
        status = int(status)
    except (TypeError, ValueError):
        status = None
    if status in (400, 401, 403, 404):
        return True
    msg = str(error).upper()
    fatal_terms = (
        "ACCESS_TOKEN_TYPE_UNSUPPORTED", "UNAUTHENTICATED", "API_KEY_INVALID",
        "API KEY NOT VALID", "PERMISSION_DENIED", "NO LONGER AVAILABLE",
        "NOT_FOUND", "PERDAY", "PER_DAY", "LIMIT: 0",
    )
    return any(term in msg for term in fatal_terms)


def make_gemini_batch_fn(
    api_key: str,
    model: str,
    ctx: dict,
    retries: int = 5,
    base_delay: float = 2,
    max_delay: float = 30,
):
    """Kembalikan fungsi f(payload) -> {row_index: (noun, modifier, conf, alasan)}.

    Retry hanya dilakukan untuk error sementara seperti 429, 5xx, timeout,
    dan gangguan koneksi. Error autentikasi/permintaan tidak diulang.
    Import Gemini dilakukan di sini agar core.py tetap bisa diuji tanpa
    library Gemini.
    """
    from google import genai
    from google.genai import types
    from pydantic import BaseModel

    class Item(BaseModel):
        row_index: int
        noun: str
        modifier: str
        confidence: int
        reason: str

    class Result(BaseModel):
        items: list[Item]

    # Semua format kunci (AIza... maupun AQ....) dikirim lewat api_key.
    # JANGAN memakai header "Authorization: Bearer" untuk kunci API:
    # itu menghasilkan 401 ACCESS_TOKEN_TYPE_UNSUPPORTED.
    client = genai.Client(api_key=api_key.strip())

    def is_retryable(error: Exception) -> bool:
        """Deteksi error sementara dari kode/status maupun pesan error."""
        status = (
            getattr(error, "status_code", None)
            or getattr(error, "code", None)
        )

        try:
            status = int(status)
        except (TypeError, ValueError):
            status = None

        if status in (408, 425, 429, 500, 502, 503, 504):
            return True

        msg = str(error).upper()
        retry_terms = (
            "503", "UNAVAILABLE",
            "429", "RESOURCE_EXHAUSTED",
            "500", "INTERNAL SERVER ERROR",
            "502", "BAD GATEWAY",
            "504", "GATEWAY TIMEOUT",
            "408", "TIMEOUT", "TIMED OUT",
            "CONNECTION RESET", "CONNECTION ERROR",
            "TEMPORARILY", "TEMPORARY",
        )
        return any(term in msg for term in retry_terms)

    def run(payload: list[dict]) -> dict:
        blocks = []

        for p in payload:
            examples = "\n".join(
                f"   - {t[:200]} => {n} | {m}"
                for t, n, m in p.get("examples", [])
            )

            block = (
                f"[row_index {p['row_index']}]\n"
                f"SHORT: {p['short'][:150]}\n"
                f"PO TEXT: {p['long'][:600]}"
            )
            if examples:
                block += (
                    "\nCONTOH SERUPA (sudah dikatalogkan):\n"
                    + examples
                )
            blocks.append(block)

        # Prompt asli katalogisasi dipertahankan.
        prompt = (
            "Anda cataloguer senior SKK Migas. Untuk SETIAP item, pilih "
            "tepat satu pasangan NOUN | MODIFIER dari DAFTAR RESMI.\n"
            "Catatan penting:\n"
            "- SHORT dibatasi 40 karakter dan sering disingkat/terpotong "
            "(PUPJOINT, X-OVER, TBG=tubing, CSG=casing). PO TEXT lebih "
            "lengkap; utamakan PO TEXT.\n"
            "- Tentukan benda utamanya, bukan peralatan tempat ia dipasang.\n"
            "- 'SPARE PARTS | --' sah dipakai untuk suku cadang yang tidak "
            "punya kelas sendiri.\n"
            "- Ikuti pola CONTOH SERUPA bila relevan. Jangan mengarang "
            "pasangan di luar daftar.\n"
            "- confidence 0-100 (jujur; turunkan bila ragu).\n\n"
            "ITEM:\n"
            + "\n\n".join(blocks)
            + "\n\nDAFTAR RESMI:\n"
            + ctx["pairs_text"]
        )

        last_error = None
        total_attempts = max(1, int(retries))

        for attempt in range(total_attempts):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=Result,
                    ),
                )

                if not getattr(response, "text", None):
                    raise ValueError(
                        "Gemini mengembalikan respons kosong."
                    )

                parsed = Result.model_validate_json(response.text)

                return {
                    item.row_index: (
                        item.noun.strip().upper(),
                        item.modifier.strip().upper(),
                        item.confidence,
                        item.reason,
                    )
                    for item in parsed.items
                }

            except Exception as error:
                last_error = error

                # Error fatal tidak sembuh dengan mengulang: hentikan tahap AI.
                if is_fatal_error(error):
                    raise AIFatalError(str(error)[:500]) from error

                # Error nonsementara lain langsung dihentikan.
                if not is_retryable(error):
                    raise

                if attempt < total_attempts - 1:
                    delay = min(
                        float(base_delay) * (2 ** attempt),
                        float(max_delay),
                    )
                    # Jitter agar retry tidak selalu terjadi bersamaan.
                    delay += random.uniform(0, min(1.0, delay * 0.2))
                    time.sleep(delay)

        raise last_error

    return run


def classify_rows(
    rows: list[dict],
    ctx: dict,
    ai_batch=None,
    batch_size: int = 5,
    progress=None
) -> list[dict]:
    """Klasifikasi: label contoh -> aturan -> AI -> REVIEW.

    Kegagalan AI pada satu batch tidak menghentikan batch berikutnya.
    Urutan hasil selalu sama dengan urutan rows.
    """
    results: dict[int, tuple] = {}
    pending: list[int] = []
    errors: dict[int, str] = {}
    fatal_message = None

    # 1. Label contoh -> aturan -> masukkan sisanya ke antrean AI.
    for i, row in enumerate(rows):
        label = ctx["labels"].get(row["code"])

        if label and label in ctx["valid"]:
            results[i] = (
                label[0], label[1], 100,
                "CONTOH", "Label dari file contoh berlabel"
            )
            continue

        rule = rule_classify(
            row["short"], row["long"], ctx
        )
        if rule:
            results[i] = rule
        else:
            pending.append(i)

    # 2. AI per batch. Default 5 material per panggilan.
    if ai_batch is not None and pending:
        safe_batch_size = max(1, int(batch_size))

        for start in range(0, len(pending), safe_batch_size):
            chunk = pending[start:start + safe_batch_size]

            payload = [
                {
                    "row_index": i,
                    "short": rows[i]["short"],
                    "long": rows[i]["long"],
                    "examples": similar_examples(
                        rows[i]["long"] or rows[i]["short"],
                        ctx
                    ),
                }
                for i in chunk
            ]

            try:
                decisions = ai_batch(payload)
                if not isinstance(decisions, dict):
                    raise TypeError(
                        "Hasil AI harus berupa dictionary berdasarkan row_index."
                    )
            except AIFatalError as error:
                decisions = {}
                fatal_message = str(error)[:500]
                for i in chunk:
                    errors[i] = fatal_message
            except Exception as error:
                decisions = {}
                message = str(error)[:500]
                for i in chunk:
                    errors[i] = message

            # 3. Validasi setiap hasil; satu hasil buruk tidak merusak baris lain.
            for i in chunk:
                decision = decisions.get(i)

                if not decision:
                    if i in errors:
                        note = f"Gagal AI: {errors[i]}"
                    else:
                        note = (
                            "AI tidak mengembalikan hasil untuk baris ini."
                        )
                    results[i] = ("", "", 0, "REVIEW", note)
                    continue

                try:
                    noun, modifier, confidence, reason = decision
                    noun = str(noun).strip().upper()
                    modifier = str(modifier).strip().upper()
                    confidence = max(0, min(100, int(confidence)))
                    reason = str(reason)

                    if (noun, modifier) in ctx["valid"]:
                        results[i] = (
                            noun, modifier, confidence, "AI", reason
                        )
                    else:
                        results[i] = (
                            "", "", 0, "REVIEW",
                            "AI mengusulkan pasangan tidak ada di master: "
                            f"{noun} | {modifier}"
                        )

                except (TypeError, ValueError) as error:
                    results[i] = (
                        "", "", 0, "REVIEW",
                        f"Format jawaban AI tidak valid: {error}"
                    )

            if fatal_message:
                # Sisa antrean tidak dikirim ke AI; ditandai REVIEW.
                for j in pending[start + len(chunk):]:
                    results[j] = (
                        "", "", 0, "REVIEW",
                        f"AI dihentikan: {fatal_message}"
                    )
                if progress:
                    progress(len(pending), len(pending))
                break

            if progress:
                progress(
                    min(start + len(chunk), len(pending)),
                    len(pending)
                )

    # 4. Susun hasil sesuai urutan material asli.
    output = []

    for i, _row in enumerate(rows):
        noun, modifier, confidence, source, note = results.get(
            i,
            (
                "", "", 0, "REVIEW",
                "Belum diklasifikasi"
                if ai_batch is None
                else "Belum diklasifikasi"
            )
        )

        output.append({
            "noun": noun,
            "modifier": modifier,
            "confidence": confidence,
            "source": source,
            "note": note,
        })

    return output


# ----------------------------------------------------------------------------
# 5. Template 1: validasi dan pembuatan file
# ----------------------------------------------------------------------------
LIMITS = {
    "code": 20,
    "noun": 30,
    "modifier": 30,
    "short": 100,
    "long": 3000,
    "uom": 10,
}


def make_short(long: str, limit: int = 100) -> str:
    """Potong PO Text rapi (di koma/spasi) agar muat batas Short Description."""
    if len(long) <= limit:
        return long

    cut = long[:limit]
    idx = max(cut.rfind(","), cut.rfind(" "))
    return (cut[:idx] if idx > 40 else cut).rstrip(" ,;")


def parse_uom_map(text: str) -> dict:
    mapping = {}
    for line in (text or "").splitlines():
        if "=" in line:
            source, target = line.split("=", 1)
            if source.strip():
                mapping[source.strip().upper()] = target.strip().upper()
    return mapping


def validate_template1(
    df: pd.DataFrame,
    valid_pairs: set
) -> pd.DataFrame:
    """df kolom: code, noun, modifier, short, long, uom. Return tabel masalah."""
    issues = []

    required = ["code", "noun", "modifier", "short", "long", "uom"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            "Kolom validasi Template 1 tidak ditemukan: "
            + ", ".join(missing)
        )

    def add(index, field, message):
        issues.append({
            "baris": index + 1,
            "kode": df.iloc[index]["code"],
            "kolom": field,
            "masalah": message,
        })

    duplicates = df["code"].duplicated(keep=False)

    for i in range(len(df)):
        row = df.iloc[i]

        if not str(row["code"]).strip():
            add(i, "code", "Kode material kosong")

        if duplicates.iloc[i]:
            add(i, "code", "Kode material duplikat")

        noun = str(row["noun"]).strip()
        modifier = str(row["modifier"]).strip()

        if not noun or not modifier:
            add(i, "noun/modifier", "Belum diklasifikasi")
        elif (noun, modifier) not in valid_pairs:
            add(
                i, "noun/modifier",
                f"Pasangan tidak ada di master: {noun} | {modifier}"
            )

        if not str(row["long"]).strip():
            add(i, "long", "Long Description kosong")

        if not str(row["uom"]).strip():
            add(i, "uom", "Uom kosong")

        for field, limit in LIMITS.items():
            length = len(str(row[field]))
            if length > limit:
                add(
                    i, field,
                    f"Melebihi {limit} karakter ({length})"
                )

    return pd.DataFrame(
        issues,
        columns=["baris", "kode", "kolom", "masalah"]
    )


FORMAT_ROWS = [
    [
        "INTEGER", "ALPHANUMERIC", "ALPHANUMERIC",
        "ALPHANUMERIC", "ALPHANUMERIC", "ALPHANUMERIC",
        "ALPHANUMERIC", "HEADER DATA", "DATA TYPE"
    ],
    [
        "No. of the material data ",
        "The Code/KEYMAP# of the material",
        "NOUN as per the CLASS master",
        "MODIFIER as per the CLASS master",
        "Short description of the material",
        "Long description of the material",
        "Uom of the material",
        None,
        "DESCRIPTION"
    ],
    ["2", "20", "30", "30", "100", "3000", "10", None, "SIZE"],
    [
        "No.", "PSC Material code", "Noun", "Modifier",
        "Material - Short Description",
        "Material - Long Description", "Uom", None, "HEADER"
    ],
]


def build_template1_bytes(df: pd.DataFrame) -> bytes:
    """Buat file Template 1. Pemotongan karakter harus divalidasi terpisah."""
    required = [
        "code", "noun", "modifier", "short", "long", "uom"
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            "Kolom untuk membuat Template 1 tidak ditemukan: "
            + ", ".join(missing)
        )

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Source_Data_Upload"

    worksheet.append([
        "No.", "PSC Material Code", "Noun", "Modifier",
        "Material - Short Description",
        "Material - Long Description", "Uom"
    ])

    for number, row in enumerate(
        df[required].itertuples(index=False, name="MaterialRow"),
        start=1
    ):
        worksheet.append([
            number,
            row.code,
            row.noun,
            row.modifier,
            row.short,
            row.long,
            row.uom,
        ])

    format_sheet = workbook.create_sheet(
        "Source_Data_Upload_Format"
    )
    for row in FORMAT_ROWS:
        format_sheet.append(row)

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
