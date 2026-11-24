"""Jalankan:  python test_core.py
Unit test selalu jalan. Uji data nyata jalan jika file ada di folder ini atau di SKK_DATA_DIR.
"""
import io
import os

import openpyxl
import pandas as pd

import core

# ---------------------------------------------------------------- unit test
def fake_master():
    pairs = [
        ("CROSS-OVER", "--"), ("PUP JOINT", "CASING"), ("PUP JOINT", "TUBING"),
        ("GASKET", "SPIRAL WOUND"), ("GASKET", "RING JOINT"), ("FUSE", "--"),
        ("SPARE PARTS", "--"), ("ACCESSORIES", "DRILLING"),
    ]
    return pd.DataFrame(pairs, columns=["NOUN", "MODIFIER"])


def test_clean_text():
    assert core.clean_text('5½” CASING') == '5-1/2" CASING'
    assert core.clean_text("ELBOW 90Â°") == "ELBOW 90°"
    assert core.clean_text("A   B\t\nC") == "A B C"
    assert core.clean_text("5-1/2GÇ¥ TBG") == '5-1/2" TBG'
    assert core.clean_text("a\xa0b") == "a b"
    assert core.clean_text("PUMP®", ascii_symbols=True) == "PUMP"
    assert core.clean_text("PUMP®") == "PUMP®"           # default: simbol tidak diubah
    assert core.clean_text('9.5/8"') == '9.5/8"'         # angka tidak pernah diubah
    assert core.clean_text(None) == ""


def test_match_text():
    assert core.match_text("pupjoint 5") == "PUP JOINT 5"
    assert core.match_text("X-OVER") == "CROSSOVER" == core.match_text("cross over")
    assert "TUBING" in core.match_text("TBG PUP")


def test_rules_and_pipeline():
    ctx = core.build_context(fake_master())
    assert core.rule_classify("X-OVER,4-1/2", "X-OVER 4-1/2 X 5", ctx)[:2] == ("CROSS-OVER", "--")
    assert core.rule_classify("PUPJOINT,5-1/2", "PUP JOINT CASING 5-1/2", ctx)[:2] == ("PUP JOINT", "CASING")
    assert core.rule_classify("TBG PUP,2-7/8", "TBG PUP 2-7/8 L80", ctx)[:2] == ("PUP JOINT", "TUBING")
    assert core.rule_classify("GASKET,SPIRAL WOUND,1/2", "GASKET SPIRAL WOUND 1/2", ctx)[:2] == ("GASKET", "SPIRAL WOUND")
    assert core.rule_classify("FUSE,10A", "FUSE 10A", ctx) is None     # noun satu-modifier -> bukan aturan

    rows = [
        {"code": "1", "short": "X-OVER", "long": "X-OVER 4-1/2"},
        {"code": "2", "short": "SPIROLIZER", "long": "SPIROLIZER 9-5/8"},
        {"code": "3", "short": "ABC", "long": "ABC"},
        {"code": "4", "short": "XYZ", "long": "XYZ"},
    ]
    # tanpa AI: yang tidak tertangkap aturan -> REVIEW dan kosong
    out = core.classify_rows(rows, ctx)
    assert out[0]["source"] == "ATURAN" and out[1]["source"] == "REVIEW" and out[1]["noun"] == ""

    # dengan AI palsu: valid diterima, tidak valid ditolak -> REVIEW, satu batch gagal tidak mematikan semua
    def fake_ai(payload):
        d = {}
        for p in payload:
            d[p["row_index"]] = ("ACCESSORIES", "DRILLING", 85, "tes") if p["row_index"] == 1 else ("NGAWUR", "X", 90, "tes")
        return d
    out = core.classify_rows(rows, ctx, ai_batch=fake_ai)
    assert out[1]["source"] == "AI" and (out[1]["noun"], out[1]["modifier"]) == ("ACCESSORIES", "DRILLING")
    assert out[2]["source"] == "REVIEW" and out[2]["noun"] == ""

    def boom(payload):
        raise RuntimeError("429")
    out = core.classify_rows(rows, ctx, ai_batch=boom)
    assert out[0]["source"] == "ATURAN" and out[1]["source"] == "REVIEW" and "429" in out[1]["note"]


def test_label_lookup_and_examples():
    lab = pd.DataFrame({
        "PSC Material Code": ["10"], "Noun": ["ACCESSORIES"], "Modifier": ["DRILLING"],
        "Material - Long Description": ["SPIROLIZER 9-5/8 HALLIBURTON"],
    })
    ctx = core.build_context(fake_master(), lab)
    out = core.classify_rows([{"code": "10", "short": "x", "long": "y"}], ctx)
    assert out[0]["source"] == "CONTOH" and out[0]["noun"] == "ACCESSORIES"
    ex = core.similar_examples("SPIROLIZER 13-3/8", ctx)
    assert ex and ex[0][1] == "ACCESSORIES"


def test_validate_and_build():
    valid = {("CROSS-OVER", "--")}
    df = pd.DataFrame([
        {"code": "1", "noun": "CROSS-OVER", "modifier": "--", "short": "S", "long": "L", "uom": "EA"},
        {"code": "1", "noun": "", "modifier": "", "short": "S" * 101, "long": "L" * 3001, "uom": ""},
        {"code": "3", "noun": "NGAWUR", "modifier": "X", "short": "S", "long": "L", "uom": "EA"},
    ])
    iss = core.validate_template1(df, valid)
    msgs = " | ".join(iss["masalah"])
    for k in ("duplikat", "Belum diklasifikasi", "tidak ada di master", "Melebihi 100", "Melebihi 3000", "Uom kosong"):
        assert k in msgs, k
    wb = openpyxl.load_workbook(io.BytesIO(core.build_template1_bytes(df)))
    assert wb.sheetnames == ["Source_Data_Upload", "Source_Data_Upload_Format"]
    ws = wb["Source_Data_Upload"]
    assert [c.value for c in ws[1]][:3] == ["No.", "PSC Material Code", "Noun"]
    assert ws.max_row == 4 and len(ws.cell(3, 6).value) == 3001   # tidak dipotong diam-diam


def test_helpers():
    assert core.parse_uom_map("pcs=pc\nBX = BOX\nngawur") == {"PCS": "PC", "BX": "BOX"}
    s = core.make_short("PUP JOINT, 5-1/2\", 17 PPF, L80, VAM TOP, " + "X" * 20 + ", " + "Y" * 60)
    assert len(s) <= 100 and not s.endswith(",")
    assert core.detect_columns(["Material", "Material Description", "PO Text", "Base Unit of Measure"]) | {} \
        and core.detect_columns(["Material", "Material Description", "PO Text", "Base Unit of Measure"])["long"] == "PO Text"


def test_is_fatal_error():
    class E(Exception):
        def __init__(self, m, code=None):
            super().__init__(m)
            self.code = code
    assert core.is_fatal_error(E("x", 401))
    assert core.is_fatal_error(E("401 UNAUTHENTICATED ACCESS_TOKEN_TYPE_UNSUPPORTED"))
    assert core.is_fatal_error(E("404 NOT_FOUND: models/gemini-2.5-flash is no longer available"))
    assert core.is_fatal_error(E("429 RESOURCE_EXHAUSTED GenerateRequestsPerDayPerProjectPerModel-FreeTier"))
    assert core.is_fatal_error(E("quota exceeded, limit: 0"))
    assert not core.is_fatal_error(E("503 UNAVAILABLE high demand", 503))
    assert not core.is_fatal_error(E("429 RESOURCE_EXHAUSTED requests per minute", 429))


def test_fatal_error_stops_ai():
    ctx = core.build_context(fake_master())
    rows = [{"code": str(i), "short": f"ABC{i}", "long": f"ABC{i}"} for i in range(7)]
    calls = {"n": 0}

    def fatal(payload):
        calls["n"] += 1
        raise core.AIFatalError("401 ACCESS_TOKEN_TYPE_UNSUPPORTED")

    out = core.classify_rows(rows, ctx, ai_batch=fatal, batch_size=2)
    assert calls["n"] == 1                       # hanya 1 panggilan, bukan 4
    assert all(o["source"] == "REVIEW" for o in out)
    assert out[0]["note"].startswith("Gagal AI") and out[6]["note"].startswith("AI dihentikan")


# ------------------------------------------------------------ uji data nyata
def find(pattern):
    for d in (os.environ.get("SKK_DATA_DIR", ""), os.path.dirname(os.path.abspath(__file__))):
        if d and os.path.isdir(d):
            for f in sorted(os.listdir(d)):
                if pattern(f.lower()):
                    return os.path.join(d, f)
    return None


def real_data_check():
    ms = find(lambda f: f.startswith("material_standard") and f.endswith(".xlsx"))
    data = find(lambda f: f in ("all.xlsx", "all.csv"))
    ex = find(lambda f: "extracted_data" in f and f.endswith(".xlsx"))
    if not (ms and data):
        print("(uji data nyata dilewati: file master/data tidak ditemukan)")
        return
    df_nm, _ = core.load_master(pd.read_excel(ms, dtype=str))
    raw = pd.read_excel(data, dtype=str).fillna("") if data.endswith("xlsx") else core.read_any(data, open(data, "rb").read())
    cols = core.detect_columns(list(raw.columns))
    raw = raw[raw[cols["code"]].str.strip() != ""]
    rows = [{"code": r[cols["code"]].strip(), "short": core.clean_text(r[cols["short"]]),
             "long": core.clean_text(r[cols["long"]]), "uom": r[cols["uom"]].strip().upper()}
            for _, r in raw.iterrows()]
    print(f"Data: {len(rows)} baris | kolom terdeteksi: {cols}")

    # a) akurasi aturan murni pada material berlabel (tanpa label lookup)
    if ex:
        lab = pd.read_excel(ex, dtype=str).fillna("")
        truth = lab.groupby("PSC Material Code").first()[["Noun", "Modifier"]]
        ctx0 = core.build_context(df_nm)  # tanpa label -> menguji aturan saja
        hit = ok = n = 0
        for r in rows:
            if r["code"] in truth.index:
                n += 1
                p = core.rule_classify(r["short"], r["long"], ctx0)
                if p:
                    hit += 1
                    ok += (p[0], p[1]) == tuple(truth.loc[r["code"]])
        print(f"Aturan pada {n} material berlabel: ter-cover {hit} ({hit/n:.0%}), benar {ok} ({ok/max(hit,1):.0%})")
        ctx = core.build_context(df_nm, lab)
    else:
        ctx = core.build_context(df_nm)

    # b) pipeline penuh tanpa AI
    out = core.classify_rows(rows, ctx)
    from collections import Counter
    print("Sumber hasil (tanpa AI):", dict(Counter(o["source"] for o in out)))

    # c) validasi + pembuatan file
    t1 = pd.DataFrame([{"code": r["code"], "noun": o["noun"], "modifier": o["modifier"],
                        "short": r["short"], "long": r["long"], "uom": r["uom"]} for r, o in zip(rows, out)])
    iss = core.validate_template1(t1, ctx["valid"])
    print("Masalah validasi:", iss["masalah"].str.replace(r"\(.*\)|:.*", "", regex=True).value_counts().to_dict())
    b = core.build_template1_bytes(t1)
    print(f"Template 1 terbuat: {len(b)/1024:.0f} KB")
    changed = sum(core.clean_text(a) != a for a in raw[cols["long"]])
    print(f"Baris PO Text yang berubah oleh pembersihan: {changed}")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK  ", name)
    real_data_check()
