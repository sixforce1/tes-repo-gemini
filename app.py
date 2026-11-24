"""app.py - Streamlit: data SAP -> bersihkan -> klasifikasi Noun|Modifier -> Template 1.

Jalankan:
    streamlit run app.py

Tahap 3 (Template 2 / ekstraksi atribut) sengaja belum ada:
dibuat setelah Template 1 diterima e-cataloging.
"""
from __future__ import annotations

import io
import os

import pandas as pd
import streamlit as st

import core


st.set_page_config(
    page_title="SKK Migas - Template 1",
    layout="wide",
)

BASE = os.path.dirname(os.path.abspath(__file__))
EDITABLE = ["Kelas", "Short (final)", "Uom"]


# ------------------------------------------------------------------
# Helper UI
# ------------------------------------------------------------------
def find_local(pred):
    for filename in sorted(os.listdir(BASE)):
        if pred(filename.lower()):
            return os.path.join(BASE, filename)
    return None


def secret(name, default=""):
    try:
        value = st.secrets.get(name, "")
    except Exception:
        value = ""
    return value or os.getenv(name, default)


@st.cache_data(show_spinner=False)
def cached_master(raw: bytes):
    return core.load_master(
        pd.read_excel(io.BytesIO(raw), dtype=str)
    )


@st.cache_data(show_spinner=False)
def cached_labeled(raw: bytes):
    return pd.read_excel(
        io.BytesIO(raw), dtype=str
    ).fillna("")


def apply_edits(base: pd.DataFrame, edits: dict) -> pd.DataFrame:
    """Terapkan edit manual di atas hasil dasar; base tidak diubah."""
    current = base.copy()

    for index, changes in edits.items():
        for column, value in changes.items():
            current.at[index, column] = value

        if "Kelas" in changes:
            current.at[index, "Sumber"] = "MANUAL"
            current.at[index, "Keyakinan"] = 100

    return current


def to_template1(current: pd.DataFrame) -> pd.DataFrame:
    parts = current["Kelas"].fillna("").astype(str).str.split(
        " | ", n=1, regex=False, expand=True
    )

    if parts.shape[1] < 2:
        parts[1] = ""

    return pd.DataFrame({
        "code": current["Kode"],
        "noun": parts[0].fillna(""),
        "modifier": parts[1].fillna(""),
        "short": current["Short (final)"],
        "long": current["Long"],
        "uom": current["Uom"],
    })


# ------------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------------
with st.sidebar:
    st.title("Pengaturan")

    st.subheader("Referensi")

    master_path = find_local(
        lambda filename:
            filename.startswith("material_standard")
            and filename.endswith(".xlsx")
    )

    master_raw = (
        open(master_path, "rb").read()
        if master_path else None
    )

    if master_raw is None:
        uploaded_master = st.file_uploader(
            "Material Standard (.xlsx)",
            type=["xlsx"],
            key="ms",
        )
        master_raw = (
            uploaded_master.getvalue()
            if uploaded_master else None
        )

    example_path = find_local(
        lambda filename:
            "extracted_data" in filename
            and filename.endswith(".xlsx")
    )

    example_raw = (
        open(example_path, "rb").read()
        if example_path else None
    )

    if example_raw is None:
        uploaded_example = st.file_uploader(
            "Contoh berlabel (Template 2 terisi), opsional",
            type=["xlsx"],
            key="ex",
        )
        example_raw = (
            uploaded_example.getvalue()
            if uploaded_example else None
        )

    st.subheader("Pembersihan & output")

    short_mode = st.radio(
        "Short Description",
        [
            "Short SAP apa adanya",
            "Dari PO Text (dipotong rapi, maks 100)",
        ],
    )

    ascii_symbols = st.checkbox(
        "Ganti simbol ® ™ • ± ³ ke ASCII",
        value=False,
        help=(
            "Nyalakan hanya jika e-cataloging menolak "
            "simbol tersebut."
        ),
    )

    uom_text = st.text_area(
        "Pemetaan Uom (SAP=SKK, satu per baris)",
        placeholder="PCS=PC\nBX=BOX",
        height=80,
    )

    st.subheader("AI (Gemini)")

    api_key = secret("GEMINI_API_KEY")

    use_ai = st.checkbox(
        "Pakai AI untuk yang tidak tertangkap aturan",
        value=bool(api_key),
        disabled=not api_key,
    )

    if not api_key:
        st.caption(
            "GEMINI_API_KEY belum diatur: baris sisa akan "
            "berstatus REVIEW untuk diisi manual."
        )

    model = st.text_input(
        "Model",
        value=secret("GEMINI_MODEL", "gemini-3.5-flash"),
        help=(
            "Isi dengan nama model Gemini yang tersedia "
            "untuk API key kamu."
        ),
    )

    # Default 5 agar beban per permintaan lebih kecil.
    batch_size = st.slider(
        "Item per panggilan AI",
        min_value=1,
        max_value=10,
        value=5,
        help=(
            "Jika sering mendapat error 503/429, gunakan "
            "batch kecil."
        ),
    )

    limit_rows = st.number_input(
        "Uji: proses N baris pertama saja (0 = semua)",
        min_value=0,
        value=0,
        step=50,
    )


# ------------------------------------------------------------------
# Load master
# ------------------------------------------------------------------
if master_raw is None:
    st.info(
        "Letakkan file Material Standard di folder aplikasi "
        "atau unggah lewat sidebar."
    )
    st.stop()

try:
    df_nm, df_ms = cached_master(master_raw)
    labeled = (
        cached_labeled(example_raw)
        if example_raw else None
    )
    ctx = core.build_context(df_nm, labeled)

except Exception as error:
    st.error(f"Gagal membaca file referensi: {error}")
    st.stop()

with st.sidebar:
    st.success(
        f"Master: {len(df_nm):,} noun|modifier\n\n"
        f"Contoh berlabel: {len(ctx['labels']):,} material"
    )


# ------------------------------------------------------------------
# Input data SAP
# ------------------------------------------------------------------
st.title("Template 1 - Source_Data_Upload")

uploaded_sap = st.file_uploader(
    "Data material mentah dari SAP (.xlsx / .csv)",
    type=["xlsx", "csv"],
)

if uploaded_sap is None:
    st.info(
        "Unggah data SAP (kolom: Material, Material Description, "
        "PO Text, Base Unit of Measure)."
    )
    st.stop()

try:
    raw_df = core.read_any(
        uploaded_sap.name,
        uploaded_sap.getvalue(),
    )
except Exception as error:
    st.error(f"Gagal membaca file SAP: {error}")
    st.stop()

cols = list(raw_df.columns)

if not cols:
    st.error("File SAP tidak memiliki kolom/header yang dapat dibaca.")
    st.stop()

guess = core.detect_columns(cols)

def column_index(name):
    return cols.index(name) if name in cols else 0


c1, c2, c3, c4 = st.columns(4)

col_code = c1.selectbox(
    "Kolom kode material",
    cols,
    index=column_index(guess["code"]),
)

col_short = c2.selectbox(
    "Kolom short (SAP)",
    cols,
    index=column_index(guess["short"]),
)

col_long = c3.selectbox(
    "Kolom long (PO Text)",
    cols,
    index=column_index(guess["long"]),
)

col_uom = c4.selectbox(
    "Kolom Uom",
    cols,
    index=column_index(guess["uom"]),
)

st.caption(f"{len(raw_df):,} baris terbaca.")
st.dataframe(raw_df.head(5), use_container_width=True)


# ------------------------------------------------------------------
# Proses klasifikasi
# ------------------------------------------------------------------
if st.button("Bersihkan & Klasifikasi", type="primary"):
    data = raw_df[
        raw_df[col_code].astype(str).str.strip() != ""
    ].copy()

    if limit_rows > 0:
        data = data.head(int(limit_rows))

    if data.empty:
        st.warning("Tidak ada baris dengan kode material yang terisi.")
        st.stop()

    uom_map = core.parse_uom_map(uom_text)

    rows = []

    for _, row in data.iterrows():
        uom = str(row[col_uom]).strip().upper()

        rows.append({
            "code": str(row[col_code]).strip(),
            "short": core.clean_text(
                row[col_short], ascii_symbols
            ),
            "long": core.clean_text(
                row[col_long], ascii_symbols
            ),
            "uom": uom_map.get(uom, uom),
        })

    ai_fn = None

    if use_ai and api_key:
        try:
            ai_fn = core.make_gemini_batch_fn(
                api_key=api_key,
                model=model.strip(),
                ctx=ctx,
                retries=5,
                base_delay=2,
                max_delay=30,
            )
        except Exception as error:
            st.error(f"AI tidak dapat diaktifkan: {error}")

    progress_bar = st.progress(
        0.0,
        text="Menerapkan label contoh & aturan...",
    )

    def on_progress(done, total):
        if total > 0:
            progress_bar.progress(
                min(done / total, 1.0),
                text=(
                    f"AI mengklasifikasi "
                    f"{done:,} / {total:,} item"
                ),
            )

    try:
        output = core.classify_rows(
            rows,
            ctx,
            ai_batch=ai_fn,
            batch_size=batch_size,
            progress=on_progress,
        )
    except Exception as error:
        progress_bar.empty()
        st.error(f"Proses klasifikasi gagal: {error}")
        st.stop()

    progress_bar.empty()

    ai_problems = [
        item["note"] for item in output
        if item["note"].startswith(("Gagal AI", "AI dihentikan"))
    ]
    st.session_state["ai_error"] = (
        {"count": len(ai_problems), "first": ai_problems[0]}
        if ai_problems else None
    )

    base = pd.DataFrame({
        "Kode": [row["code"] for row in rows],
        "Kelas": [
            (
                f"{item['noun']} | {item['modifier']}"
                if item["noun"] else ""
            )
            for item in output
        ],
        "Sumber": [item["source"] for item in output],
        "Keyakinan": [item["confidence"] for item in output],
        "Catatan": [item["note"] for item in output],
        "Short SAP": [row["short"] for row in rows],
        "Short (final)": [
            (
                row["short"]
                if short_mode.startswith("Short SAP")
                else core.make_short(row["long"] or row["short"])
            )
            for row in rows
        ],
        "Long": [row["long"] for row in rows],
        "Uom": [row["uom"] for row in rows],
    })

    st.session_state["base"] = base
    st.session_state["edits"] = {}
    st.session_state["run_id"] = (
        st.session_state.get("run_id", 0) + 1
    )


if "base" not in st.session_state:
    st.stop()


# ------------------------------------------------------------------
# Review hasil
# ------------------------------------------------------------------
base = st.session_state["base"]
edits = st.session_state["edits"]
current = apply_edits(base, edits)

st.subheader("Review hasil")

ai_error = st.session_state.get("ai_error")
if ai_error:
    st.error(
        f"AI gagal untuk {ai_error['count']:,} baris. "
        f"Pesan pertama: {ai_error['first']}"
    )

summary = st.empty()

only_review = st.checkbox(
    "Hanya tampilkan yang perlu review "
    "(kelas kosong atau keyakinan < 70)",
    value=True,
)

# Keanggotaan filter dihitung dari base (stabil). Jika dihitung dari current,
# baris yang baru diedit langsung hilang dari tabel dan posisi baris bergeser,
# sehingga data_editor bisa menerapkan edit ke baris yang salah.
if only_review:
    mask = (
        (base["Kelas"].fillna("") == "")
        | (base["Keyakinan"] < 70)
    )
else:
    mask = pd.Series(True, index=base.index)

view = current.loc[mask].copy()

pairs = [""] + [
    f"{noun} | {modifier}"
    for noun, modifier in zip(
        df_nm["NOUN"],
        df_nm["MODIFIER"],
    )
]

edited = st.data_editor(
    view,
    key=f"ed_{st.session_state['run_id']}_{only_review}",
    use_container_width=True,
    hide_index=True,
    height=520,
    disabled=[
        column for column in view.columns
        if column not in EDITABLE
    ],
    column_config={
        "Kelas": st.column_config.SelectboxColumn(
            "Kelas (Noun | Modifier)",
            options=pairs,
            width="medium",
        ),
        "Keyakinan": st.column_config.NumberColumn(
            format="%d"
        ),
    },
)

# Simpan perubahan user ke dictionary edit berdasarkan index asli.
for index in edited.index:
    for column in EDITABLE:
        new_value = edited.at[index, column]

        if (
            new_value is None
            or (
                isinstance(new_value, float)
                and pd.isna(new_value)
            )
        ):
            new_value = ""
        else:
            new_value = str(new_value)

        if new_value != str(view.at[index, column]):
            edits.setdefault(index, {})[column] = new_value

st.session_state["edits"] = edits
current = apply_edits(base, edits)


with summary.container():
    counts = current["Sumber"].value_counts().to_dict()

    metrics = st.columns(6)
    metrics[0].metric("Total", f"{len(current):,}")

    for index, source in enumerate(
        ["CONTOH", "ATURAN", "AI", "MANUAL", "REVIEW"],
        start=1,
    ):
        metrics[index].metric(
            source,
            f"{counts.get(source, 0):,}",
        )

    empty_count = int(
        (current["Kelas"].fillna("") == "").sum()
    )

    if empty_count:
        st.warning(
            f"{empty_count:,} baris belum punya Noun | Modifier. "
            "Isi lewat tabel di bawah, atau aktifkan AI di sidebar."
        )


# ------------------------------------------------------------------
# Validasi & unduh
# ------------------------------------------------------------------
st.subheader("Validasi & unduh")

template1 = to_template1(current)
issues = core.validate_template1(
    template1,
    ctx["valid"],
)

if issues.empty:
    st.success("Semua baris lolos validasi.")
else:
    st.error(
        f"{len(issues):,} masalah pada "
        f"{issues['baris'].nunique():,} baris."
    )

    issue_summary = (
        issues.groupby("masalah")
        .size()
        .rename("jumlah")
        .reset_index()
        .sort_values("jumlah", ascending=False)
    )

    st.dataframe(
        issue_summary,
        use_container_width=True,
    )

    with st.expander("Lihat semua masalah"):
        st.dataframe(
            issues,
            use_container_width=True,
        )


download1, download2 = st.columns(2)

download1.download_button(
    "Download Template 1 (Source_Data_Upload.xlsx)",
    data=core.build_template1_bytes(template1),
    file_name="Source_Data_Upload.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    type="primary",
)

download2.download_button(
    "Download laporan klasifikasi (audit)",
    data=current.to_csv(index=False).encode("utf-8-sig"),
    file_name="laporan_klasifikasi.csv",
    mime="text/csv",
)

st.caption(
    "Template 1 tidak memuat kolom Sumber/Keyakinan; "
    "laporan audit memuatnya."
)
