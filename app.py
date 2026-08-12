"""
Aplikasi Pencatat Temuan Patroli Tower
========================================
Tempel teks laporan patroli (format grup WA), aplikasi akan otomatis
memecahnya per tower & per temuan, lalu menyimpannya ke file Excel.

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""

import re
import os
from datetime import date

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
DATA_FILE = "data_temuan.xlsx"
SHEET_NAME = "Temuan"
COLUMNS = [
    "ID",
    "Tanggal Patroli",
    "Tower",
    "PTD",
    "Lantai",
    "Temuan",
    "Kategori",
    "PIC",
    "Status",
    "Keterangan Progress",
    "Tanggal Update",
    "Waktu Input",
]
STATUS_OPTIONS = ["Baru", "Dalam Proses", "Selesai"]
PROGRESS_OPTIONS = ["On Progress", "Selesai"]

st.set_page_config(page_title="Pencatat Temuan Patroli", page_icon="🧱", layout="wide")


# ---------------------------------------------------------------------------
# Fungsi bantu
# ---------------------------------------------------------------------------
def load_data() -> pd.DataFrame:
    """Muat data yang sudah tersimpan, atau buat DataFrame kosong jika belum ada."""
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            # Data lama mungkin belum punya ID -> isi otomatis
            if df["ID"].isna().any() or (df["ID"] == "").any():
                df["ID"] = range(1, len(df) + 1)
            df["ID"] = df["ID"].astype(int)
            # Pastikan kolom teks tetap bertipe string (bukan NaN/float) setelah round-trip Excel
            kolom_teks = [c for c in COLUMNS if c != "ID"]
            for col in kolom_teks:
                df[col] = df[col].fillna("").astype(str)
                df.loc[df[col] == "nan", col] = ""
            return df[COLUMNS]
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNS)


def next_id(df: pd.DataFrame) -> int:
    if df.empty or "ID" not in df.columns or df["ID"].isna().all():
        return 1
    return int(df["ID"].max()) + 1


def save_data(df: pd.DataFrame) -> None:
    df.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False)


def tebak_kategori(teks: str) -> str:
    """Menebak kategori temuan secara sederhana berdasarkan kata kunci."""
    teks_lower = teks.lower()
    aturan = [
        ("Tembok Retak", ["retak"]),
        ("Gompal/Terkupas", ["gompal", "terkupas", "kupas"]),
        ("Berlubang", ["berlubang", "bolong"]),
        ("Berjamur", ["jamur", "berjamur"]),
        ("Panel/Elektrikal", ["panel", "listrik", "mcb"]),
        ("Keamanan/Akses", ["tidak terkunci", "kunci"]),
    ]
    for kategori, kata_kunci in aturan:
        if any(k in teks_lower for k in kata_kunci):
            return kategori
    return "Lainnya"


def parse_laporan(teks: str) -> list[dict]:
    """
    Pecah teks laporan mentah menjadi list temuan.
    Mendukung variasi heading seperti:
      "Temuan patroli tower Clifford sbb :"
      "Temuan Patroli Tower Belmont Sbb:"
      "Temuan patroli publik Area"           (tanpa nama tower / tanpa "sbb")
    dan penomoran seperti "1)." , "1.)", "2.)" dsb.
    Baris "@Nama Orang" di akhir blok dianggap sebagai PIC / penanggung jawab.
    """
    records = []

    pola_heading = re.compile(
        r"Temuan\s*Patroli\s*(?:Tower\s+(?P<tower>[A-Za-z]+)|(?P<area>Publik\s*Area|Public\s*Area))"
        r"\s*(?:sbb\s*:?)?",
        re.IGNORECASE,
    )
    matches = list(pola_heading.finditer(teks))

    if not matches:
        return records

    for i, m in enumerate(matches):
        if m.group("area"):
            tower = "Publik Area"
        else:
            tower = m.group("tower").strip().title()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(teks)
        blok = teks[start:end]

        pic_match = re.search(r"@(.+)", blok)
        pic = pic_match.group(1).strip() if pic_match else ""
        blok_tanpa_pic = re.sub(r"@.+", "", blok)

        for baris in blok_tanpa_pic.split("\n"):
            baris = baris.strip()
            if not baris:
                continue
            item_match = re.match(r"^\d+\s*[.)]+\s*(.+)", baris)
            if not item_match:
                continue
            temuan = re.sub(r"\s+", " ", item_match.group(1).strip())

            lantai_m = re.search(r"(?:Lt\.?|Lantai)\s*(\d+)", temuan, re.IGNORECASE)
            ptd_m = re.search(r"PTD\s*(\d+)", temuan, re.IGNORECASE)

            records.append(
                {
                    "Tower": tower,
                    "PTD": ptd_m.group(1) if ptd_m else "",
                    "Lantai": lantai_m.group(1) if lantai_m else "",
                    "Temuan": temuan,
                    "Kategori": tebak_kategori(temuan),
                    "PIC": pic,
                    "Status": "Baru",
                }
            )

    return records


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🧱 Pencatat Temuan Patroli Tower")
st.caption(
    "Tempel teks laporan patroli dari grup WA, sistem akan otomatis memecahnya "
    "menjadi baris-baris temuan siap simpan ke spreadsheet."
)

if "hasil_parse" not in st.session_state:
    st.session_state.hasil_parse = None

tab_input, tab_riwayat = st.tabs(["📋 Input Laporan Baru", "📚 Riwayat Temuan"])

# ----------------------------- TAB INPUT -----------------------------------
with tab_input:
    col_a, col_b = st.columns([2, 1])

    with col_a:
        teks_laporan = st.text_area(
            "Tempel teks laporan di sini",
            height=350,
            placeholder=(
                "Temuan patroli tower Clifford sbb :\n"
                "1). Termonitor Tembok Retak Di PTD 2 Lt 11\n"
                "2.) Termonitor Tembok Terkupas Di Lt 01 PTD 1\n"
                "..."
            ),
        )

    with col_b:
        tanggal_patroli = st.date_input("Tanggal Patroli", value=date.today())
        st.write("")
        parse_clicked = st.button("🔍 Parse Laporan", type="primary", use_container_width=True)
        st.info(
            "Setelah di-parse, kamu masih bisa mengedit tiap baris (kategori, "
            "status, PIC) sebelum disimpan ke spreadsheet.",
            icon="✏️",
        )

    if parse_clicked:
        if not teks_laporan.strip():
            st.warning("Teks laporan masih kosong.")
        else:
            hasil = parse_laporan(teks_laporan)
            if not hasil:
                st.error(
                    "Tidak ada temuan yang terdeteksi. Pastikan format masih memuat "
                    "'Temuan Patroli Tower ... sbb :' dan penomoran seperti '1).'."
                )
            else:
                for r in hasil:
                    r["Tanggal Patroli"] = tanggal_patroli
                st.session_state.hasil_parse = pd.DataFrame(hasil)
                st.success(f"{len(hasil)} temuan berhasil terdeteksi dari {teks_laporan.count('Temuan')} laporan. Silakan cek & edit di bawah sebelum menyimpan.")

    if st.session_state.hasil_parse is not None:
        st.subheader("Pratinjau & Edit Sebelum Disimpan")
        edited_df = st.data_editor(
            st.session_state.hasil_parse,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status", options=["Baru", "Dalam Proses", "Selesai"]
                ),
                "Kategori": st.column_config.SelectboxColumn(
                    "Kategori",
                    options=[
                        "Tembok Retak",
                        "Gompal/Terkupas",
                        "Berlubang",
                        "Berjamur",
                        "Panel/Elektrikal",
                        "Keamanan/Akses",
                        "Lainnya",
                    ],
                ),
                "Tanggal Patroli": st.column_config.DateColumn("Tanggal Patroli"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="editor_temuan",
        )

        if st.button("💾 Simpan ke Spreadsheet", type="primary"):
            df_lama = load_data()
            df_baru = edited_df.copy()
            mulai_id = next_id(df_lama)
            df_baru["ID"] = range(mulai_id, mulai_id + len(df_baru))
            df_baru["Waktu Input"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            df_baru["Keterangan Progress"] = ""
            df_baru["Tanggal Update"] = ""
            for col in COLUMNS:
                if col not in df_baru.columns:
                    df_baru[col] = ""
            df_final = pd.concat([df_lama, df_baru[COLUMNS]], ignore_index=True)
            save_data(df_final)
            st.session_state.hasil_parse = None
            st.success(f"Tersimpan! Total data di spreadsheet sekarang: {len(df_final)} baris.")
            st.rerun()

# ----------------------------- TAB RIWAYAT ----------------------------------
with tab_riwayat:
    df_all = load_data()
    if df_all.empty:
        st.info("Belum ada data tersimpan.")
    else:
        jml_baru = int((df_all["Status"] == "Baru").sum())
        jml_proses = int((df_all["Status"] == "Dalam Proses").sum())
        jml_selesai = int((df_all["Status"] == "Selesai").sum())

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Temuan", len(df_all))
        m2.metric("🆕 Baru", jml_baru)
        m3.metric("🔧 Dalam Proses", jml_proses)
        m4.metric("✅ Selesai", jml_selesai)

        st.divider()

        col1, col2, col3 = st.columns(3)
        with col1:
            filter_tower = st.multiselect("Filter Tower", sorted(df_all["Tower"].dropna().unique()))
        with col2:
            filter_status = st.multiselect("Filter Status", STATUS_OPTIONS)
        with col3:
            cari = st.text_input("Cari kata kunci di kolom Temuan")

        df_filtered = df_all.copy()
        if filter_tower:
            df_filtered = df_filtered[df_filtered["Tower"].isin(filter_tower)]
        if filter_status:
            df_filtered = df_filtered[df_filtered["Status"].isin(filter_status)]
        if cari:
            df_filtered = df_filtered[df_filtered["Temuan"].str.contains(cari, case=False, na=False)]

        st.subheader("🔄 Update Progres / Tandai Selesai")
        st.caption(
            "Ubah kolom **Status**, **Keterangan Progress**, dan **PIC** langsung di tabel, "
            "lalu klik tombol simpan di bawah. Kolom lain terkunci (tidak bisa diubah)."
        )

        kolom_terkunci = [c for c in COLUMNS if c not in ("Status", "Keterangan Progress", "PIC")]

        edited_riwayat = st.data_editor(
            df_filtered,
            column_config={
                "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS),
                "Keterangan Progress": st.column_config.SelectboxColumn(
                    "Keterangan Progress", options=PROGRESS_OPTIONS
                ),
                "PIC": st.column_config.TextColumn("PIC", help="Nama penanggung jawab"),
                "ID": st.column_config.NumberColumn("ID", disabled=True),
            },
            disabled=kolom_terkunci,
            hide_index=True,
            use_container_width=True,
            key="editor_riwayat",
        )

        if st.button("💾 Simpan Perubahan Status/Progres", type="primary"):
            df_updated = df_all.set_index("ID")
            edited_indexed = edited_riwayat.set_index("ID")
            hari_ini = date.today().strftime("%Y-%m-%d")
            jumlah_berubah = 0

            for temuan_id, baris_baru in edited_indexed.iterrows():
                baris_lama = df_updated.loc[temuan_id]
                status_berubah = baris_lama["Status"] != baris_baru["Status"]
                catatan_berubah = baris_lama["Keterangan Progress"] != baris_baru["Keterangan Progress"]
                pic_berubah = baris_lama["PIC"] != baris_baru["PIC"]
                if status_berubah or catatan_berubah or pic_berubah:
                    df_updated.loc[temuan_id, "Status"] = baris_baru["Status"]
                    df_updated.loc[temuan_id, "Keterangan Progress"] = baris_baru["Keterangan Progress"]
                    df_updated.loc[temuan_id, "PIC"] = baris_baru["PIC"]
                    df_updated.loc[temuan_id, "Tanggal Update"] = hari_ini
                    jumlah_berubah += 1

            df_updated = df_updated.reset_index()[COLUMNS]
            save_data(df_updated)
            st.success(f"{jumlah_berubah} baris berhasil diperbarui.")
            st.rerun()

        st.caption(f"Menampilkan {len(df_filtered)} dari {len(df_all)} total temuan.")

        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "rb") as f:
                st.download_button(
                    "⬇️ Download Spreadsheet (Excel)",
                    data=f,
                    file_name=DATA_FILE,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
