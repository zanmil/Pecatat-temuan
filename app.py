"""
Aplikasi Pencatat Temuan Patroli Tower
========================================
Tempel teks laporan patroli (format grup WA) — baik temuan bangunan (tembok
retak, dsb) maupun temuan CCTV mati — aplikasi otomatis memecahnya jadi
baris-baris data dan menyimpannya ke file CSV (data_temuan.csv).

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""

import re
import io
import os
from datetime import date

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
DATA_FILE = "data_temuan.csv"

COLUMNS = [
    "ID",
    "Tanggal Patroli",
    "Jenis Temuan",
    "Tower",
    "PTD",
    "Lantai",
    "DVR",
    "Channel",
    "Temuan",
    "Kategori",
    "PIC",
    "Status",
    "Keterangan Progress",
    "Tanggal Update",
    "Waktu Input",
]

JENIS_OPTIONS = ["Bangunan", "CCTV"]
STATUS_OPTIONS = ["Baru", "Dalam Proses", "Selesai"]
PROGRESS_OPTIONS = ["On Progress", "Selesai"]
KATEGORI_OPTIONS = [
    "Tembok Retak",
    "Gompal/Terkupas",
    "Berlubang",
    "Berjamur",
    "Panel/Elektrikal",
    "Keamanan/Akses",
    "CCTV Mati",
    "Lainnya",
]

st.set_page_config(page_title="Pencatat Temuan Patroli", page_icon="🧱", layout="wide")


# ---------------------------------------------------------------------------
# Fungsi Penyimpanan (CSV)
# ---------------------------------------------------------------------------
def load_data() -> pd.DataFrame:
    """Ambil semua data dari file CSV, atau DataFrame kosong kalau belum ada."""
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_csv(DATA_FILE, dtype=str)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df = df.fillna("")
            if (df["ID"] == "").any():
                df["ID"] = range(1, len(df) + 1)
            df["ID"] = df["ID"].astype(int)
            return df[COLUMNS]
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNS)


def next_id(df: pd.DataFrame) -> int:
    if df.empty:
        return 1
    return int(df["ID"].max()) + 1


def save_semua(df: pd.DataFrame) -> None:
    df.to_csv(DATA_FILE, index=False)


def insert_rows(df_baru: pd.DataFrame) -> None:
    """Tambahkan baris-baris temuan baru ke CSV (ID otomatis lanjut dari data lama)."""
    df_lama = load_data()
    mulai_id = next_id(df_lama)
    df_baru = df_baru.copy()
    df_baru["ID"] = range(mulai_id, mulai_id + len(df_baru))
    for col in COLUMNS:
        if col not in df_baru.columns:
            df_baru[col] = ""
    df_final = pd.concat([df_lama, df_baru[COLUMNS]], ignore_index=True)
    save_semua(df_final)


def export_ke_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Temuan", index=False)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Parser Laporan (Bangunan + CCTV)
# ---------------------------------------------------------------------------
def tebak_kategori(teks: str) -> str:
    """Menebak kategori temuan bangunan secara sederhana berdasarkan kata kunci."""
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
    Pecah teks laporan mentah menjadi list temuan. Mendukung dua jenis heading:

      1) Temuan bangunan:
         "Temuan patroli tower Clifford sbb :"      (per tower, numbering "1).")
         "Temuan patroli publik Area"                (tanpa nama tower/"sbb")

      2) Temuan CCTV mati:
         "Kamera CCTV yang mati 20 titik"            (bullet "* DVR X Ch Y (lokasi) [catatan]")

    Baris "@Nama Orang" di blok bangunan dianggap PIC. Tower/PTD/Lantai untuk
    item CCTV otomatis diekstrak dari teks lokasi dalam kurung (mis. "Twr A PTD2 Lt 11").
    """
    records: list[dict] = []

    pola_heading = re.compile(
        r"Temuan\s*Patroli\s*(?:Tower\s+(?P<tower>[A-Za-z]+)|(?P<area>Publik\s*Area|Public\s*Area))\s*(?:sbb\s*:?)?"
        r"|Kamera\s*CCTV\s*yang\s*mati\s*(?:\d+\s*titik)?",
        re.IGNORECASE,
    )
    matches = list(pola_heading.finditer(teks))
    if not matches:
        return records

    for i, m in enumerate(matches):
        heading_text = m.group(0)
        jenis = "CCTV" if re.search(r"cctv", heading_text, re.IGNORECASE) else "Bangunan"

        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(teks)
        blok = teks[start:end]

        if jenis == "Bangunan":
            tower = "Publik Area" if m.group("area") else m.group("tower").strip().title()

            pic_match = re.search(r"@(.+)", blok)
            pic = pic_match.group(1).strip() if pic_match else ""
            blok_bersih = re.sub(r"@.+", "", blok)

            for baris in blok_bersih.split("\n"):
                baris = baris.strip()
                if not baris:
                    continue
                item_match = re.match(r"^\d+\s*[.)]+\s*(.+)", baris)
                if not item_match:
                    continue
                temuan = re.sub(r"\s+", " ", item_match.group(1).strip())
                lantai_m = re.search(r"(?:Lt\.?|Lantai)\s*(\d+)", temuan, re.IGNORECASE)
                ptd_m = re.search(r"PTD\s*(\d+)", temuan, re.IGNORECASE)

                records.append({
                    "Jenis Temuan": "Bangunan",
                    "Tower": tower,
                    "PTD": ptd_m.group(1) if ptd_m else "",
                    "Lantai": lantai_m.group(1) if lantai_m else "",
                    "DVR": "",
                    "Channel": "",
                    "Temuan": temuan,
                    "Kategori": tebak_kategori(temuan),
                    "PIC": pic,
                    "Status": "Baru",
                })
        else:  # CCTV
            for baris in blok.split("\n"):
                baris = baris.strip()
                if not baris:
                    continue
                item_match = re.match(
                    r"^\*\s*DVR\s*(?P<dvr>\d+)\s*Ch\s*(?P<ch>\d+)\s*\(\s*(?P<lokasi>[^)]+?)\s*\)\s*(?P<catatan>.*)$",
                    baris,
                    re.IGNORECASE,
                )
                if not item_match:
                    continue
                dvr = item_match.group("dvr")
                ch = item_match.group("ch")
                lokasi = item_match.group("lokasi").strip()
                catatan = item_match.group("catatan").strip()

                tower_m = re.search(r"Twr\s*([A-Za-z0-9]+)", lokasi, re.IGNORECASE)
                tower = f"Twr {tower_m.group(1).upper()}" if tower_m else ""
                lantai_m = re.search(r"(?:Lt\.?|Lantai)\s*(\d+)", lokasi, re.IGNORECASE)
                ptd_m = re.search(r"PTD\s*(\d+)", lokasi, re.IGNORECASE)

                temuan_text = f"Kamera CCTV mati: DVR {dvr} Ch {ch} ({lokasi})"
                if catatan:
                    temuan_text += f" - {catatan}"

                records.append({
                    "Jenis Temuan": "CCTV",
                    "Tower": tower,
                    "PTD": ptd_m.group(1) if ptd_m else "",
                    "Lantai": lantai_m.group(1) if lantai_m else "",
                    "DVR": dvr,
                    "Channel": ch,
                    "Temuan": temuan_text,
                    "Kategori": "CCTV Mati",
                    "PIC": "",
                    "Status": "Baru",
                })

    return records


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🧱 Pencatat Temuan Patroli Tower")
st.caption(
    "Tempel teks laporan patroli dari grup WA (temuan bangunan atau CCTV mati), "
    "sistem otomatis memecahnya jadi baris-baris data dan menyimpan ke file CSV."
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
            height=380,
            placeholder=(
                "Temuan patroli tower Clifford sbb :\n"
                "1). Termonitor Tembok Retak Di PTD 2 Lt 11\n"
                "2.) Termonitor Tembok Terkupas Di Lt 01 PTD 1\n\n"
                "Kamera CCTV yang mati 2 titik\n"
                "* DVR 4 Ch 5 (Twr A PTD2 kluar roof)\n"
                "* DVR 11 Ch 3 (Twr C Lt 19 Loby Lift) instalasi kabel rusak"
            ),
        )

    with col_b:
        tanggal_patroli = st.date_input("Tanggal Patroli", value=date.today())
        st.write("")
        parse_clicked = st.button("🔍 Parse Laporan", type="primary", use_container_width=True)
        st.info(
            "Mendukung dua format sekaligus dalam satu teks: temuan bangunan "
            "('Temuan patroli tower ... sbb :') dan temuan CCTV mati "
            "('Kamera CCTV yang mati ... titik').",
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
                    "'Temuan Patroli Tower ... sbb :' / 'Temuan patroli publik Area' "
                    "(numbering '1).') atau 'Kamera CCTV yang mati ... titik' (bullet '*')."
                )
            else:
                for r in hasil:
                    r["Tanggal Patroli"] = tanggal_patroli
                st.session_state.hasil_parse = pd.DataFrame(hasil)
                jml_cctv = sum(1 for r in hasil if r["Jenis Temuan"] == "CCTV")
                jml_bangunan = len(hasil) - jml_cctv
                st.success(
                    f"{len(hasil)} temuan terdeteksi ({jml_bangunan} bangunan, {jml_cctv} CCTV). "
                    "Silakan cek & edit di bawah sebelum menyimpan."
                )

    if st.session_state.hasil_parse is not None:
        st.subheader("Pratinjau & Edit Sebelum Disimpan")
        edited_df = st.data_editor(
            st.session_state.hasil_parse,
            column_config={
                "Jenis Temuan": st.column_config.SelectboxColumn("Jenis Temuan", options=JENIS_OPTIONS),
                "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS),
                "Kategori": st.column_config.SelectboxColumn("Kategori", options=KATEGORI_OPTIONS),
                "Tanggal Patroli": st.column_config.DateColumn("Tanggal Patroli"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="editor_temuan",
        )

        if st.button("💾 Simpan ke CSV", type="primary"):
            df_baru = edited_df.copy()
            df_baru["Waktu Input"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            df_baru["Keterangan Progress"] = ""
            df_baru["Tanggal Update"] = ""

            insert_rows(df_baru)

            st.session_state.hasil_parse = None
            total = len(load_data())
            st.success(f"Tersimpan ke {DATA_FILE}! Total data sekarang: {total} baris.")
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
        jml_cctv = int((df_all["Jenis Temuan"] == "CCTV").sum())

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Temuan", len(df_all))
        m2.metric("📷 CCTV Mati", jml_cctv)
        m3.metric("🆕 Baru", jml_baru)
        m4.metric("🔧 Dalam Proses", jml_proses)
        m5.metric("✅ Selesai", jml_selesai)

        st.divider()

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            filter_jenis = st.multiselect("Filter Jenis Temuan", JENIS_OPTIONS)
        with col2:
            filter_tower = st.multiselect("Filter Tower", sorted([t for t in df_all["Tower"].unique() if t]))
        with col3:
            filter_status = st.multiselect("Filter Status", STATUS_OPTIONS)
        with col4:
            cari = st.text_input("Cari kata kunci di kolom Temuan")

        df_filtered = df_all.copy()
        if filter_jenis:
            df_filtered = df_filtered[df_filtered["Jenis Temuan"].isin(filter_jenis)]
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
            df_all_indexed = df_all.set_index("ID")
            edited_indexed = edited_riwayat.set_index("ID")
            hari_ini = date.today().strftime("%Y-%m-%d")
            jumlah_berubah = 0

            for temuan_id, baris_baru in edited_indexed.iterrows():
                baris_lama = df_all_indexed.loc[temuan_id]
                status_berubah = baris_lama["Status"] != baris_baru["Status"]
                catatan_berubah = baris_lama["Keterangan Progress"] != baris_baru["Keterangan Progress"]
                pic_berubah = baris_lama["PIC"] != baris_baru["PIC"]
                if status_berubah or catatan_berubah or pic_berubah:
                    df_all_indexed.loc[temuan_id, "Status"] = baris_baru["Status"]
                    df_all_indexed.loc[temuan_id, "Keterangan Progress"] = baris_baru["Keterangan Progress"]
                    df_all_indexed.loc[temuan_id, "PIC"] = baris_baru["PIC"]
                    df_all_indexed.loc[temuan_id, "Tanggal Update"] = hari_ini
                    jumlah_berubah += 1

            df_final = df_all_indexed.reset_index()[COLUMNS]
            save_semua(df_final)
            st.success(f"{jumlah_berubah} baris berhasil diperbarui.")
            st.rerun()

        st.caption(f"Menampilkan {len(df_filtered)} dari {len(df_all)} total temuan.")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                "⬇️ Download CSV",
                data=df_all.to_csv(index=False).encode("utf-8"),
                file_name=DATA_FILE,
                mime="text/csv",
                use_container_width=True,
            )
        with col_dl2:
            st.download_button(
                "⬇️ Export ke Excel",
                data=export_ke_excel_bytes(df_all),
                file_name="data_temuan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.expander("📄 Lihat Isi File CSV Mentah"):
            st.caption(f"Isi apa adanya dari file **{os.path.abspath(DATA_FILE)}**.")
            if os.path.exists(DATA_FILE):
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    st.code(f.read(), language="text")
