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
DATA_FILE_DINDING = "data_temuan_dinding.csv"
DATA_FILE_CCTV = "data_temuan_cctv.csv"

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

STATUS_OPTIONS = ["Baru", "Dalam Proses", "Selesai"]
PROGRESS_OPTIONS = ["On Progress", "Selesai"]
KATEGORI_DINDING_OPTIONS = [
    "Tembok Retak",
    "Gompal/Terkupas",
    "Berlubang",
    "Berjamur",
    "Panel/Elektrikal",
    "Keamanan/Akses",
    "Lainnya",
]

st.set_page_config(page_title="Pencatat Temuan Patroli", page_icon="🧱", layout="wide")


# ---------------------------------------------------------------------------
# Fungsi Penyimpanan (CSV) — dipakai bersama oleh kedua fitur
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
    """Simpan data gabungan ke data_temuan.csv, SEKALIGUS otomatis regenerate
    data_temuan_dinding.csv dan data_temuan_cctv.csv — jadi kedua file itu
    selalu up-to-date tanpa perlu klik download manual."""
    df.to_csv(DATA_FILE, index=False)
    df[df["Jenis Temuan"] == "Bangunan"].to_csv(DATA_FILE_DINDING, index=False)
    df[df["Jenis Temuan"] == "CCTV"].to_csv(DATA_FILE_CCTV, index=False)


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


def terapkan_perubahan_status(df_all: pd.DataFrame, edited: pd.DataFrame) -> int:
    """Terapkan perubahan Status/Keterangan Progress/PIC dari tabel yang diedit
    (bisa berupa subset/filtered) ke seluruh data, lalu simpan. Return jumlah baris berubah."""
    df_all_indexed = df_all.set_index("ID")
    edited_indexed = edited.set_index("ID")
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

    save_semua(df_all_indexed.reset_index()[COLUMNS])
    return jumlah_berubah


def export_ke_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Temuan", index=False)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Parser — Temuan Dinding/Bangunan
# ---------------------------------------------------------------------------
def tebak_kategori_dinding(teks: str) -> str:
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


def parse_dinding(teks: str) -> list[dict]:
    """
    Parser khusus temuan dinding/bangunan. Mendukung heading:
      "Temuan patroli tower Clifford sbb :"      (per tower, numbering "1).")
      "Temuan patroli publik Area"                (tanpa nama tower/"sbb")
    Baris "@Nama Orang" di akhir blok dianggap PIC.
    """
    records: list[dict] = []

    pola_heading = re.compile(
        r"Temuan\s*Patroli\s*(?:Tower\s+(?P<tower>[A-Za-z]+)|(?P<area>Publik\s*Area|Public\s*Area))\s*(?:sbb\s*:?)?",
        re.IGNORECASE,
    )
    matches = list(pola_heading.finditer(teks))
    if not matches:
        return records

    for i, m in enumerate(matches):
        tower = "Publik Area" if m.group("area") else m.group("tower").strip().title()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(teks)
        blok = teks[start:end]

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
                "Kategori": tebak_kategori_dinding(temuan),
                "PIC": pic,
                "Status": "Baru",
            })

    return records


# ---------------------------------------------------------------------------
# Parser — Temuan CCTV Mati
# ---------------------------------------------------------------------------
def parse_cctv(teks: str) -> list[dict]:
    """
    Parser khusus temuan CCTV mati. Mengenali tiap baris bullet:
      "* DVR 4 Ch 5 (Twr A PTD2 kluar roof)"
      "* DVR 11 Ch 3 (Twr C Lt 19 Loby Lift) instalasi kabel rusak"
    Tidak butuh heading khusus — setiap baris yang cocok pola bullet akan
    langsung dianggap 1 temuan, jadi boleh ada beberapa section "Kamera CCTV
    yang mati ... titik" sekaligus dalam satu teks.
    """
    records: list[dict] = []
    pola_item = re.compile(
        r"^\*\s*DVR\s*(?P<dvr>\d+)\s*Ch\s*(?P<ch>\d+)\s*\(\s*(?P<lokasi>[^)]+?)\s*\)\s*(?P<catatan>.*)$",
        re.IGNORECASE,
    )

    for baris in teks.split("\n"):
        baris = baris.strip()
        if not baris:
            continue
        item_match = pola_item.match(baris)
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
st.caption("Dua fitur terpisah: Temuan Dinding/Bangunan dan Temuan CCTV Mati — masing-masing punya input & riwayat sendiri.")

for key in ("hasil_parse_dinding", "hasil_parse_cctv"):
    if key not in st.session_state:
        st.session_state[key] = None

tab_input, tab_riwayat = st.tabs(["📋 Input Laporan Baru", "📚 Riwayat Temuan"])

# ============================================================================
# TAB INPUT
# ============================================================================
with tab_input:
    subtab_in_dinding, subtab_in_cctv = st.tabs(["🧱 Temuan Dinding/Bangunan", "📷 Temuan CCTV Mati"])

    # ------------------------- INPUT: DINDING/BANGUNAN --------------------
    with subtab_in_dinding:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            teks_dinding = st.text_area(
                "Tempel teks laporan temuan dinding/bangunan di sini",
                height=350,
                placeholder=(
                    "Temuan patroli tower Clifford sbb :\n"
                    "1). Termonitor Tembok Retak Di PTD 2 Lt 11\n"
                    "2.) Termonitor Tembok Terkupas Di Lt 01 PTD 1\n"
                    "@Iqbaal Kensington"
                ),
                key="teks_dinding",
            )
        with col_b:
            tanggal_dinding = st.date_input("Tanggal Patroli", value=date.today(), key="tgl_dinding")
            st.write("")
            parse_dinding_clicked = st.button(
                "🔍 Parse Laporan Dinding", type="primary", use_container_width=True
            )
            st.info(
                "Format: 'Temuan patroli tower ... sbb :' atau "
                "'Temuan patroli publik Area', dengan penomoran '1).'.",
                icon="✏️",
            )

        if parse_dinding_clicked:
            if not teks_dinding.strip():
                st.warning("Teks laporan masih kosong.")
            else:
                hasil = parse_dinding(teks_dinding)
                if not hasil:
                    st.error(
                        "Tidak ada temuan dinding yang terdeteksi. Pastikan format masih memuat "
                        "'Temuan Patroli Tower ... sbb :' / 'Temuan patroli publik Area' dan "
                        "penomoran seperti '1).'."
                    )
                else:
                    for r in hasil:
                        r["Tanggal Patroli"] = tanggal_dinding
                    st.session_state.hasil_parse_dinding = pd.DataFrame(hasil)
                    st.success(f"{len(hasil)} temuan dinding terdeteksi. Cek & edit di bawah sebelum menyimpan.")

        if st.session_state.hasil_parse_dinding is not None:
            st.subheader("Pratinjau & Edit Sebelum Disimpan")
            edited_dinding = st.data_editor(
                st.session_state.hasil_parse_dinding,
                column_config={
                    "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS),
                    "Kategori": st.column_config.SelectboxColumn("Kategori", options=KATEGORI_DINDING_OPTIONS),
                    "Tanggal Patroli": st.column_config.DateColumn("Tanggal Patroli"),
                },
                num_rows="dynamic",
                use_container_width=True,
                key="editor_input_dinding",
            )

            if st.button("💾 Simpan Temuan Dinding ke CSV", type="primary"):
                df_baru = edited_dinding.copy()
                df_baru["Waktu Input"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                df_baru["Keterangan Progress"] = ""
                df_baru["Tanggal Update"] = ""
                insert_rows(df_baru)
                st.session_state.hasil_parse_dinding = None
                total = int((load_data()["Jenis Temuan"] == "Bangunan").sum())
                st.success(f"Tersimpan! Total temuan dinding sekarang: {total} baris.")
                st.rerun()

    # ------------------------- INPUT: CCTV MATI ----------------------------
    with subtab_in_cctv:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            teks_cctv = st.text_area(
                "Tempel teks laporan CCTV mati di sini",
                height=350,
                placeholder=(
                    "Kamera CCTV yang mati 2 titik\n"
                    "* DVR 4 Ch 5 (Twr A PTD2 kluar roof)\n"
                    "* DVR 11 Ch 3 (Twr C Lt 19 Loby Lift) instalasi kabel rusak"
                ),
                key="teks_cctv",
            )
        with col_b:
            tanggal_cctv = st.date_input("Tanggal Patroli", value=date.today(), key="tgl_cctv")
            st.write("")
            parse_cctv_clicked = st.button(
                "🔍 Parse Laporan CCTV", type="primary", use_container_width=True
            )
            st.info(
                "Format tiap baris: '* DVR <no> Ch <no> (lokasi) [catatan opsional]'. "
                "Tower/PTD/Lantai otomatis diambil dari teks lokasi.",
                icon="✏️",
            )

        if parse_cctv_clicked:
            if not teks_cctv.strip():
                st.warning("Teks laporan masih kosong.")
            else:
                hasil = parse_cctv(teks_cctv)
                if not hasil:
                    st.error(
                        "Tidak ada temuan CCTV yang terdeteksi. Pastikan tiap baris berformat "
                        "'* DVR <no> Ch <no> (lokasi)'."
                    )
                else:
                    for r in hasil:
                        r["Tanggal Patroli"] = tanggal_cctv
                    st.session_state.hasil_parse_cctv = pd.DataFrame(hasil)
                    st.success(f"{len(hasil)} titik CCTV mati terdeteksi. Cek & edit di bawah sebelum menyimpan.")

        if st.session_state.hasil_parse_cctv is not None:
            st.subheader("Pratinjau & Edit Sebelum Disimpan")
            edited_cctv = st.data_editor(
                st.session_state.hasil_parse_cctv,
                column_config={
                    "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS),
                    "Tanggal Patroli": st.column_config.DateColumn("Tanggal Patroli"),
                },
                num_rows="dynamic",
                use_container_width=True,
                key="editor_input_cctv",
            )

            if st.button("💾 Simpan Temuan CCTV ke CSV", type="primary"):
                df_baru = edited_cctv.copy()
                df_baru["Waktu Input"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                df_baru["Keterangan Progress"] = ""
                df_baru["Tanggal Update"] = ""
                insert_rows(df_baru)
                st.session_state.hasil_parse_cctv = None
                total = int((load_data()["Jenis Temuan"] == "CCTV").sum())
                st.success(f"Tersimpan! Total titik CCTV mati sekarang: {total} baris.")
                st.rerun()

# ============================================================================
# TAB RIWAYAT
# ============================================================================
with tab_riwayat:
    df_all = load_data()

    if df_all.empty:
        st.info("Belum ada data tersimpan.")
    else:
        subtab_riw_dinding, subtab_riw_cctv = st.tabs(["🧱 Riwayat Dinding/Bangunan", "📷 Riwayat CCTV Mati"])

        # ------------------------- RIWAYAT: DINDING/BANGUNAN --------------
        with subtab_riw_dinding:
            df_dinding = df_all[df_all["Jenis Temuan"] == "Bangunan"].copy()

            if df_dinding.empty:
                st.info("Belum ada temuan dinding/bangunan tersimpan.")
            else:
                jml_baru = int((df_dinding["Status"] == "Baru").sum())
                jml_proses = int((df_dinding["Status"] == "Dalam Proses").sum())
                jml_selesai = int((df_dinding["Status"] == "Selesai").sum())

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Temuan Dinding", len(df_dinding))
                m2.metric("🆕 Baru", jml_baru)
                m3.metric("🔧 Dalam Proses", jml_proses)
                m4.metric("✅ Selesai", jml_selesai)

                st.divider()

                col1, col2, col3 = st.columns(3)
                with col1:
                    filter_tower_d = st.multiselect(
                        "Filter Tower", sorted([t for t in df_dinding["Tower"].unique() if t]), key="ft_dinding"
                    )
                with col2:
                    filter_status_d = st.multiselect("Filter Status", STATUS_OPTIONS, key="fs_dinding")
                with col3:
                    cari_d = st.text_input("Cari kata kunci di kolom Temuan", key="cari_dinding")

                df_dinding_filtered = df_dinding.copy()
                if filter_tower_d:
                    df_dinding_filtered = df_dinding_filtered[df_dinding_filtered["Tower"].isin(filter_tower_d)]
                if filter_status_d:
                    df_dinding_filtered = df_dinding_filtered[df_dinding_filtered["Status"].isin(filter_status_d)]
                if cari_d:
                    df_dinding_filtered = df_dinding_filtered[
                        df_dinding_filtered["Temuan"].str.contains(cari_d, case=False, na=False)
                    ]

                st.subheader("🔄 Update Progres / Tandai Selesai")
                kolom_terkunci = [c for c in COLUMNS if c not in ("Status", "Keterangan Progress", "PIC")]

                edited_riwayat_d = st.data_editor(
                    df_dinding_filtered,
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
                    key="editor_riwayat_dinding",
                )

                if st.button("💾 Simpan Perubahan (Dinding)", type="primary", key="simpan_riwayat_dinding"):
                    jumlah_berubah = terapkan_perubahan_status(df_all, edited_riwayat_d)
                    st.success(f"{jumlah_berubah} baris temuan dinding berhasil diperbarui.")
                    st.rerun()

                st.caption(f"Menampilkan {len(df_dinding_filtered)} dari {len(df_dinding)} total temuan dinding.")
                st.caption(
                    f"📁 File **{DATA_FILE_DINDING}** otomatis ter-update di server setiap ada "
                    "simpan/perubahan — tombol di bawah cuma buat download salinannya ke komputer kamu."
                )

                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        "⬇️ Download CSV (Dinding)",
                        data=df_dinding.to_csv(index=False).encode("utf-8"),
                        file_name="data_temuan_dinding.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key="dl_csv_dinding",
                    )
                with col_dl2:
                    st.download_button(
                        "⬇️ Export ke Excel (Dinding)",
                        data=export_ke_excel_bytes(df_dinding),
                        file_name="data_temuan_dinding.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_xlsx_dinding",
                    )

        # ------------------------- RIWAYAT: CCTV MATI ----------------------
        with subtab_riw_cctv:
            df_cctv = df_all[df_all["Jenis Temuan"] == "CCTV"].copy()

            if df_cctv.empty:
                st.info("Belum ada temuan CCTV mati tersimpan.")
            else:
                jml_baru = int((df_cctv["Status"] == "Baru").sum())
                jml_proses = int((df_cctv["Status"] == "Dalam Proses").sum())
                jml_selesai = int((df_cctv["Status"] == "Selesai").sum())

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Titik CCTV Mati", len(df_cctv))
                m2.metric("🆕 Baru", jml_baru)
                m3.metric("🔧 Dalam Proses", jml_proses)
                m4.metric("✅ Selesai", jml_selesai)

                st.divider()

                col1, col2, col3 = st.columns(3)
                with col1:
                    filter_tower_c = st.multiselect(
                        "Filter Tower", sorted([t for t in df_cctv["Tower"].unique() if t]), key="ft_cctv"
                    )
                with col2:
                    filter_status_c = st.multiselect("Filter Status", STATUS_OPTIONS, key="fs_cctv")
                with col3:
                    cari_c = st.text_input("Cari (DVR/Channel/lokasi)", key="cari_cctv")

                df_cctv_filtered = df_cctv.copy()
                if filter_tower_c:
                    df_cctv_filtered = df_cctv_filtered[df_cctv_filtered["Tower"].isin(filter_tower_c)]
                if filter_status_c:
                    df_cctv_filtered = df_cctv_filtered[df_cctv_filtered["Status"].isin(filter_status_c)]
                if cari_c:
                    df_cctv_filtered = df_cctv_filtered[
                        df_cctv_filtered["Temuan"].str.contains(cari_c, case=False, na=False)
                    ]

                st.subheader("🔄 Update Progres / Tandai Selesai")
                kolom_terkunci = [c for c in COLUMNS if c not in ("Status", "Keterangan Progress", "PIC")]

                edited_riwayat_c = st.data_editor(
                    df_cctv_filtered,
                    column_config={
                        "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS),
                        "Keterangan Progress": st.column_config.SelectboxColumn(
                            "Keterangan Progress", options=PROGRESS_OPTIONS
                        ),
                        "PIC": st.column_config.TextColumn("PIC", help="Nama teknisi penanggung jawab"),
                        "ID": st.column_config.NumberColumn("ID", disabled=True),
                    },
                    disabled=kolom_terkunci,
                    hide_index=True,
                    use_container_width=True,
                    key="editor_riwayat_cctv",
                )

                if st.button("💾 Simpan Perubahan (CCTV)", type="primary", key="simpan_riwayat_cctv"):
                    jumlah_berubah = terapkan_perubahan_status(df_all, edited_riwayat_c)
                    st.success(f"{jumlah_berubah} baris temuan CCTV berhasil diperbarui.")
                    st.rerun()

                st.caption(f"Menampilkan {len(df_cctv_filtered)} dari {len(df_cctv)} total titik CCTV mati.")
                st.caption(
                    f"📁 File **{DATA_FILE_CCTV}** otomatis ter-update di server setiap ada "
                    "simpan/perubahan — tombol di bawah cuma buat download salinannya ke komputer kamu."
                )

                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        "⬇️ Download CSV (CCTV)",
                        data=df_cctv.to_csv(index=False).encode("utf-8"),
                        file_name="data_temuan_cctv.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key="dl_csv_cctv",
                    )
                with col_dl2:
                    st.download_button(
                        "⬇️ Export ke Excel (CCTV)",
                        data=export_ke_excel_bytes(df_cctv),
                        file_name="data_temuan_cctv.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_xlsx_cctv",
                    )

        st.divider()
        with st.expander("📄 Lihat Isi File CSV Mentah di Server"):
            pilihan_file = st.radio(
                "Pilih file",
                ["Gabungan (data_temuan.csv)", "Dinding saja (data_temuan_dinding.csv)", "CCTV saja (data_temuan_cctv.csv)"],
                horizontal=True,
            )
            path_terpilih = {
                "Gabungan (data_temuan.csv)": DATA_FILE,
                "Dinding saja (data_temuan_dinding.csv)": DATA_FILE_DINDING,
                "CCTV saja (data_temuan_cctv.csv)": DATA_FILE_CCTV,
            }[pilihan_file]

            st.caption(f"Isi apa adanya dari file **{os.path.abspath(path_terpilih)}**.")
            if os.path.exists(path_terpilih):
                with open(path_terpilih, "r", encoding="utf-8") as f:
                    st.code(f.read(), language="text")
            else:
                st.info("File belum ada — belum pernah ada data jenis ini yang disimpan.")
