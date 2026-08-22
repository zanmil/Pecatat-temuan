"""
Aplikasi Pencatat Temuan Patroli Tower
========================================
Tempel teks laporan patroli (format grup WA), aplikasi akan otomatis
memecahnya per tower & per temuan, lalu menyimpannya ke DATABASE SQLite
(file temuan_patroli.db) sehingga data yang sudah dikirim tersimpan permanen.

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""

import re
import io
import sqlite3
from datetime import date

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
DB_FILE = "temuan_patroli.db"
TABLE_NAME = "temuan"

# Nama kolom yang ditampilkan di UI (urutan ini juga dipakai di DataFrame)
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

# Nama kolom asli di tabel database (snake_case)
DB_COLUMN_MAP = {
    "ID": "id",
    "Tanggal Patroli": "tanggal_patroli",
    "Tower": "tower",
    "PTD": "ptd",
    "Lantai": "lantai",
    "Temuan": "temuan",
    "Kategori": "kategori",
    "PIC": "pic",
    "Status": "status",
    "Keterangan Progress": "keterangan_progress",
    "Tanggal Update": "tanggal_update",
    "Waktu Input": "waktu_input",
}
DB_TO_DISPLAY = {v: k for k, v in DB_COLUMN_MAP.items()}

STATUS_OPTIONS = ["Baru", "Dalam Proses", "Selesai"]
PROGRESS_OPTIONS = ["On Progress", "Selesai"]

st.set_page_config(page_title="Pencatat Temuan Patroli", page_icon="🧱", layout="wide")


# ---------------------------------------------------------------------------
# Fungsi Database (SQLite)
# ---------------------------------------------------------------------------
def get_conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_FILE)


def init_db() -> None:
    """Buat tabel database kalau belum ada."""
    with get_conn() as conn:
        conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal_patroli TEXT,
                tower TEXT,
                ptd TEXT,
                lantai TEXT,
                temuan TEXT,
                kategori TEXT,
                pic TEXT,
                status TEXT,
                keterangan_progress TEXT,
                tanggal_update TEXT,
                waktu_input TEXT
            )
            """
        )
        conn.commit()


def load_data() -> pd.DataFrame:
    """Ambil semua data dari database sebagai DataFrame (kolom pakai nama tampilan)."""
    init_db()
    with get_conn() as conn:
        df = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME} ORDER BY id", conn)

    if df.empty:
        return pd.DataFrame(columns=COLUMNS)

    df = df.rename(columns=DB_TO_DISPLAY)
    df["ID"] = df["ID"].astype(int)

    kolom_teks = [c for c in COLUMNS if c != "ID"]
    for col in kolom_teks:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)
        df.loc[df[col] == "nan", col] = ""

    return df[COLUMNS]


def insert_rows(df_baru: pd.DataFrame) -> None:
    """Simpan baris-baris temuan baru ke database (ID otomatis dari database)."""
    init_db()
    kolom_db = [
        "tanggal_patroli", "tower", "ptd", "lantai", "temuan",
        "kategori", "pic", "status", "keterangan_progress",
        "tanggal_update", "waktu_input",
    ]
    with get_conn() as conn:
        cur = conn.cursor()
        for _, row in df_baru.iterrows():
            nilai = (
                str(row.get("Tanggal Patroli", "")),
                str(row.get("Tower", "")),
                str(row.get("PTD", "")),
                str(row.get("Lantai", "")),
                str(row.get("Temuan", "")),
                str(row.get("Kategori", "")),
                str(row.get("PIC", "")),
                str(row.get("Status", "")),
                str(row.get("Keterangan Progress", "")),
                str(row.get("Tanggal Update", "")),
                str(row.get("Waktu Input", "")),
            )
            placeholder = ", ".join(["?"] * len(kolom_db))
            cur.execute(
                f"INSERT INTO {TABLE_NAME} ({', '.join(kolom_db)}) VALUES ({placeholder})",
                nilai,
            )
        conn.commit()


def update_row(temuan_id: int, status: str, keterangan: str, pic: str, tanggal_update: str) -> None:
    """Update status/progres/PIC satu baris temuan berdasarkan ID."""
    with get_conn() as conn:
        conn.execute(
            f"""
            UPDATE {TABLE_NAME}
            SET status = ?, keterangan_progress = ?, pic = ?, tanggal_update = ?
            WHERE id = ?
            """,
            (status, keterangan, pic, tanggal_update, int(temuan_id)),
        )
        conn.commit()


def export_ke_excel_bytes(df: pd.DataFrame) -> bytes:
    """Ubah DataFrame jadi file Excel (bytes) untuk didownload — dipakai untuk export saja,
    bukan sebagai sumber data utama (sumber data utama tetap database SQLite)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Temuan", index=False)
    return buffer.getvalue()


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
    "menjadi baris-baris temuan dan menyimpannya ke database (SQLite)."
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
            "status, PIC) sebelum disimpan ke database.",
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
                    "dan penomoran seperti '1).'."
                )
            else:
                for r in hasil:
                    r["Tanggal Patroli"] = tanggal_patroli
                st.session_state.hasil_parse = pd.DataFrame(hasil)
                st.success(
                    f"{len(hasil)} temuan berhasil terdeteksi. "
                    "Silakan cek & edit di bawah sebelum menyimpan."
                )

    if st.session_state.hasil_parse is not None:
        st.subheader("Pratinjau & Edit Sebelum Disimpan")
        edited_df = st.data_editor(
            st.session_state.hasil_parse,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status", options=STATUS_OPTIONS
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

        if st.button("💾 Simpan ke Database", type="primary"):
            df_baru = edited_df.copy()
            df_baru["Waktu Input"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            df_baru["Keterangan Progress"] = ""
            df_baru["Tanggal Update"] = ""

            insert_rows(df_baru)

            st.session_state.hasil_parse = None
            total = len(load_data())
            st.success(f"Tersimpan ke database! Total data sekarang: {total} baris.")
            st.rerun()

# ----------------------------- TAB RIWAYAT ----------------------------------
with tab_riwayat:
    df_all = load_data()
    if df_all.empty:
        st.info("Belum ada data tersimpan di database.")
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
                    update_row(
                        temuan_id,
                        baris_baru["Status"],
                        baris_baru["Keterangan Progress"],
                        baris_baru["PIC"],
                        hari_ini,
                    )
                    jumlah_berubah += 1

            st.success(f"{jumlah_berubah} baris berhasil diperbarui di database.")
            st.rerun()

        st.caption(f"Menampilkan {len(df_filtered)} dari {len(df_all)} total temuan.")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                "⬇️ Export ke Excel",
                data=export_ke_excel_bytes(df_all),
                file_name="data_temuan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            with open(DB_FILE, "rb") as f:
                st.download_button(
                    "⬇️ Download File Database (.db)",
                    data=f,
                    file_name=DB_FILE,
                    mime="application/octet-stream",
                    use_container_width=True,
                )
