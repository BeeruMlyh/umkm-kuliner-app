"""Project Charter: Analisis Penjualan UMKM Kuliner (Studi Kasus 2) - Kelompok 6.
Jalankan: streamlit run app.py   (file Excel harus satu folder dengan app.py)"""
import io
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

DATA_FILE = Path(__file__).with_name("Data_Transaksi_Simulasi_UMKM_Kuliner.xlsx")
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
PRI, ACC, MUTED = "#0E5A52", "#E09F1F", "#98A2B3"

st.set_page_config(page_title="Analisis Penjualan UMKM Kuliner", page_icon="🍜", layout="wide")

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stMarkdown, button, input { font-family: 'Plus Jakarta Sans', system-ui, sans-serif; }
#MainMenu, footer { visibility: hidden; }
.block-container { padding-top: 1.6rem; max-width: 1180px; }
.hero { background: #0E3B38; color: #fff; border-radius: 16px; padding: 34px 40px; margin-bottom: 22px; }
.hero h1 { color: #fff; font-size: 2.3rem; font-weight: 800; letter-spacing: -0.02em; margin: 10px 0 8px; line-height: 1.15; padding: 0; }
.hero p { color: #D5E5E2; font-size: 1.05rem; max-width: 62ch; margin: 0 0 14px; }
.hero .meta { color: #9FC4BE; font-size: .85rem; }
.badge { display: inline-block; background: #E09F1F; color: #1F2A37; font-weight: 700; font-size: .75rem; padding: 4px 10px; border-radius: 99px; }
.card { background: #fff; border: 1px solid #E4E7EC; border-radius: 12px; padding: 20px 22px; margin-bottom: 16px; }
.card h3 { margin: 0 0 8px; font-size: 1.05rem; font-weight: 700; color: #0E5A52; padding: 0; }
.card p, .card li, .card td { font-size: .93rem; color: #344054; line-height: 1.6; }
.card ul, .card ol { margin: 4px 0 0; padding-left: 20px; }
.card table { width: 100%; border-collapse: collapse; }
.card td { padding: 7px 0; border-bottom: 1px solid #EEF0F3; vertical-align: top; }
.card td:first-child { font-weight: 700; color: #0E5A52; width: 110px; }
.insight { background: #FFF7E6; border: 1px solid #F3D9A0; border-radius: 12px; padding: 14px 18px; margin: 4px 0 18px; color: #4A3B12; font-size: .95rem; }
[data-testid="stMetric"] { background: #fff; border: 1px solid #E4E7EC; border-radius: 12px; padding: 14px 18px; }
[data-testid="stMetricLabel"] p { color: #667085; font-size: .82rem; }
[data-testid="stMetricValue"] { color: #0E3B38; font-weight: 800; }
.stTabs [data-baseweb="tab-list"] { gap: 6px; border-bottom: 1px solid #E4E7EC; }
.stTabs [data-baseweb="tab"] { font-weight: 600; padding: 10px 16px; }
.note { color: #667085; font-size: .82rem; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def rp(n):
    return "Rp " + f"{n:,.0f}".replace(",", ".")


def num(n):
    return f"{n:,.0f}".replace(",", ".")


def card(judul, isi):
    st.markdown(f'<div class="card"><h3>{judul}</h3>{isi}</div>', unsafe_allow_html=True)


@st.cache_data(show_spinner="Membaca data...")
def muat(sumber):
    x = pd.ExcelFile(io.BytesIO(sumber) if isinstance(sumber, bytes) else sumber)
    menu = x.parse("Daftar_Menu", usecols="A:C")
    df = x.parse("Data_Bersih", usecols="A:F").dropna(subset=["ID_Transaksi"])
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Total"] = df["Jumlah"] * df["Harga"]
    df["JamBulat"] = df["Jam"].astype(str).str.split(":").str[0].astype(int)
    df["Kategori"] = df["Menu"].map(menu.set_index("Menu")["Kategori"])
    df["Hari"] = df["Tanggal"].dt.dayofweek.map(dict(enumerate(HARI)))
    log = x.parse("Log_Pembersihan")
    n_mentah = len(x.parse("Data_Mentah", usecols="A").dropna())
    return df, log, n_mentah


def ringkas_menu(d):
    g = d.groupby("Menu").agg(Jumlah_Terjual=("Jumlah", "sum"), Pendapatan=("Total", "sum")).reset_index()
    g["Kontribusi"] = g["Pendapatan"] / g["Pendapatan"].sum()
    return g


def ringkas_jam(d):
    g = d.groupby("JamBulat").agg(Transaksi=("ID_Transaksi", "nunique"), Item=("Jumlah", "sum"), Pendapatan=("Total", "sum"))
    g = g.reindex(range(8, 21), fill_value=0).reset_index().rename(columns={"JamBulat": "JamN"})
    g["Jam"] = g["JamN"].map(lambda h: f"{h:02d}:00")
    return g


def ringkas_hari(d):
    g = d.groupby("Hari").agg(Pendapatan=("Total", "sum"), Hari_Data=("Tanggal", "nunique")).reindex(HARI).dropna().reset_index()
    g["Rata_per_Hari"] = g["Pendapatan"] / g["Hari_Data"]
    return g


def style(ch):
    return ch.configure_view(stroke=None).configure_axis(grid=False, domain=False, labelColor="#475467", ticks=False)


def grafik_batang(d, kolom, kat, horizontal=True, tinggi=None):
    d = d.copy()
    d["Top"] = d[kolom] == d[kolom].max()
    d["Label"] = d[kolom].map(num)
    warna = alt.Color("Top:N", scale=alt.Scale(domain=[True, False], range=[ACC, PRI]), legend=None)
    tip = [kat, alt.Tooltip("Label:N", title=kolom)]
    if horizontal:
        base = alt.Chart(d).encode(y=alt.Y(f"{kat}:N", sort="-x", title=None), x=alt.X(f"{kolom}:Q", title=None, axis=None))
        batang = base.mark_bar(cornerRadiusEnd=4, size=18).encode(color=warna, tooltip=tip)
        teks = base.mark_text(align="left", dx=5, color="#344054", fontSize=11).encode(text="Label:N")
        ch = (batang + teks).properties(height=tinggi or 30 * len(d))
    else:
        base = alt.Chart(d).encode(x=alt.X(f"{kat}:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)), y=alt.Y(f"{kolom}:Q", title=None, axis=None))
        batang = base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(color=warna, tooltip=tip)
        teks = base.mark_text(dy=-7, color="#344054", fontSize=11).encode(text="Label:N")
        ch = (batang + teks).properties(height=tinggi or 280)
    return style(ch)


# ---------------------------------------------------------------- sumber data
sumber = None
if DATA_FILE.exists():
    sumber = str(DATA_FILE)
else:
    up = st.sidebar.file_uploader("Unggah Data_Transaksi_Simulasi_UMKM_Kuliner.xlsx", type="xlsx")
    if up:
        sumber = up.getvalue()
if sumber is None:
    st.error("File data tidak ditemukan. Letakkan Data_Transaksi_Simulasi_UMKM_Kuliner.xlsx satu folder dengan app.py, atau unggah lewat sidebar.")
    st.stop()
df, log, n_mentah = muat(sumber)

st.markdown(
    '<div class="hero"><span class="badge">Data simulasi</span>'
    "<h1>Analisis Penjualan UMKM Kuliner</h1>"
    "<p>Menu apa yang paling menguntungkan dan jam berapa warung paling ramai? Project charter berbasis 30 hari data transaksi.</p>"
    '<div class="meta">Pengantar Data Sains, Studi Kasus 2 &nbsp;|&nbsp; Kelompok 6 &nbsp;|&nbsp; Dosen pengampu: Nurfiyah S.T., M.Kom &nbsp;|&nbsp; Disusun 18 Oktober 2026</div></div>',
    unsafe_allow_html=True,
)

t1, t2, t3, t4 = st.tabs(["Charter", "Data dan pembersihan", "Dashboard analisis", "Rencana dan evaluasi"])

# ---------------------------------------------------------------- 1. charter
with t1:
    a, b = st.columns(2)
    with a:
        card("Latar belakang dan masalah",
             "<p>UMKM kuliner mencatat transaksi setiap hari, tetapi catatan itu jarang diolah. Menu yang terasa laris belum tentu menyumbang pendapatan terbesar, dan jam yang terasa ramai belum tentu paling menghasilkan.</p>"
             "<p>Selama 30 hari tercatat sekitar 1.400 transaksi dari 12 menu (data simulasi), tetapi belum ada ringkasan menu terlaris, menu berpendapatan terbesar, dan jam puncak. Keputusan stok dan jadwal persiapan masih berdasarkan perkiraan.</p>")
        card("Tiga pertanyaan analitik",
             "<ol><li>Menu apa yang paling banyak terjual dan menu apa yang berpendapatan terbesar, beserta % kontribusinya?</li>"
             "<li>Pada jam berapa pendapatan tertinggi (pembanding: jumlah transaksi)?</li>"
             "<li>Rekomendasi apa (minimal 3) yang dapat disusun untuk persediaan dan jadwal persiapan?</li></ol>")
        card("Ruang lingkup",
             "<p><b>Masuk:</b> jumlah dan pendapatan per menu, pola penjualan per jam, visualisasi, rekomendasi bisnis.</p>"
             "<p><b>Tidak masuk:</b> peramalan, laba bersih (data biaya tidak tersedia), kepuasan pelanggan, pemasaran di luar data transaksi.</p>"
             "<p>Tanpa data biaya, “menguntungkan” berarti pendapatan terbesar.</p>")
    with b:
        card("Tujuan SMART",
             "<table><tr><td>Specific</td><td>Mengidentifikasi menu terlaris, menu berpendapatan tertinggi, dan jam berpendapatan tertinggi.</td></tr>"
             "<tr><td>Measurable</td><td>Peringkat menu, % kontribusi pendapatan, jam puncak, minimal 3 rekomendasi bisnis.</td></tr>"
             "<tr><td>Achievable</td><td>Analisis deskriptif dan visualisasi dengan Excel, Google Sheets, atau Python/Colab.</td></tr>"
             "<tr><td>Relevant</td><td>Mendukung keputusan persediaan dan jadwal operasional pemilik UMKM.</td></tr>"
             "<tr><td>Time-bound</td><td>Data 30 hari; proyek selesai dalam 6 minggu (19 Oktober–29 November 2026).</td></tr></table>")
        card("Pemangku kepentingan",
             "<table><tr><td>Pemilik UMKM</td><td>Pengambil keputusan utama: menu unggulan, persediaan, jadwal persiapan.</td></tr>"
             "<tr><td>Pegawai</td><td>Pencatatan konsisten; tugas sesuai jam sibuk.</td></tr>"
             "<tr><td>Pelanggan</td><td>Pihak terdampak: menu tersedia, layanan tidak terlalu lama saat ramai.</td></tr>"
             "<tr><td>Dosen</td><td>Menilai pemahaman konsep dasar sains data.</td></tr></table>")
        card("Rumus utama",
             "<p>Total pendapatan = jumlah terjual × harga jual per unit<br>Pendapatan menu = Σ pendapatan transaksi menu tersebut<br>Kontribusi (%) = pendapatan menu ÷ total pendapatan × 100%</p>")

# ---------------------------------------------------------------- 2. data
with t2:
    k = st.columns(4)
    k[0].metric("Baris data mentah", num(n_mentah))
    k[1].metric("Baris data bersih", num(len(df)))
    k[2].metric("Baris dikeluarkan", num(n_mentah - len(df)))
    k[3].metric("Temuan di log", num(len(log)))
    a, b = st.columns([1.1, 1])
    with a:
        st.subheader("Jenis masalah pada data mentah")
        tm = log["Masalah"].value_counts().rename_axis("Masalah").reset_index(name="Jumlah")
        st.altair_chart(grafik_batang(tm, "Jumlah", "Masalah"), use_container_width=True)
        st.markdown('<p class="note">Kesalahan sengaja disuntikkan ke data simulasi sebagai bahan latihan pembersihan. Setiap perubahan tercatat di sheet Log_Pembersihan.</p>', unsafe_allow_html=True)
    with b:
        st.subheader("Sembilan langkah persiapan data")
        langkah = [
            ("Periksa kelengkapan data", "Mengetahui lokasi dan jumlah nilai kosong."),
            ("Tangani data duplikat", "Pencatatan ganda menggandakan jumlah dan pendapatan."),
            ("Seragamkan format tanggal dan jam", "Pengelompokan per jam memerlukan format seragam."),
            ("Periksa harga dan jumlah tidak valid", "Nilai nol, negatif, atau janggal merusak perhitungan."),
            ("Tangani nilai kosong", "Harga diisi dari harga baku menu; jam diisi dari baris lain pada transaksi yang sama."),
            ("Hitung total pendapatan per baris", "Menyiapkan variabel turunan."),
            ("Buat variabel turunan", "Hari, kelompok jam, dan kategori menu untuk membaca pola."),
            ("Kelompokkan per menu dan per jam", "Menyiapkan tabel ringkasan."),
            ("Pastikan hasil konsisten", "Total setelah pengelompokan harus sama dengan total data bersih."),
        ]
        for i, (j, s) in enumerate(langkah, 1):
            with st.expander(f"{i}. {j}"):
                st.write(s)
    st.subheader("Contoh log pembersihan")
    st.dataframe(log.head(40), hide_index=True, use_container_width=True)
    st.download_button("Unduh data bersih (CSV)", df[["ID_Transaksi", "Tanggal", "Menu", "Jumlah", "Jam", "Harga", "Total"]].to_csv(index=False).encode("utf-8"),
                       file_name="data_bersih_umkm_kuliner.csv", mime="text/csv")

# ---------------------------------------------------------------- 3. dashboard
with t3:
    with st.sidebar:
        st.markdown("### Filter dashboard")
        kat = st.multiselect("Kategori menu", sorted(df["Kategori"].dropna().unique()), default=sorted(df["Kategori"].dropna().unique()))
        hari = st.multiselect("Hari", HARI, default=HARI)
        st.caption("Filter berlaku pada tab Dashboard analisis.")
    d = df[df["Kategori"].isin(kat) & df["Hari"].isin(hari)]
    if d.empty:
        st.warning("Tidak ada data untuk filter ini. Pilih minimal satu kategori dan satu hari.")
    else:
        tot = d["Total"].sum()
        trx = d["ID_Transaksi"].nunique()
        k = st.columns(4)
        k[0].metric("Total pendapatan", rp(tot))
        k[1].metric("Jumlah transaksi", num(trx))
        k[2].metric("Item terjual", num(d["Jumlah"].sum()))
        k[3].metric("Rata-rata per transaksi", rp(tot / trx))
        m, j, h = ringkas_menu(d), ringkas_jam(d), ringkas_hari(d)
        q, r = m.loc[m["Jumlah_Terjual"].idxmax()], m.loc[m["Pendapatan"].idxmax()]
        pk = j.loc[j["Pendapatan"].idxmax()]
        sama = q["Menu"] == r["Menu"]
        kontrib = f"{r['Kontribusi']:.1%}".replace(".", ",")
        teks_menu = (f"<b>{q['Menu']}</b> unggul di jumlah terjual dan pendapatan." if sama else
                     f"Menu terlaris adalah <b>{q['Menu']}</b> ({num(q['Jumlah_Terjual'])} item), tetapi pendapatan terbesar datang dari <b>{r['Menu']}</b> ({rp(r['Pendapatan'])}, {kontrib} dari total). Terlaris belum tentu terbesar.")
        st.markdown(f'<div class="insight">{teks_menu} Jam pendapatan tertinggi: <b>{pk["Jam"]}–{pk["JamN"]:02d}:59</b> ({rp(pk["Pendapatan"])}).</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            st.subheader("Menu")
            ukuran = st.radio("Ukuran", ["Jumlah terjual", "Pendapatan"], horizontal=True, label_visibility="collapsed")
            kol = "Jumlah_Terjual" if ukuran == "Jumlah terjual" else "Pendapatan"
            st.altair_chart(grafik_batang(m, kol, "Menu"), use_container_width=True)
        with b:
            st.subheader("Jam")
            ukj = st.radio("Ukuran jam", ["Pendapatan", "Transaksi"], horizontal=True, label_visibility="collapsed")
            st.altair_chart(grafik_batang(j, ukj, "Jam", horizontal=False, tinggi=360), use_container_width=True)
        st.subheader("Rata-rata pendapatan per hari dalam minggu")
        st.altair_chart(grafik_batang(h, "Rata_per_Hari", "Hari", horizontal=False, tinggi=220), use_container_width=True)
        st.subheader("Tabel ringkasan menu")
        tabel = m.sort_values("Pendapatan", ascending=False).assign(
            Peringkat_Terjual=lambda x: x["Jumlah_Terjual"].rank(ascending=False, method="min").astype(int),
            Peringkat_Pendapatan=lambda x: x["Pendapatan"].rank(ascending=False, method="min").astype(int))
        st.dataframe(tabel, hide_index=True, use_container_width=True, column_config={
            "Pendapatan": st.column_config.NumberColumn("Pendapatan (Rp)", format="%d"),
            "Kontribusi": st.column_config.NumberColumn("Kontribusi", format="percent")})
        selisih = m["Pendapatan"].sum() - tot
        (st.success if selisih == 0 else st.error)(f"Cek konsistensi: total per menu dikurangi total data bersih = Rp {num(selisih)}.")
        st.markdown('<p class="note">Seluruh angka berasal dari data simulasi dan hanya berlaku untuk data ini, bukan kondisi usaha nyata.</p>', unsafe_allow_html=True)

# ---------------------------------------------------------------- 4. rencana
with t4:
    a, b = st.columns([1.1, 1])
    with a:
        st.subheader("Jadwal enam minggu")
        st.dataframe(pd.DataFrame({
            "Minggu": ["1 (19–25 Okt)", "2 (26 Okt–1 Nov)", "3 (2–8 Nov)", "4 (9–15 Nov)", "5 (16–22 Nov)", "6 (23–29 Nov)"],
            "Tahapan": ["Perumusan masalah dan tujuan", "Pengumpulan/pembuatan data", "Pembersihan dan persiapan data", "Analisis dan visualisasi", "Evaluasi hasil", "Laporan dan presentasi"],
            "Keluaran": ["Masalah, tujuan SMART, pertanyaan analitik", "Dataset mentah 30 hari", "Dataset bersih dan log pembersihan", "Tabel ringkasan dan grafik", "Pemeriksaan indikator, rekomendasi", "Laporan akhir dan bahan presentasi"]}),
            hide_index=True, use_container_width=True)
        st.subheader("Risiko dan mitigasi")
        st.dataframe(pd.DataFrame({
            "Risiko": ["Data simulasi tidak lengkap atau tidak konsisten", "Data simulasi tidak mencerminkan kondisi nyata", "Kesalahan perhitungan pendapatan", "Kesalahan pengelompokan waktu", "Kesimpulan terlalu luas"],
            "Mitigasi": ["Periksa sejak awal; catat setiap perbaikan di log", "Nyatakan sebagai simulasi; batasi kesimpulan pada data tersebut", "Periksa rumus pada sampel baris; bandingkan total sebelum dan sesudah pengelompokan", "Seragamkan format jam; periksa hasil pada sampel", "Cantumkan batasan 30 hari dan satu usaha"]}),
            hide_index=True, use_container_width=True)
    with b:
        card("Indikator keberhasilan",
             "<ul><li>100% baris lolos pemeriksaan dan tercatat di log pembersihan.</li>"
             "<li>Peringkat menu (jumlah terjual dan pendapatan) beserta % kontribusi tersusun.</li>"
             "<li>Jam dengan pendapatan tertinggi teridentifikasi, jumlah transaksi sebagai pembanding.</li>"
             "<li>Minimal 3 grafik yang angkanya sama dengan tabel.</li>"
             "<li>Minimal 3 rekomendasi, masing-masing mengacu pada satu angka temuan.</li>"
             "<li>Selisih total sesudah pengelompokan dan total data bersih = Rp 0.</li></ul>")
        card("Rencana penyampaian hasil",
             "<ul><li>Dataset bersih dan tabel ringkasan</li><li>Dashboard visualisasi penjualan</li><li>Laporan interpretasi (±10 halaman)</li><li>Minimal tiga rekomendasi untuk pemilik UMKM</li></ul>")
        card("Catatan penting",
             "<p>Dampak bisnis (kenaikan pendapatan, pengurangan sisa bahan) belum dapat dipastikan dan perlu diuji pemilik usaha. Proyek ini tidak menjanjikan peningkatan pendapatan.</p>")
