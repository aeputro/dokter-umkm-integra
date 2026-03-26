import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import sqlite3
import io
import urllib.parse
from datetime import datetime
from fpdf import FPDF

# --- CONFIG & SECURITY ---
ADMIN_PASSWORD = "Integra_2026" 
NAMA_KONSULTAN = "Arjuno Eko Putro"

# --- 1. DATABASE SETUP ---
conn = sqlite3.connect('database_leads_umkm.db', check_same_thread=False)
c = conn.cursor()
c.execute('''
    CREATE TABLE IF NOT EXISTS leads_umkm (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama_user TEXT,
        nama_usaha TEXT,
        kategori_usaha TEXT,
        email TEXT,
        no_hp TEXT,
        waktu_daftar TEXT
    )
''')
conn.commit()

# --- 2. FUNGSI HELPER ---

class PDFReport(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(0, 10, 'LAPORAN DIAGNOSIS KEUANGAN UMKM', 0, 1, 'C')
        self.set_draw_color(0, 102, 204)
        self.line(10, 22, 200, 22)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Hasil Analisis Otomatis | Konsultan: {NAMA_KONSULTAN} | Hal {self.page_no()}', 0, 0, 'C')

def generate_pdf(name, usaha, oz, hpp, op, profit, insights):
    pdf = PDFReport()
    pdf.add_page()
    
    # Identitas
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"Pemilik: {name}", ln=1)
    pdf.cell(0, 8, f"Nama Usaha: {usaha}", ln=1)
    pdf.cell(0, 8, f"Tanggal Laporan: {datetime.now().strftime('%d %B %Y')}", ln=1)
    pdf.ln(5)

    # Ringkasan Angka
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Helvetica', 'B', 11)
    pdf.cell(95, 10, " KATEGORI", 1, 0, 'L', True)
    pdf.cell(95, 10, " JUMLAH", 1, 1, 'L', True)
    
    pdf.set_font('Helvetica', '', 11)
    pdf.cell(95, 10, " Total Omzet", 1, 0)
    pdf.cell(95, 10, f" Rp {oz:,.0f}", 1, 1)
    pdf.cell(95, 10, " Total HPP (Modal)", 1, 0)
    pdf.cell(95, 10, f" Rp {hpp:,.0f}", 1, 1)
    pdf.cell(95, 10, " Total Operasional", 1, 0)
    pdf.cell(95, 10, f" Rp {op:,.0f}", 1, 1)
    pdf.set_font('Helvetica', 'B', 11)
    pdf.cell(95, 10, " LABA/RUGI BERSIH", 1, 0, 'L', True)
    pdf.cell(95, 10, f" Rp {profit:,.0f}", 1, 1, 'L', True)
    pdf.ln(10)

    # GRAFIK PERBANDINGAN DI PDF (Manual Draw)
    pdf.set_font('Helvetica', 'B', 13)
    pdf.cell(0, 10, "Visualisasi Komposisi Biaya:", ln=1)
    
    total_biaya = hpp + op
    if oz > 0:
        # Bar Omzet (Full 100%)
        pdf.set_fill_color(0, 200, 0) # Hijau
        pdf.rect(10, pdf.get_y(), 180, 8, 'F')
        pdf.set_font('Helvetica', 'B', 8)
        pdf.set_text_color(255, 255, 255)
        pdf.text(12, pdf.get_y()+6, f"OMZET (100%) - Rp {oz:,.0f}")
        pdf.ln(10)
        
        # Bar HPP
        lebar_hpp = (hpp / oz) * 180 if oz > hpp else 180
        pdf.set_fill_color(200, 0, 0) # Merah
        pdf.rect(10, pdf.get_y(), lebar_hpp, 8, 'F')
        pdf.set_text_color(0, 0, 0)
        pdf.text(12, pdf.get_y()+12, f"Porsi HPP: {int((hpp/oz)*100)}%")
        pdf.ln(15)

    # INSIGHT STRATEGIS
    pdf.set_font('Helvetica', 'B', 13)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(0, 10, "Insight & Rekomendasi Strategis:", ln=1)
    pdf.set_font('Helvetica', '', 11)
    pdf.set_text_color(0, 0, 0)
    
    lebar_teks = pdf.w - 20
    for i in insights:
        clean_text = "".join([c for c in str(i) if ord(c) < 256])
        for char in ['📈', '📉', '⚠️', 'ℹ️', '✅', '🏁', '🎯', '🚀', '💸']:
            clean_text = clean_text.replace(char, '')
        pdf.multi_cell(lebar_teks, 8, f"- {clean_text.strip()}")
    
    return bytes(pdf.output())

def get_excel_template():
    df_t = pd.DataFrame({
        'Tanggal': ['2026-03-01', '2026-03-02', '2026-03-03'],
        'Keterangan': ['Penjualan Produk', 'Beli Stok Bahan', 'Gaji & Listrik'],
        'Kategori': ['Omzet', 'HPP', 'Operasional'],
        'Jumlah': [10000000, 4000000, 2000000]
    })
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_t.to_excel(writer, index=False)
    return out.getvalue()

def buat_link_wa(nomor, nama, usaha):
    no = nomor.replace(" ", "").replace("-", "").replace("+", "")
    if no.startswith("0"): no = "62" + no[1:]
    msg = f"Halo {nama}, saya {NAMA_KONSULTAN}. Saya telah melihat data bisnis {usaha} Anda. Ada potensi menarik yang bisa kita kembangkan bersama..."
    return f"https://wa.me/{no}?text={urllib.parse.quote(msg)}"

# --- 3. UI ---
st.set_page_config(page_title="UMKM Financial Doctor", layout="wide")

if 'admin_logged_in' not in st.session_state: st.session_state.admin_logged_in = False
if 'registered' not in st.session_state: st.session_state.registered = False

with st.sidebar:
    st.title("🛡️ Panel")
    if not st.session_state.admin_logged_in:
        p_in = st.text_input("Sandi Admin", type="password")
        if st.button("Login Admin"):
            if p_in == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.rerun()
    else:
        st.success("Admin Aktif")
        if st.button("🚪 Logout Admin"):
            st.session_state.admin_logged_in = False
            st.rerun()

# --- 4. LOGIKA TAMPILAN ---

if st.session_state.admin_logged_in:
    st.header("👨‍💻 Admin Dashboard")
    df_l = pd.read_sql_query('SELECT * FROM leads_umkm ORDER BY id DESC', conn)
    if not df_l.empty:
        st.subheader("📱 Hubungi Database UMKM")
        for i, r in df_l.iterrows():
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            c1.write(f"**{r['nama_user']}**")
            c2.write(f"{r['nama_usaha']}")
            c3.write(r['no_hp'])
            c4.link_button("💬 Chat WA", buat_link_wa(r['no_hp'], r['nama_user'], r['nama_usaha']))
            st.divider()
    else: st.info("Belum ada data.")

elif not st.session_state.registered:
    st.title("🩺 UMKM Financial Doctor")
    st.subheader("Diagnosis Bisnis Anda Secara Instan")
    with st.form("reg"):
        co1, co2 = st.columns(2)
        with co1:
            u_n = st.text_input("Nama Anda *")
            u_u = st.text_input("Nama Usaha *")
            u_k = st.selectbox("Bidang", ["Kuliner", "Retail", "Jasa", "Fashion", "Lainnya"])
        with co2:
            u_e = st.text_input("Email *")
            u_h = st.text_input("Nomor WA *")
        if st.form_submit_button("Mulai Diagnosis 🚀"):
            if u_n and u_u and u_h:
                tgl = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute('INSERT INTO leads_umkm (nama_user, nama_usaha, kategori_usaha, email, no_hp, waktu_daftar) VALUES (?,?,?,?,?,?)',
                          (u_n, u_u, u_k, u_e, u_h, tgl))
                conn.commit()
                st.session_state.registered, st.session_state.user_info = True, {"nama": u_n, "usaha": u_u}
                st.rerun()

else:
    st.title(f"Dashboard: {st.session_state.user_info['usaha']}")
    c1, c2 = st.columns(2)
    c1.download_button("📥 Unduh Template", get_excel_template(), "template.xlsx")
    up = c2.file_uploader("Upload Excel Anda", type=["xlsx"])

    if up:
        df = pd.read_excel(up)
        oz = df[df['Kategori'] == 'Omzet']['Jumlah'].sum()
        hpp = df[df['Kategori'] == 'HPP']['Jumlah'].sum()
        op = df[df['Kategori'] == 'Operasional']['Jumlah'].sum()
        profit = oz - hpp - op
        margin = (profit/oz*100) if oz > 0 else 0
        
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Omzet", f"Rp {oz:,.0f}")
        m2.metric("Laba Bersih", f"Rp {profit:,.0f}", delta=f"Margin {margin:.1f}%")
        m3.metric("Total Beban", f"Rp {hpp+op:,.0f}")

        # --- GRAFIK DASHBOARD ---
        st.subheader("📊 Visualisasi Keuangan")
        fig = go.Figure(data=[
            go.Bar(name='Pemasukan', x=['Omzet'], y=[oz], marker_color='green'),
            go.Bar(name='HPP (Modal)', x=['Biaya'], y=[hpp], marker_color='orange'),
            go.Bar(name='Operasional', x=['Biaya'], y=[op], marker_color='red')
        ])
        fig.update_layout(barmode='group', height=400)
        st.plotly_chart(fig, use_container_width=True)

        # --- SMART AI INSIGHTS ---
        st.subheader("🧠 Rekomendasi Dokter")
        insights = []
        if profit > 0: insights.append(f"✅ Bisnis Profit: Anda menghasilkan Laba Bersih Rp {profit:,.0f}.")
        else: insights.append(f"⚠️ Bisnis Minus: Anda merugi Rp {abs(profit):,.0f}. Segera efisiensi biaya.")
        
        if margin < 15: insights.append("📉 Margin Tipis: Keuntungan di bawah 15%. Coba evaluasi supplier atau naikkan harga.")
        
        ratio_hpp = hpp/oz if oz > 0 else 0
        if ratio_hpp < 1:
            bep = op / (1 - ratio_hpp)
            insights.append(f"🏁 Target BEP: Minimal Omzet harus Rp {bep:,.0f} agar tidak merugi.")
            if oz < bep: insights.append("🚀 Perhatian: Anda butuh tambahan penjualan untuk mencapai titik impas.")
        
        for i in insights: st.info(i)

        if st.button("📄 Generate Laporan PDF"):
            pdf_b = generate_pdf(st.session_state.user_info['nama'], st.session_state.user_info['usaha'], oz, hpp, op, profit, insights)
            st.download_button("⬇️ Download PDF", pdf_b, "Laporan_Dokter_UMKM.pdf", "application/pdf")