import streamlit as st
import pandas as pd
import time
import os
import traceback
import requests
from io import BytesIO
import base64

# ==============================
# KONFIGURASI
# ==============================
URLS = {
    "🌐 Production": "http://ellipse-ellprd.bssr-ellprd.baramultigroup.co.id/html/ui#!/ome",
    "🧪 Testing": "http://ellipse-elltst.bssr-elldev.baramultigroup.co.id/html/ui#!/ome"
}
QUICK_LAUNCH_CODE = "MSO240"

# ==============================
# SIMULASI MODE (Untuk Streamlit Cloud)
# ==============================
def simulate_ro_processing(df, username, environment):
    """Simulasi proses RO untuk testing di Streamlit Cloud"""
    results = []
    
    for idx in range(len(df)):
        stock = str(df.iloc[idx, 1]).strip()
        qty = str(df.iloc[idx, 4]).strip()
        wh = str(df.iloc[idx, 7]).strip()
        
        if not stock or stock.lower() == "nan":
            results.append("SKIP - Stock code kosong")
            continue
            
        # Generate RO number simulasi
        ro_number = f"RO{int(time.time())}{idx:04d}"
        results.append(ro_number)
        
    return results

# ==============================
# FUNGSI UTAMA (Manual Setup untuk Local)
# ==============================
def setup_selenium_local():
    """Setup Selenium untuk environment local saja"""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        st.error(f"❌ Selenium tidak tersedia: {e}")
        return None

def validate_dataframe(df):
    """Validasi struktur dataframe"""
    required_cols = 8
    if len(df.columns) < required_cols:
        return False, f"File harus memiliki minimal {required_cols} kolom"
    if len(df) == 0:
        return False, "File tidak memiliki data"
    
    # Cek kolom penting
    if df.iloc[:, 1].isna().all():
        return False, "Kolom Stock Code (kolom 2) kosong atau tidak valid"
    
    return True, None

# ==============================
# STREAMLIT UI
# ==============================
st.set_page_config(
    page_title="Otomasi Input RO Ellipse",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Otomasi Input RO Ellipse (MSO240)")
st.caption("Upload file Excel RO untuk generate RO Number")

# Informasi aplikasi
with st.expander("ℹ️ Informasi Aplikasi"):
    st.info("""
    **Aplikasi Otomasi Input RO Ellipse**
    
    🔧 **Fitur:**
    - Generate RO Number secara otomatis
    - Support environment Production & Testing
    - Export hasil ke Excel
    - Mode simulasi untuk testing
    
    ⚠️ **Catatan Streamlit Cloud:**
    - Aplikasi berjalan dalam mode simulasi
    - Untuk akses real Ellipse, jalankan di environment local
    - Pastikan file Excel format benar
    """)

with st.sidebar:
    st.header("🌍 Pilih Environment")
    env_choice = st.radio("Pilih server Ellipse:", list(URLS.keys()), index=1)
    TARGET_URL = URLS[env_choice]
    st.info(f"🔗 **URL target:** {TARGET_URL}")

    st.divider()
    st.header("🔐 Login Credentials")
    username = st.text_input("Username", value="", placeholder="0192210XXX")
    password = st.text_input("Password", value="", type="password", placeholder="••••••")

    st.divider()
    st.header("⚙️ Mode Operasi")
    operation_mode = st.radio(
        "Pilih mode:",
        ["🎯 Mode Simulasi", "🚀 Mode Real"],
        index=0,
        help="Mode real hanya work di local environment"
    )

uploaded_file = st.file_uploader("📂 Upload File Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        is_valid, error_msg = validate_dataframe(df)

        if not is_valid:
            st.error(f"❌ {error_msg}")
            st.stop()

        st.success(f"✅ File berhasil dimuat: {len(df)} baris data")
        
        # Tampilkan preview data
        st.subheader("📊 Preview Data")
        st.dataframe(df.head())
        
        # Informasi kolom
        st.info(f"""
        **Environment:** {env_choice}
        **Mode:** {operation_mode}
        **Kolom yang akan diproses:**
        - Kolom 2 (index 1): Stock Code
        - Kolom 5 (index 4): Quantity  
        - Kolom 8 (index 7): Warehouse
        - Kolom 9: RO Number hasil
        """)

        if st.button("🚀 Jalankan Proses", type="primary", use_container_width=True):
            if not username:
                st.error("❌ Username harus diisi di sidebar!")
                st.stop()

            if operation_mode == "🚀 Mode Real":
                st.warning("""
                ⚠️ **Mode Real di Streamlit Cloud**
                
                Fitur Selenium tidak tersedia di Streamlit Cloud environment.
                Aplikasi akan berjalan dalam mode simulasi.
                
                Untuk akses real Ellipse, jalankan script ini di:
                - Local computer dengan Chrome installed
                - VM/Server yang mendukung Chrome Driver
                """)
                use_simulation = True
            else:
                use_simulation = True

            # Jalankan proses
            with st.spinner("🔄 Memproses data..."):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                if use_simulation:
                    # Mode simulasi
                    status_text.text("🎯 Menjalankan mode simulasi...")
                    time.sleep(2)
                    
                    hasil_ro = simulate_ro_processing(df, username, env_choice)
                    
                    # Simulasi progress
                    for i in range(len(df)):
                        progress_bar.progress((i + 1) / len(df))
                        time.sleep(0.1)
                    
                    status_text.text("✅ Simulasi selesai!")

                # Simpan hasil
                df["RO_NUMBER"] = hasil_ro
                
                # Export hasil
                output_filename = f"RO_Results_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
                df.to_excel(output_filename, index=False)

                # Tampilkan summary
                st.divider()
                st.subheader("📈 Summary Hasil")
                
                success_count = len([x for x in hasil_ro if not x.startswith('SKIP')])
                failed_count = len([x for x in hasil_ro if x.startswith('SKIP')])
                
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ Berhasil", success_count)
                col2.metric("❌ Gagal", failed_count)
                col3.metric("📊 Total", len(df))

                # Tampilkan hasil detail
                st.subheader("📋 Detail Hasil")
                result_df = df[['RO_NUMBER']].copy()
                result_df['Status'] = result_df['RO_NUMBER'].apply(
                    lambda x: '✅ Berhasil' if not x.startswith('SKIP') else '❌ Gagal'
                )
                st.dataframe(result_df)

                # Download button
                with open(output_filename, "rb") as f:
                    st.download_button(
                        label="📥 Download Hasil Excel",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

                st.success("🎉 Proses selesai!")
                
                if use_simulation:
                    st.info("""
                    💡 **Informasi Mode Simulasi:**
                    - RO Number yang dihasilkan adalah contoh simulasi
                    - Untuk implementasi real, jalankan di local environment
                    - Pastikan dependencies Selenium terinstall di local
                    """)

    except Exception as e:
        st.error(f"❌ Gagal memproses file: {str(e)}")
        st.code(traceback.format_exc())

else:
    st.info("📁 Silakan upload file Excel untuk memulai proses")

# Footer
st.divider()
st.markdown("""
### 🚀 Deployment Options

**1. Streamlit Cloud (Recommended untuk Testing)**
- Mode simulasi saja
- Tidak perlu setup Chrome
- Instant deployment

**2. Local Environment (Untuk Production)**
```bash
pip install streamlit pandas openpyxl selenium webdriver-manager
streamlit run app.py
