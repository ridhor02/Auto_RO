import streamlit as st
import pandas as pd
import time
import os
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import chromedriver_autoinstaller

# ==============================
# KONFIGURASI
# ==============================
URLS = {
    "🌐 Production": "http://ellipse-ellprd.bssr-ellprd.baramultigroup.co.id/html/ui#!/ome",
    "🧪 Testing": "http://ellipse-elltst.bssr-elldev.baramultigroup.co.id/html/ui#!/ome"
}
QUICK_LAUNCH_CODE = "MSO240"
SELENIUM_WAIT = 20
MAX_RETRIES = 3

# ==============================
# SETUP UNTUK STREAMLIT CLOUD
# ==============================
def setup_driver_for_streamlit():
    """Setup ChromeDriver khusus untuk Streamlit Cloud"""
    try:
        # Install ChromeDriver dengan opsi yang lebih kompatibel
        chromedriver_autoinstaller.install()
        
        # Setup Chrome options untuk Streamlit Cloud
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-features=VizDisplayCompositor")
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-software-rasterizer")
        
        # Untuk bypass automation detection
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Gunakan Chrome yang sudah terinstall di system
        driver = webdriver.Chrome(options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        return driver
        
    except Exception as e:
        st.error(f"❌ Gagal setup ChromeDriver: {e}")
        
        # Fallback: coba tanpa chromedriver_autoinstaller
        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except Exception as fallback_error:
            st.error(f"❌ Fallback juga gagal: {fallback_error}")
            raise

# ==============================
# HELPER FUNCTIONS
# ==============================
def retry_element_interaction(wait, by, value, action_func, max_retries=MAX_RETRIES):
    """Retry mechanism untuk interaksi elemen"""
    for attempt in range(max_retries):
        try:
            element = wait.until(EC.element_to_be_clickable((by, value)))
            action_func(element)
            return True
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
            time.sleep(1)
    return False

def safe_fill_field(driver, wait, field_id, value, use_dropdown=False):
    """Mengisi field dengan aman dan retry"""
    try:
        if use_dropdown:
            dropdown = wait.until(EC.element_to_be_clickable((By.ID, field_id)))
            dropdown.click()
            time.sleep(0.5)
            active_elem = driver.switch_to.active_element
            active_elem.send_keys(value)
            active_elem.send_keys(Keys.ENTER)
        else:
            field = wait.until(EC.element_to_be_clickable((By.ID, field_id)))
            field.clear()
            field.send_keys(value)
            field.send_keys(Keys.ENTER)
        return True, None
    except Exception as e:
        return False, str(e)

def get_ro_number_from_attribute(driver, wait, element_id="RO_NO1I_3", max_retries=5):
    """
    Ambil nilai RO dari atribut 'data-value' karena field readonly & disabled.
    RO Number ter-generate otomatis setelah Stock Code + Quantity diisi.
    """
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            driver.switch_to.default_content()
            st.write(f"🔎 Percobaan {attempt}: mencari elemen {element_id}")

            # Cari frame yang berisi elemen RO
            found = False
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            for i, frame in enumerate(frames):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(frame)
                    if driver.find_elements(By.ID, element_id):
                        found = True
                        st.write(f"✅ Elemen ditemukan di frame ke-{i}")
                        break
                except Exception:
                    continue

            if not found:
                driver.switch_to.default_content()
                st.warning("⚠️ Elemen tidak ditemukan di semua frame")

            elem = wait.until(EC.presence_of_element_located((By.ID, element_id)))
            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
            time.sleep(1)

            # Coba berbagai atribut untuk mendapatkan nilai
            value = (elem.get_attribute("data-value") or 
                    elem.get_attribute("value") or 
                    elem.text or
                    elem.get_attribute("innerText"))

            if value and value.strip() and value.strip().lower() != "none":
                cleaned_value = value.strip()
                st.success(f"✅ RO Number ditemukan: {cleaned_value}")
                return cleaned_value, None

            last_err = f"Percobaan {attempt}: data-value kosong"
            st.warning(last_err)
            time.sleep(2)
            continue

        except Exception as e:
            last_err = f"Percobaan {attempt} error: {e}"
            st.warning(f"⚠️ {last_err}")
            time.sleep(2)
            continue

    return None, last_err or "Gagal ambil RO Number"

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

def switch_to_correct_frame(driver, element_id):
    """Switch ke frame yang mengandung element tertentu"""
    driver.switch_to.default_content()
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    
    for frame in frames:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            if driver.find_elements(By.ID, element_id):
                return True
        except Exception:
            continue
    
    driver.switch_to.default_content()
    return False

# ==============================
# STREAMLIT UI
# ==============================
st.set_page_config(
    page_title="Otomasi Input RO Ellipse",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Otomasi Input RO Ellipse (MSO240)")
st.caption("Upload file Excel RO lalu klik Jalankan Otomasi")

# Informasi untuk Streamlit Cloud
with st.expander("ℹ️ Informasi Aplikasi"):
    st.info("""
    **Aplikasi Otomasi Input RO Ellipse**
    
    🔧 **Fitur:**
    - Auto input data RO ke sistem Ellipse
    - Support environment Production & Testing
    - Generate RO Number otomatis
    - Export hasil ke Excel
    
    ⚠️ **Catatan:**
    - Pastikan file Excel memiliki minimal 8 kolom
    - Koneksi internet stabil diperlukan
    - Proses mungkin memakan waktu beberapa menit
    - Pastikan credentials login valid
    """)

with st.sidebar:
    st.header("🌍 Pilih Environment")
    env_choice = st.radio("Pilih server Ellipse:", list(URLS.keys()), index=1)
    TARGET_URL = URLS[env_choice]
    st.info(f"🔗 **URL aktif:** {TARGET_URL}")

    st.divider()
    st.header("🔐 Login Credentials")
    username = st.text_input("Username", value="", placeholder="0192210XXX")
    password = st.text_input("Password", value="", type="password", placeholder="••••••")

    st.divider()
    st.header("⚙️ Settings")
    headless_mode = st.checkbox("Headless Mode", value=True, help="Browser berjalan di background")
    auto_download = st.checkbox("Auto download hasil", value=True)
    retry_on_error = st.checkbox("Retry otomatis jika gagal", value=True)

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
        **Kolom yang akan diproses:**
        - Kolom 2 (index 1): Stock Code
        - Kolom 5 (index 4): Quantity  
        - Kolom 8 (index 7): Warehouse
        - Kolom 9 (index 8): RO Number hasil
        """)

        if st.button("🚀 Jalankan Otomasi", type="primary", use_container_width=True):
            if not username or not password:
                st.error("❌ Username dan Password harus diisi di sidebar!")
                st.stop()

            # Warning untuk proses lama
            st.warning("⏳ Proses mungkin memakan waktu beberapa menit. Jangan tutup browser!")
            st.info("💡 Pastikan koneksi internet stabil")

            driver = None
            try:
                # Setup WebDriver untuk Streamlit Cloud
                with st.spinner("🔄 Menyiapkan browser..."):
                    driver = setup_driver_for_streamlit()
                    wait = WebDriverWait(driver, SELENIUM_WAIT)
                st.success("✅ Browser siap")

                # Login
                with st.spinner(f"🔐 Login ke {env_choice}..."):
                    driver.get(TARGET_URL)
                    time.sleep(3)
                    
                    # Switch to login frame jika ada
                    switch_to_correct_frame(driver, "username")
                    
                    username_field = wait.until(EC.element_to_be_clickable((By.ID, "username")))
                    username_field.send_keys(username)
                    
                    password_field = driver.find_element(By.ID, "password")
                    password_field.send_keys(password)
                    
                    login_button = driver.find_element(By.ID, "login")
                    login_button.click()
                    time.sleep(5)
                st.success("✅ Login berhasil")

                # Buka MSO240
                with st.spinner("📋 Membuka MSO240..."):
                    driver.switch_to.default_content()
                    quick = wait.until(EC.element_to_be_clickable((By.ID, "quicklaunch")))
                    quick.send_keys(QUICK_LAUNCH_CODE + Keys.ENTER)
                    time.sleep(8)  # Beri waktu lebih untuk loading
                st.success("✅ MSO240 terbuka")

                # Inisialisasi progress
                progress_bar = st.progress(0)
                status_text = st.empty()
                success_count = 0
                failed_count = 0
                hasil_ro = []

                # Process each row
                for idx in range(len(df)):
                    stock = str(df.iloc[idx, 1]).strip()
                    qty = str(df.iloc[idx, 4]).strip()
                    wh = str(df.iloc[idx, 7]).strip()

                    status_text.text(f"📝 Memproses baris {idx+1}/{len(df)}: {stock}")

                    if not stock or stock.lower() == "nan":
                        st.warning(f"⏭️ Baris {idx+1}: Stock code kosong, dilewati")
                        hasil_ro.append("SKIP - Stock code kosong")
                        failed_count += 1
                        progress_bar.progress((idx + 1) / len(df))
                        continue

                    try:
                        # Switch ke frame yang benar
                        driver.switch_to.default_content()
                        switch_to_correct_frame(driver, "STOCK_CODE1I_3")

                        # === Tahap 1: Isi Stock Code ===
                        success, error = safe_fill_field(driver, wait, "STOCK_CODE1I_3", stock)
                        if not success:
                            raise Exception(f"Gagal isi Stock Code: {error}")
                        time.sleep(3)  # Beri waktu untuk generate RO

                        # === Tahap 2: Ambil RO Number ===
                        ro_no, error_msg = get_ro_number_from_attribute(driver, wait, "RO_NO1I_3", max_retries=3)
                        if ro_no:
                            # === Tahap 3: Isi Quantity jika RO berhasil ===
                            qty_field = wait.until(EC.element_to_be_clickable((By.ID, "QTY_UOI1I_25")))
                            qty_field.clear()
                            qty_field.send_keys(qty)
                            qty_field.send_keys(Keys.TAB)
                            time.sleep(2)

                            # === Tahap 4: Isi Warehouse jika ada ===
                            if wh and wh.lower() != "nan":
                                try:
                                    wh_success, wh_error = safe_fill_field(driver, wait, "WHSE1I_4", wh)
                                    if not wh_success:
                                        st.warning(f"⚠️ Gagal isi warehouse: {wh_error}")
                                except Exception as wh_e:
                                    st.warning(f"⚠️ Error isi warehouse: {wh_e}")

                            # === Tahap 5: Isi Field Tambahan ===
                            try:
                                vat_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "VAT_CODE1I_39_dropdown")))
                                vat_dropdown.click()
                                time.sleep(0.5)
                                active_elem = driver.switch_to.active_element
                                active_elem.send_keys("00")
                                active_elem.send_keys(Keys.ENTER)
                                time.sleep(1)
                            except Exception as vat_e:
                                st.warning(f"⚠️ Skip VAT code: {vat_e}")

                            try:
                                proc_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "PROCESS_CODE1I_8_dropdown")))
                                proc_dropdown.click()
                                time.sleep(0.3)
                                driver.switch_to.active_element.send_keys("S")
                                driver.switch_to.active_element.send_keys(Keys.ENTER)
                                time.sleep(1)
                            except Exception as proc_e:
                                st.warning(f"⚠️ Skip process code: {proc_e}")

                            # Simpan RO
                            try:
                                save_btn = wait.until(EC.element_to_be_clickable((By.ID, "toolbar-2")))
                                save_btn.click()
                                time.sleep(3)
                                st.success(f"✅ Baris {idx+1}: RO {ro_no} berhasil disimpan")
                            except Exception as save_e:
                                st.warning(f"⚠️ Gagal save RO: {save_e}")

                            hasil_ro.append(ro_no)
                            success_count += 1
                        else:
                            hasil_ro.append(f"GAGAL - {error_msg}")
                            failed_count += 1
                            st.error(f"❌ Baris {idx+1}: {error_msg}")

                    except Exception as e:
                        error_msg = f"Error: {str(e)}"
                        hasil_ro.append(error_msg)
                        failed_count += 1
                        st.error(f"❌ Baris {idx+1}: {error_msg}")

                    progress_bar.progress((idx + 1) / len(df))

                # Simpan hasil
                if len(hasil_ro) == len(df):
                    df["RO_NUMBER"] = hasil_ro
                else:
                    # Pad hasil_ro jika panjangnya tidak sama
                    while len(hasil_ro) < len(df):
                        hasil_ro.append("ERROR - Process incomplete")
                    df["RO_NUMBER"] = hasil_ro

                # Export hasil
                output_filename = f"RO_Results_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
                df.to_excel(output_filename, index=False)

                # Tampilkan summary
                st.divider()
                st.subheader("📈 Summary Hasil")
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ Berhasil", success_count)
                col2.metric("❌ Gagal", failed_count)
                col3.metric("📊 Total", len(df))

                # Tampilkan hasil detail
                st.subheader("📋 Detail Hasil")
                result_df = df[['RO_NUMBER']].copy()
                result_df['Status'] = result_df['RO_NUMBER'].apply(
                    lambda x: '✅ Berhasil' if not x.startswith(('SKIP', 'GAGAL', 'ERROR')) else '❌ Gagal'
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

            except Exception as e:
                st.error(f"❌ Error Fatal: {str(e)}")
                st.code(traceback.format_exc())
            finally:
                if driver:
                    try:
                        driver.quit()
                        st.info("🔒 Browser ditutup")
                    except:
                        pass

    except Exception as e:
        st.error(f"❌ Gagal memproses file: {str(e)}")
        st.code(traceback.format_exc())

else:
    st.info("📁 Silakan upload file Excel untuk memulai otomasi")

# Footer
st.divider()
st.caption("🤖 Otomasi RO Ellipse v2.0 | Streamlit Cloud Ready")
