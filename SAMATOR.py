import streamlit as st
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys as keyboard_keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time

st.set_page_config(page_title="SO Automator", page_icon="💊", layout="wide")
st.title("SAMATOR 💊")
st.markdown('Upload your Excel file to begin automation')
st.sidebar.header("Farmacare login")
user_email = st.sidebar.text_input("Email")
user_pass = st.sidebar.text_input("Password", type="password")

file = st.file_uploader("Upload Excel file", type=["xlsx"])

if file:
    xl=pd.ExcelFile(file)
    sheet_names = xl.sheet_names
    sheet = st.selectbox("Select sheet", sheet_names)
    df = pd.read_excel(file, sheet_name=sheet)
    st.write(f"Preview of '{sheet}' sheet:")
    st.dataframe(df.head())

    if st.button("Start Automation"):
        if not user_email or not user_pass:
            st.error("Please enter your credentials first")
        else:
            st.info("Starting automation...")
            options = Options()
            prefs = {"profile.default_content_setting_values.notifications": 2}
            options.add_experimental_option("prefs", prefs)
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-blink-features=AutomationControlled")
            service = Service("/usr/bin/chromedriver")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            wait = WebDriverWait(driver, 30)
            try:
                #logging in
                st.info("Logging in")
                driver.get("https://app.farmacare.id/auth/signin")
                wait.until(EC.presence_of_element_located((By.NAME, "login-email"))).send_keys(user_email)
                driver.find_element(By.NAME, "login-password").send_keys(user_pass)
                driver.find_element(By.ID, "btn-login").click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@data-testid="company_name"]'))).click()
                for i in range(3):
                    try:
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@data-testid="sidebar-menu-/catalog-parent"]'))).click()
                        break
                    except TimeoutException:
                        if i == 2: raise
                        time.sleep(2)
                inventory_url = driver.current_url
                wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Cari barang"]'))).click()
                st.success("Login successful, starting automation process...")            
                #starting loop
                progress_bar = st.progress(0)
                status_text = st.empty()
                for index, row in df.iterrows():
                    status_text.text(f"Processing row {index + 1} of {len(df)}")
                    try:
                        # cari barang
                        search_field = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Cari barang"]')))
                        search_field.send_keys(keyboard_keys.CONTROL + "a")
                        search_field.send_keys(keyboard_keys.BACKSPACE)
                        search_field.send_keys(str(row['Nama Barang']).strip(), keyboard_keys.ENTER)
                        medicine_name = str(row['Nama Barang'])
                        try:
                            wait.until(EC.text_to_be_present_in_element((By.XPATH, "//*[@data-cellid='cell_products-0']"), medicine_name))
                            suggestion = driver.find_element(By.XPATH, "//*[@data-cellid='cell_products-0']")
                            suggestion.click()
                        except TimeoutException:
                            st.warning(f"Row {index + 1}: {medicine_name} - Item not found. Skipping.")
                            driver.get(inventory_url)
                            continue
                        # navigating the stock update page
                        bfr_stock_el = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@data-testid='stok-value']")))
                        bfr_stock_txt = bfr_stock_el.text
                        current_stock = int(''.join(filter(str.isdigit, bfr_stock_txt)))
                        aftr_stock = current_stock + row['Selisih']
                        # handling negative stock scenario
                        if aftr_stock < 0:
                            st.warning(f"Row {index + 1}: {row['Nama Barang']} - Stock negative. Skipping update.")
                            driver.back()
                            wait.until(EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Cari barang"]'))).clear()
                            continue
                        # updating stock
                        else:
                            driver.find_element(By.ID, "text-stock-btn-update-stock").click()
                            stock_input_field = wait.until(EC.presence_of_element_located((By.ID, "st-input-pkg-qty-item-0")))
                            stock_input_field.send_keys(keyboard_keys.CONTROL + "a")
                            stock_input_field.send_keys(keyboard_keys.BACKSPACE)
                            stock_input_field.send_keys(str(aftr_stock))
                            driver.find_element(By.ID, "text-st-btn-save").click()
                            save_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Simpan')]")))
                            save_btn.click()
                            # adding comment
                            comment_field = wait.until(EC.presence_of_element_located((By.ID, "text-box-update-stock-reason")))
                            comment_field.send_keys(row['Keterangan'])
                            wait.until(EC.element_to_be_clickable((By.ID, "text-modal-reason-save-action"))).click()
                            target_stock = str(aftr_stock)
                            wait.until(EC.text_to_be_present_in_element((By.XPATH, "//*[@data-testid='stok-value']"), target_stock))
                            st.success(f"Row {index + 1}: {row['Nama Barang']} - Stock updated successfully to {aftr_stock}.")
                            #navigating back to search
                            driver.get(inventory_url)
                        #back to loop
                    except Exception as e:
                        st.error(f"Error processing row {index + 1}: {e}")
                        continue
            finally:
                driver.quit()
                st.success("Automation completed!")
