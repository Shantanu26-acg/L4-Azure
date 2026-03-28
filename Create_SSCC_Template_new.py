from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from openpyxl import Workbook,load_workbook
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import tempfile
import pandas as pd

options = Options()
options.add_argument(f'--user-data-dir={tempfile.mkdtemp()}')
# options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver=webdriver.Chrome(options=options)
wait=WebDriverWait(driver, 15)

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="SSCC_template", dtype=str)

df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

url=df_logindata.iloc[0, 0]
username=df_logindata.iloc[0, 1]
password=df_logindata.iloc[0, 2]

def create_sscc_template(driver, wait, template_name, business_partner, number_generator, extension_digit, bus_ptn, location):
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Menu']"))).click()
    element=wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Serial Number Management']")))
    driver.execute_script("arguments[0].scrollIntoView();", element)
    driver.execute_script("arguments[0].click();", element)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[normalize-space()='Create SSCC Template']"))).click()

    wait.until(EC.element_to_be_clickable((By.ID, "templateName"))).send_keys(template_name)

    print(template_name)

    wait.until(EC.element_to_be_clickable((By.ID, "businessPartnerGCP"))).click()
    business_partner_option=wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space()='{business_partner}']")))
    driver.execute_script("arguments[0].click();", business_partner_option)

    wait.until(EC.element_to_be_clickable((By.ID, "serialNumberGenerator"))).click()
    num_gen_opt=wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space()='{number_generator}']")))
    driver.execute_script("arguments[0].click();", num_gen_opt)

    wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[normalize-space()='{extension_digit}']"))).click()

    wait.until(EC.element_to_be_clickable((By.ID, "businessPartnerName"))).click()
    bus_ptn_opt=wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[normalize-space()='{bus_ptn}']")))
    driver.execute_script("arguments[0].click();", bus_ptn_opt)

    wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[normalize-space()='{location}']"))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Add']"))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Submit']"))).click()

    try:
        toast_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))

        print("Toast message:", toast_element.text.strip())
    except:
        print("No toast message appeared")

    time.sleep(5)

    driver.refresh()

driver.get(url)
wait.until(EC.element_to_be_clickable((By.NAME,'identifier'))).send_keys(username)
driver.find_element(By.NAME,'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
# create_sscc_template(driver, wait, "", "", "", "", "", "")

for idx, row in df.iterrows():
    create_sscc_template(driver,
                         wait,
                         template_name=row["Temp_name"],
                         business_partner=row["Business part"],
                         number_generator=row["SN_Gen"],
                         extension_digit=row["extension"],
                         bus_ptn=row["Business part1"],
                         location=row["Location"]
                         )
