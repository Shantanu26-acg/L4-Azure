from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import pandas as pd
import time
import tempfile

options = Options()
options.add_argument(f'--user-data-dir={tempfile.mkdtemp()}')
# options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver=webdriver.Chrome(options=options)
driver.maximize_window()
wait=WebDriverWait(driver, 15)

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="SNG_Template", dtype=str)

df_login_data=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

url=df_login_data.iloc[0, 0]
username=df_login_data.iloc[0, 1]
password=df_login_data.iloc[0, 2]

def create_template(driver, wait, temp_name, SN_Gen, product, encoding, business_part, location, num_type, gen_type, num_len, start_range, pool_size, no_of_levels, packaging_part1, packaging_part2, packaging_part3, packaging_part4):
    for i in range(int(no_of_levels)):
        productpackaging = locals().get(f"packaging_part{i + 1}")
        print("productpackaging", productpackaging)

        if productpackaging.split("-")[-1]!='Pallet':
            wait.until(EC.presence_of_element_located((By.ID, 'templateName'))).send_keys(temp_name+"_"+str(i))

            driver.find_element(By.XPATH, "//span[text()='Serial Number Generator']").click()
            driver.find_element(By.XPATH, "//span[text()='" + str(SN_Gen) + "']").click()

            dropdown = driver.find_element(By.XPATH, "//span[text()='Product Name - Product Code']")
            driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            dropdown.click()
            time.sleep(1)
            option = driver.find_element(By.XPATH, f"//span[text()='{str(product)}']")
            driver.execute_script("arguments[0].scrollIntoView(true);", option)
            option.click()

            dropdown = driver.find_element(By.XPATH, "//span[text()='Select Packing']")
            driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            dropdown.click()
            option = driver.find_element(By.XPATH, f"//span[text()='{str(productpackaging)}']")
            driver.execute_script("arguments[0].scrollIntoView(true);", option)
            option.click()

            el = driver.find_element(By.XPATH, "//span[text()='Select BusinessPartner']")
            driver.execute_script("arguments[0].scrollIntoView(true)", el)

            driver.find_element(By.XPATH, "//span[text()='Select BusinessPartner']").click()
            time.sleep(3)
            driver.find_element(By.XPATH, "//span[text()='" + str(business_part) + "']").click()

            driver.find_element(
                By.XPATH, f"//label[contains(normalize-space(.), '{location.strip()}')]"
            ).click()

            driver.find_element(By.XPATH, "//span[text()='Add']").click()

            driver.find_element(By.XPATH, "//button[text()='Next']").click()

            time.sleep(2)

            wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='NumberType']/span[contains(text(),'Select')]"))).click()
            driver.find_element(By.XPATH, "//span[text()='" + str(num_type) + "']").click()

            if num_type == 'Numeric':
                driver.find_element(By.XPATH, "//div[@id='generationType']/span[contains(text(),'Select')]").click()
                driver.find_element(By.XPATH, "//span[text()='" + str(gen_type) + "']").click()

            driver.find_element(By.ID, "length").send_keys(str(num_len))

            driver.find_element(By.XPATH, "//button[text()='Next']").click()

            time.sleep(3)

            # pool details
            wait.until(EC.presence_of_element_located((By.NAME, "initialPoolSize"))).send_keys(str(pool_size))

            # click on submit
            driver.find_element(By.XPATH, "//button[text()='Submit']").click()
            message = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='p-toast-detail']"))).get_attribute("innerText")
            print(message)

            time.sleep(3)

            if "successfully" in message or "something" in message:
                print("Done for template", str(productpackaging))
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='+ New Template']"))).click()
            elif "Duplicate" in message:
                driver.find_element(By.XPATH, "//button[text()='Cancel']").click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='+ New Template']"))).click()
            else:
                driver.find_element(By.XPATH, "//button[text()='Cancel']").click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='+ New Template']"))).click()
        else:
            pass

driver.get(url)
wait.until(EC.presence_of_element_located((By.NAME, 'identifier'))).send_keys(username)
driver.find_element(By.NAME, 'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()='Serial Number Management']"))).click()
wait.until(EC.visibility_of_element_located((By.XPATH, "//li[text()='Create Template']"))).click()

# create_template(driver, wait,"abcd", "ACG LifeSciences Cloud", "Methylphenidate HCl ER Tb 27mg-Product055", "URN", "Amneal_SRB", "Amneal_SRB(Office) - 0365162000009.0", "Numeric", "Sequential", 8, 10000, 500000, 1, "00365162233094-Primary", "", "", "")

for idx, row in df.iterrows():
    create_template(
        driver,
        wait,
        row['Temp_name'],
        row['SN_Gen'],
        row['Product'],
        row['encoding type-download'],
        row['Business part'],
        row['Location'],
        row['Number type'],
        row['Gen Type'],
        row['Num_len'],
        row['Start Range'],
        row['pool size'],
        row['No.of levels'],
        row['Packaging level - 1'],
        row['Packaging level - 2'],
        row['Packaging level - 3'],
        row['Packaging level - 4']
    )
