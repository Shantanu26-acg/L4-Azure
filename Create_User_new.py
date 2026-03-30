from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd
import tempfile

options = Options()
options.add_argument(f'--user-data-dir={tempfile.mkdtemp()}')
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver=webdriver.Chrome(options=options)
wait=WebDriverWait(driver, 15)

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="Users", dtype={
    "Phone": str
})

df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

url=df_logindata.iloc[0, 0]
username=df_logindata.iloc[0, 1]
password=df_logindata.iloc[0, 2]

print(url, username, password)

def create_user(driver, wait, name, role, email, location, country, phone_number):
    wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()= 'User Management']"))).click()
    driver.find_element(By.XPATH, "//li[text()= 'Add User']").click()

    time.sleep(3)

    wait.until(EC.element_to_be_clickable((By.ID, "userName"))).send_keys(name)

    driver.find_element(By.XPATH, "//span[text()='Select role']").click()

    rol = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{role}']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", rol)
    driver.find_element(By.XPATH, f"//span[text()='{role}']").click()

    driver.find_element(By.ID, 'emailAddress').send_keys(email)

    print(email)

    # print(location)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Select location']"))).click()

    loc=wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{location}']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", loc)
    driver.find_element(By.XPATH, f"//span[text()='{location}']").click()

    driver.find_element(By.XPATH, "//button[@title='India']").click()
    driver.find_element(By.XPATH, f"//span[text()='{country}']").click()

    driver.find_element(By.ID, "custom-phone-input").send_keys(str(phone_number))
    time.sleep(2)

    # driver.find_element(By.XPATH, "//button[normalize-space(text())='Next']").click()

    next_btn=wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Next']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
    driver.execute_script("arguments[0].click();", next_btn)
    # driver.find_element(By.XPATH, "//button[@aria-label='Next']").click()

    time.sleep(3)

    submit_btn=wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Submit']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
    driver.execute_script("arguments[0].click();", submit_btn)

    try:
        toast_element=wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))

        print("Toast Message:", toast_element.text.strip())
    except:
        print("No toast element appeared")

    time.sleep(5)

driver.get(url)
wait.until(EC.element_to_be_clickable((By.NAME,'identifier'))).send_keys(username)
driver.find_element(By.NAME,'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
# # create_user(driver, wait, "abcd", "Admin", "abcd@acg-world.com", "ACG_Phy_0108_2", "India", 9876545678)

for idx, row in df.iterrows():
    create_user(driver,
                wait,
                row['Usern'],
                row['Rolen'],
                row['Emailid'],
                row['Location'],
                row['Country'],
                row['Phone']
                )