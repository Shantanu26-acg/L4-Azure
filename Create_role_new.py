# from selenium import webdriver
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.select import Select
# from selenium.webdriver.chrome.options import Options
# import time
# import pandas as pd
# import tempfile
#
# options = Options()
# options.add_argument(f'--user-data-dir={tempfile.mkdtemp()}')
# options.add_argument("--headless")
# options.add_argument("--disable-gpu")
# options.add_argument("--no-sandbox")
# options.add_argument("--disable-dev-shm-usage")
#
# driver=webdriver.Chrome(options=options)
# wait=WebDriverWait(driver, 15)
#
# df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="Rights_test_case")
#
# df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")
#
# print(df.columns.tolist())
#
# url=df_logindata.iloc[0, 0]
# username=df_logindata.iloc[0, 1]
# password=df_logindata.iloc[0, 2]
#
# def create_role(driver, wait, role_name, role_desc):
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Menu']"))).click()
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Role Management']"))).click()
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//li[normalize-space()='Add Role']"))).click()
#
#     wait.until(EC.element_to_be_clickable((By.ID, "roleName"))).send_keys(role_name)
#
#     wait.until(EC.element_to_be_clickable((By.NAME, "roleDescription"))).send_keys(role_desc)
#
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='Select All']"))).click()
#
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Submit']"))).click()
#
#     try:
#         toast_element=wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))
#
#         print("Toast Message:", toast_element.text.strip())
#     except:
#         print("No toast element appeared")
#
#     time.sleep(10)
#
# driver.get(url)
# wait.until(EC.presence_of_element_located((By.NAME,'identifier'))).send_keys(username)
# driver.find_element(By.NAME,'password').send_keys(password)
# driver.find_element(By.XPATH, "//button[text()='Login']").click()
# # create_role(driver, wait, "Admin_new", "Testing")
#
# for idx, row in df.iterrows():
#     create_role(driver,
#                 wait,
#                 row["Role name"],
#                 row["Role description"]
#             )














from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
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

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="Rights_test_case")

df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

print(df.columns.tolist())

url=df_logindata.iloc[0, 0]
username=df_logindata.iloc[0, 1]
password=df_logindata.iloc[0, 2]

def create_role(driver, wait, role_name, role_desc):
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Menu']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Role Management']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//li[normalize-space()='Add Role']"))).click()

    wait.until(EC.element_to_be_clickable((By.ID, "roleName"))).send_keys(role_name)

    wait.until(EC.element_to_be_clickable((By.NAME, "roleDescription"))).send_keys(role_desc)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='Select All']"))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Submit']"))).click()

    try:
        toast_element=wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))

        print("Toast Message:", toast_element.text.strip())
    except:
        print("No toast element appeared")

    time.sleep(10)

driver.get(url)
time.sleep(10)
# wait.until(EC.presence_of_element_located((By.NAME,'identifier'))).send_keys(username)
elem=wait.until(EC.visibility_of_element_located((By.NAME,'identifier')))
driver.execute_script("arguments[0].scrollIntoView();", elem)
wait.until(EC.element_to_be_clickable((By.NAME, 'identifier')))
elem.clear()
elem.send_keys(username)
driver.find_element(By.NAME,'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
# create_role(driver, wait, "Admin_new", "Testing")

for idx, row in df.iterrows():
    create_role(driver,
                wait,
                row["Role name"],
                row["Role description"]
            )
