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
# options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver=webdriver.Chrome(options=options)
wait=WebDriverWait(driver, 15)

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="partner", dtype={
    'LocationIdentifier': str,
    'Postalcode': str,
    'phnumber': str,
    'GLN': str,
    'Location Number': str
})

df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

url=df_logindata.iloc[0, 0]
username=df_logindata.iloc[0, 1]
password=df_logindata.iloc[0, 2]

print(url, username, password)

def create_location(driver, wait, loc_name, loc_type, entity, loc_id_type, loc_id, bus_ent, phy_site, loc_num, country, state, city, address, postal_code, poc_name, poc_email, phone_number, website):
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Menu']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()= 'Master Data']"))).click()
    driver.find_element(By.XPATH, "//li[text()= 'Location master data']").click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()= '+ New Location']"))).click()

    wait.until(EC.presence_of_element_located((By.ID, 'locationName'))).send_keys(loc_name)

    print(loc_name)

    driver.find_element(By.XPATH, "//span[text()='Select location type']").click()
    driver.find_element(By.XPATH, f"//span[text()='{loc_type}']").click()

    time.sleep(5)

    driver.find_element(By.XPATH, "//span[text()='Select entity type']").click()
    el_e = driver.find_element(By.XPATH, f"//span[text()='{entity}']")
    driver.execute_script("arguments[0].scrollIntoView(true);", el_e)
    driver.find_element(By.XPATH, f"//span[text()='{entity}']").click()

    if loc_type=='Business Partner':
        driver.find_element(By.XPATH, "//input[@placeholder='Enter location identifier']").send_keys(loc_id)
        if len(loc_id)<12:
            driver.find_element(By.ID, 'locationNumber').send_keys(loc_num)
            # print(loc_num)
    if loc_type=="Physical Site":
        driver.find_element(By.XPATH, "//span[text()='Select business entity']").click()
        driver.find_element(By.XPATH, f"//span[text()='{bus_ent}']").click()

        driver.find_element(By.XPATH, "//span[text()='Identifier type']").click()
        driver.find_element(By.XPATH, f"//span[text()='{loc_id_type}']").click()
    if loc_type=='Internal Site':
        driver.find_element(By.XPATH, "//span[text()='Select physical site']").click()
        driver.find_element(By.XPATH, f"//span[text()='{phy_site}']").click()

    time.sleep(2)

    driver.find_element(By.XPATH, "//button[normalize-space(text())='Next']").click()

    time.sleep(5)

    driver.find_element(By.XPATH, "//span[text()='Select country']").click()
    time.sleep(2)
    driver.find_element(By.XPATH, f"(//span[text()='{country}'])[2]").click()

    driver.find_element(By.ID, "state").send_keys(state)

    driver.find_element(By.ID, 'city').send_keys(city)

    driver.find_element(By.ID, 'address').send_keys(address)

    driver.find_element(By.ID, 'postalCode').send_keys(postal_code)

    driver.find_element(By.ID, 'contactPersonName').send_keys(poc_name)

    driver.find_element(By.ID, 'contactEmailId').send_keys(poc_email)

    driver.find_element(By.XPATH, "//button[@title='India']").click()

    element = driver.find_element(By.XPATH, "//span[text()='" + str(country) + "']")
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    driver.find_element(By.XPATH, f"//span[text()='{country}']").click()

    driver.find_element(By.ID, "custom-phone-input").send_keys(str(phone_number))

    driver.find_element(By.ID, 'website').send_keys(website)

    time.sleep(2)

    driver.find_element(By.XPATH, "//button[normalize-space(text())='Submit']").click()

    try:
        toast_element=wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))

        print("Toast Message:", toast_element.text.strip())
    except:
        print("No toast message appeared")

    time.sleep(5)


driver.get(url)
wait.until(EC.element_to_be_clickable((By.NAME,'identifier'))).send_keys(username)
driver.find_element(By.NAME,'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
# create_location(driver, wait, "abc", "Internal Site", "Bin", "GLN/SGLN", 100001, "Macleods Pharmaceuticals Ltd", "Amneal_SRB(Office)", 1111111, "India","Maharashtra" ,"Mumbai" ,"abcd" ,"400067" ,"Shantanu","shantanu.mule@email.com", "9867766319", "https://shantanu.com")

# create_location(driver, wait, "abc", "Business Partner", "Manufacturer", "GLN/SGLN", 987678, "Macleods Pharmaceuticals Ltd", "Amneal_SRB(Office)", 1111111, "India","Maharashtra" ,"Mumbai" ,"abcd" ,"400067" ,"Shantanu","shantanu.mule@email.com", "9867766319", "https://shantanu.com")

for idx, row in df.iterrows():
    create_location(driver,
                    wait,
                    row['LocationName'],
                    row['Location Type'],
                    row['Entity'],
                    row['Location ID Type'],
                    row['LocationIdentifier'],
                    row['Business Entity'],
                    row['Physical Site'],
                    row['Location Number'],
                    row['Country'],
                    row['State'],
                    row['City'],
                    row['Address'],
                    row['Postalcode'],
                    row['Contactperson'],
                    row['email'],
                    row['phnumber'],
                    row['website']
                    )
