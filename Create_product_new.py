from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
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

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="Product", dtype=str)

def create_product(driver, wait, row_idx, prd_id, prd_name, prd_desc, manufacturer, no_of_levels, item_ref1, packaging_level1, level_indicator1, item_ref2, packaging_level2, level_indicator2, item_ref3, packaging_level3, level_indicator3, item_ref4, packaging_level4, level_indicator4, qty_level1, qty_level2, qty_level3):
    wait.until(EC.presence_of_element_located((By.ID, 'productIdentifierType'))).click()
    checkbox=driver.find_element(By.XPATH, "//span[text()='Product Code']")
    driver.execute_script("arguments[0].scrollIntoView(true)", checkbox)
    driver.find_element(By.XPATH, "//span[text()='Product Code']").click()

    driver.find_element(By.ID, 'productIdentifier').send_keys(prd_id)

    driver.find_element(By.ID, 'productName').send_keys(prd_name)

    driver.find_element(By.ID, "productDescription").send_keys(prd_desc)

    driver.find_element(By.ID, 'manufacturer').click()
    manu=wait.until(EC.presence_of_element_located((By.XPATH, f"//li[@role='option']//span[text()='{manufacturer}']")))
    driver.execute_script("arguments[0].scrollIntoView(true)", manu)
    manu.click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Next']"))).click()

    for i in range(int(no_of_levels)):
        packaging_code=wait.until(EC.presence_of_element_located((By.ID, 'packagingCodeType')))
        driver.execute_script("arguments[0].scrollIntoView(true)", packaging_code)
        packaging_code.click()
        gtin_format = driver.find_element(By.XPATH, "//*[@id='dropdownItem_0']/span")
        driver.execute_script("arguments[0].scrollIntoView(true)", gtin_format)
        driver.find_element(By.XPATH, "//*[@id='dropdownItem_0']/span").click()

        itemref=locals().get(f"item_ref{i+1}")
        packaginglevel=locals().get(f"packaging_level{i+1}")
        levelindicator=locals().get(f"level_indicator{i+1}")

        print(itemref)

        driver.find_element(By.ID, 'packagingProductIdentifier').send_keys(str(itemref))

        driver.find_element(By.ID, 'packagingLevel').click()
        level=wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[@role='option']//span[text()='{packaginglevel}']")))
        driver.execute_script("arguments[0].scrollIntoView(true)", level)
        level.click()

        driver.find_element(By.ID, 'packagingLevelIndicator').send_keys(levelindicator)

        if i+1!=1:
            driver.find_element(By.ID, 'aggregationYes').click()

            childlevel=locals().get(f"packaging_level{i}")
            print(childlevel)
            qtylevel=locals().get(f"qty_level{i}")
            print(qtylevel)

            driver.find_element(By.ID, 'childPackagingLevel_0').click()
            child_level = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[@role='option']//span[starts-with(text(), '{childlevel}')]")))
            driver.execute_script("arguments[0].scrollIntoView(true)", child_level)
            child_level.click()

            driver.find_element(By.ID, 'maxQuantity_0').send_keys(qtylevel)
            gtin = driver.find_element(By.XPATH,"//input[@placeholder='LevelIndicator + GCP + Item Reference + Check Digit']").get_attribute("value")
            print("gtin", gtin)

            df.loc[row_idx, f"GTIN{i + 1}"] = gtin

            df.to_excel("ACG_Common_Workbook.xlsx", sheet_name="Product", index=False)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Add']"))).click()
        elif i+1==1:
            gtin = driver.find_element(By.XPATH,"//input[@placeholder='LevelIndicator + GCP + Item Reference + Check Digit']").get_attribute("value")
            print("gtin", gtin)

            df.loc[row_idx, f"GTIN{i + 1}"] = gtin

            df.to_excel("ACG_Common_Workbook.xlsx", sheet_name="Product", index=False)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Add']"))).click()

        time.sleep(2)

        # level = locals().get(f"PackagingLevel{i + 1}")
        # level_indicator = locals().get(f"LevelIndicator{i + 1}")
        # quantity = locals().get(f"Quantity{i + 1}")
        # print(level, level_indicator, quantity)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Submit']"))).click()

    message = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='p-toast-detail']"))).get_attribute("innerText")

    if "successfully" in message:
        elm = EC.element_to_be_clickable((By.XPATH,"//span[text()= '+ New Product']"))
        WebDriverWait(driver,3).until(elm)
    else:
        driver.find_element(By.XPATH,"//span[text()='Cancel']").click()

    wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='+ New Product']"))).click()

    time.sleep(2)


driver.get("https://lifesciences-sit.acgi.in/")
wait.until(EC.presence_of_element_located((By.NAME, 'identifier'))).send_keys("mukesh.tiwari04@hotmail.com")
driver.find_element(By.NAME, 'password').send_keys("Acgi@12345")
driver.find_element(By.XPATH, "//button[text()='Login']").click()
time.sleep(5)
wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()= 'Master Data']"))).click()
driver.find_element(By.XPATH, "//li[text()= 'Add Product']").click()

# create_product(driver, wait, 11111, "abc", "xyz", "Amneal_SRB", 4, 11111, "Primary", 5, 11111, "Bundle", 8, 11111, "Shipper", 9, 11111, "Pallet", 4, 24, 4, 2)


for idx, row in df.iterrows():
    create_product(
        driver,
        wait,
        idx,
        row['ProductIdentifier'],
        row['ProductName'],
        row['Product Description'],
        row['Manufacturer'],
        row['No. of levels'],
        row['Item_ref1'],
        row['Packaging unit1'],
        row['Indicator1'],
        row['Item_ref2'],
        row['Packaging unit2'],
        row['Indicator2'],
        row['Item_ref3'],
        row['Packaging unit3'],
        row['Indicator3'],
        row['Item_ref4'],
        row['Packaging unit4'],
        row['Indicator4'],
        row['Qtylevel1'],
        row['Qtylevel2'],
        row['Qtylevel3']
    )