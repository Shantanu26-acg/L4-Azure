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
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver=webdriver.Chrome(options=options)
wait=WebDriverWait(driver, 15)

df=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="Product+SNG", dtype=str)

df_logindata=pd.read_excel("ACG_Common_Workbook.xlsx", sheet_name="URL_Login_cred_Tenant")

url=df_logindata.iloc[0, 0]
username=df_logindata.iloc[0, 1]
password=df_logindata.iloc[0, 2]

def create_product(driver, wait, row_idx, prd_id, prd_name, prd_desc, manufacturer, no_of_levels, item_ref1, packaging_level1, level_indicator1, item_ref2, packaging_level2, level_indicator2, item_ref3, packaging_level3, level_indicator3, item_ref4, packaging_level4, level_indicator4, qty_level1, qty_level2, qty_level3):
    time.sleep(5)
    wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
    wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()= 'Master Data']"))).click()
    driver.find_element(By.XPATH, "//li[text()= 'Add Product']").click()
    wait.until(EC.presence_of_element_located((By.ID, 'productIdentifierType'))).click()
    checkbox=driver.find_element(By.XPATH, "//span[text()='Product Code']")
    driver.execute_script("arguments[0].scrollIntoView(true)", checkbox)
    driver.find_element(By.XPATH, "//span[text()='Product Code']").click()

    driver.find_element(By.ID, 'productIdentifier').send_keys(prd_id)

    driver.find_element(By.ID, 'productName').send_keys(prd_name)

    driver.find_element(By.ID, "productDescription").send_keys(prd_desc)

    driver.find_element(By.ID, 'manufacturer').click()
    manu=wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[@role='option']//span[text()='{manufacturer}']")))
    driver.execute_script("arguments[0].scrollIntoView(true)", manu)
    driver.execute_script("arguments[0].click()", manu)
    # manu.click()

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

        # print(itemref)

        driver.find_element(By.ID, 'packagingProductIdentifier').send_keys(str(itemref))

        driver.find_element(By.ID, 'packagingLevel').click()
        level=wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[@role='option']//span[text()='{packaginglevel}']")))
        driver.execute_script("arguments[0].scrollIntoView(true)", level)
        level.click()

        driver.find_element(By.ID, 'packagingLevelIndicator').send_keys(levelindicator)

        if i+1!=1:
            driver.find_element(By.ID, 'aggregationYes').click()

            childlevel=locals().get(f"packaging_level{i}")
            # print(childlevel)
            qtylevel=locals().get(f"qty_level{i}")
            # print(qtylevel)

            driver.find_element(By.ID, 'childPackagingLevel_0').click()
            child_level = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[@role='option']//span[starts-with(text(), '{childlevel}')]")))
            driver.execute_script("arguments[0].scrollIntoView(true)", child_level)
            child_level.click()

            driver.find_element(By.ID, 'maxQuantity_0').send_keys(qtylevel)
            gtin = driver.find_element(By.XPATH,"//input[@placeholder='LevelIndicator + GCP + Item Reference + Check Digit']").get_attribute("value")
            # print("gtin", gtin)

            df.loc[row_idx, f"GTIN{i + 1}"] = gtin

            # df.to_excel("ACG_Common_Workbook.xlsx", sheet_name="Product", index=False)
            with pd.ExcelWriter("ACG_Common_Workbook.xlsx", mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Product+SNG", index=False)

            productpack = gtin + "-" + packaginglevel
            # print(productpack)

            df.loc[row_idx, f"Packaging level - {i + 1}"] = productpack

            with pd.ExcelWriter("ACG_Common_Workbook.xlsx", mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Product+SNG", index=False)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Add']"))).click()
        elif i+1==1:
            gtin = driver.find_element(By.XPATH,"//input[@placeholder='LevelIndicator + GCP + Item Reference + Check Digit']").get_attribute("value")
            # print("gtin", gtin)

            df.loc[row_idx, f"GTIN{i + 1}"] = gtin

            with pd.ExcelWriter("ACG_Common_Workbook.xlsx", mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Product+SNG", index=False)

            productpack=gtin+"-"+packaginglevel
            # print(productpack)

            df.loc[row_idx, f"Packaging level - {i + 1}"] = productpack

            with pd.ExcelWriter("ACG_Common_Workbook.xlsx", mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Product+SNG", index=False)

            # df.to_excel("ACG_Common_Workbook.xlsx", sheet_name="Product+SNG", index=False)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Add']"))).click()

        df.loc[row_idx, f"Product"] = prd_name+"-"+prd_id
        # print(prd_name+"-"+prd_id)

        with pd.ExcelWriter("ACG_Common_Workbook.xlsx", mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Product+SNG", index=False)

        time.sleep(2)

        # level = locals().get(f"PackagingLevel{i + 1}")
        # level_indicator = locals().get(f"LevelIndicator{i + 1}")
        # quantity = locals().get(f"Quantity{i + 1}")
        # print(level, level_indicator, quantity)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Submit']"))).click()

    try:
        toast_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-toast-detail')]")))

        print("Toast Message:", toast_element.text.strip())
    except:
        print("No toast message appeared")

    # message = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='p-toast-detail']"))).get_attribute("innerText")
    #
    # if "successfully" in message:
    #     print("Product Created Successfully")

    time.sleep(2)

def create_template(driver, wait, temp_name, SN_Gen, product, encoding, business_part, location, num_type, gen_type, num_len, pool_size, no_of_levels, packaging_part1, packaging_part2, packaging_part3, packaging_part4):
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
            option_prod = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{str(product)}']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", option_prod)
            driver.execute_script("arguments[0].click();", option_prod)
            # option.click()

            dropdown = driver.find_element(By.XPATH, "//span[text()='Select Packing']")
            driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            dropdown.click()
            option = driver.find_element(By.XPATH, f"//span[text()='{str(productpackaging)}']")
            driver.execute_script("arguments[0].scrollIntoView(true);", option)
            driver.execute_script("arguments[0].click();", option)
            # option.click()

            el = driver.find_element(By.XPATH, "//span[text()='Select BusinessPartner']")
            driver.execute_script("arguments[0].scrollIntoView(true)", el)

            driver.find_element(By.XPATH, "//span[text()='Select BusinessPartner']").click()
            time.sleep(3)
            driver.find_element(By.XPATH, "//span[text()='" + str(business_part) + "']").click()

            time.sleep(5)

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


        # if i+1<int(no_of_levels):
        #     wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='+ New Template']"))).click()



def create_product_sng():
    for idx, row in df.iterrows():
        create_product(
            driver,
            wait,
            idx,
            # row['GCP'],
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

        wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
        element=wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()='Serial Number Management']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        driver.execute_script("arguments[0].click();", element)
        wait.until(EC.visibility_of_element_located((By.XPATH, "//li[text()='Create Template']"))).click()

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
            row['pool size'],
            row['No.of levels'],
            row['Packaging level - 1'],
            row['Packaging level - 2'],
            row['Packaging level - 3'],
            row['Packaging level - 4']
        )


driver.get(url)
time.sleep(10)
wait.until(EC.presence_of_element_located((By.NAME, 'identifier'))).send_keys(username)
driver.find_element(By.NAME, 'password').send_keys(password)
driver.find_element(By.XPATH, "//button[text()='Login']").click()
# wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Menu']"))).click()
# wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()='Serial Number Management']"))).click()
# wait.until(EC.visibility_of_element_located((By.XPATH, "//li[text()='Create Template']"))).click()


# create_template(driver, wait, "ACG45364787", "ACG LifeSciences Cloud", "Raloxifene Hydrochloride Tablets\\, USP 60mg-Product030", "URN", "Amneal_SRB", "Amneal_SRB(Office) - 0365162000009.0", "Alphanumeric", "Random", 8, 500000, 3, "00365162157109-Primary", "50365162157104-Shipper", "80365162157105-Pallet", "")

create_product_sng()







