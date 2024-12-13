import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException
import gspread

chromedriver_path = "chromedriver.exe"  

service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

def find_element_with_retry(driver, by_method, value, retry=200):
    attempt = 0
    while attempt < retry:
        try:
            time.sleep(0.1)
            element = driver.find_element(by_method, value)
            return element
        except NoSuchElementException:
            print(f"Attempt {attempt + 1}: Element not found.")
            attempt += 1
    print("Exceeded maximum retries. Couldn't find the element.")
    raise NoSuchElementException(f"Element with {by_method}: {value} not found after {retry} attempts")

def click_element_with_retry(driver, by_method, value, retry=200):
    attempt = 0
    while attempt < retry:
        try:
            element = find_element_with_retry(driver, by_method, value, retry)
            element.click()
            return
        except ElementClickInterceptedException:
            print(f"Attempt {attempt + 1}: Element not clickable, retrying...")
            time.sleep(0.5)
            attempt += 1
    print("Exceeded maximum retries. Couldn't click the element.")
    raise ElementClickInterceptedException(f"Element with {by_method}: {value} could not be clicked after {retry} attempts")


gc = gspread.service_account('service_account.json')
spreadsheet = gc.open('JHC')
sheet = spreadsheet.get_worksheet(0)

urls = sheet.get_all_values()
urls = [url[0] for url in urls]

for index, url in enumerate(urls):
    if (url == ''):
        print(f"============ Skipping {index} == {url} ============")
        continue
    print(f"============ Start updating == {index} == {url} ============")
    driver.get(url)
    
    target = find_element_with_retry(driver,By.XPATH, "//*[contains(text(), '型號')]")
    driver.execute_script("arguments[0].scrollIntoView();", target)
    click_element_with_retry(driver, By.XPATH, "//*[contains(text(), '查看門市庫存')]", 200)
    time.sleep(3) 
    regions = ['香港區', '九龍區', '新界區', '離島區/偏遠地區', '澳門區']
    store_data = []

    for region in regions:
        select_element = Select(find_element_with_retry(driver, By.ID, 'region-root-W41'))
        print("Got region select element.")
        try:
            print("Trying to select region: " + region)         
            select_element.select_by_visible_text(region)
            print("Selected region: " + region)
        except NoSuchElementException:
            print(f"Not found region: {region}")
        
        time.sleep(3) 
        
        product_title_element = find_element_with_retry(driver, By.CLASS_NAME, 'productFullDetail-productName-uvJ')
        print("Product title element found." + product_title_element.text)
        print("===================")
        product_title = product_title_element.text.strip()

        address_list = find_element_with_retry(driver, By.CLASS_NAME, 'shoplistRoot', 200)
        li_elements = address_list.find_elements(By.TAG_NAME, 'li') 
        for li in li_elements:
            lines = li.text.split('\n') 
            if len(lines) >= 4:  
                address = lines[0].strip()  
                phone = lines[1].strip()     
                opening_hours = lines[2].strip()  
                stock_status = lines[3].strip()    
            
                store_data.append({
                    'Region': region,
                    'Product Title': product_title,
                    'Address': address,
                    'Phone': phone,
                    'Opening Hours': opening_hours,
                    'Stock Status': stock_status
                })
    df = pd.DataFrame(store_data)
    print(f"============{url} Information Done============")

    sheet_index = index + 1  
    sheet = spreadsheet.get_worksheet(sheet_index)
    sheet.clear() 

    values = [df.columns.values.tolist()] + df.values.tolist()
    sheet.update('A1', values) 

    print(f"============== Updated {len(values)} Rows to Google Sheets Sheet {sheet_index}.")

    sheet_index += 1  

print("===================")
driver.quit()
