import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException
import gspread

# 设置 ChromeDriver 路径
chromedriver_path = "chromedriver.exe"  # 替换为你的 ChromeDriver 路径

# 启动浏览器
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

# Google Sheets 初始化
gc = gspread.service_account('service_account.json')
spreadsheet = gc.open('JHC')
sheet = spreadsheet.get_worksheet(0)
# 所有链接
urls = sheet.get_all_values()
urls = [url[0] for url in urls]
# urls = [
#     "https://www.jhceshop.com/hk_tc/6912504332232.html",
#     "https://www.jhceshop.com/hk_tc/4891203502394.html",
#     "https://www.jhceshop.com/hk_tc/4891203053902.html",
#     "https://www.jhceshop.com/hk_tc/4891203499502.html",
#     "https://www.jhceshop.com/hk_tc/4891203051403.html",
#     "https://www.jhceshop.com/hk_tc/4891203501748.html",
#     "https://www.jhceshop.com/hk_tc/6912504338371.html",
#     "https://www.jhceshop.com/hk_tc/068060463494.html",
#     "https://www.jhceshop.com/hk_tc/051131830271.html"
# ]


# 存储每个链接的工作表索引

# 遍历每个链接
for index, url in enumerate(urls):
    if (url == ''):
        print(f"============ Skipping {index} == {url} ============")
        continue
    print(f"============ Start updating == {index} == {url} ============")
    driver.get(url)
    

    # 等待页面完全加载
    target = find_element_with_retry(driver,By.XPATH, "//*[contains(text(), '型號')]")
    driver.execute_script("arguments[0].scrollIntoView();", target)
    click_element_with_retry(driver, By.XPATH, "//*[contains(text(), '查看門市庫存')]", 200)
    time.sleep(3)  # 可根据实际情况调整等待时间

    # 选择区域并提取店铺信息
    regions = ['香港區', '九龍區', '新界區', '離島區/偏遠地區', '澳門區']
    store_data = []

    for region in regions:
        # 选择区域
        select_element = Select(find_element_with_retry(driver, By.ID, 'region-root-W41'))
        print("Got region select element.")
        try:
            print("Trying to select region: " + region)         
            select_element.select_by_visible_text(region)
            print("Selected region: " + region)
        except NoSuchElementException:
            print(f"Not found region: {region}")
        
        # 等待页面更新
        time.sleep(3)  # 可根据实际情况调整等待时间
        
        # 获取产品标题
        product_title_element = find_element_with_retry(driver, By.CLASS_NAME, 'productFullDetail-productName-uvJ')
        print("Product title element found." + product_title_element.text)
        print("===================")
        product_title = product_title_element.text.strip()  # 获取并去除前后空格

        # 提取店铺信息
        address_list = find_element_with_retry(driver, By.CLASS_NAME, 'shoplistRoot', 200)
        li_elements = address_list.find_elements(By.TAG_NAME, 'li')  # 选择合适的选择器
        for li in li_elements:
            lines = li.text.split('\n')  # 按行分割文本
            if len(lines) >= 4:  # 确保有足够的行
                address = lines[0].strip()  # 地址
                phone = lines[1].strip()     # 电话
                opening_hours = lines[2].strip()  # 开门时间
                stock_status = lines[3].strip()    # 存货状态
            
                store_data.append({
                    'Region': region,
                    'Product Title': product_title,  # 添加产品标题
                    'Address': address,
                    'Phone': phone,
                    'Opening Hours': opening_hours,
                    'Stock Status': stock_status
                })
    # 转换为 DataFrame
    df = pd.DataFrame(store_data)
    print(f"============{url} Information Done============")

    # 获取相应的工作表    
    sheet_index = index + 1  # 从第8个工作表开始
    sheet = spreadsheet.get_worksheet(sheet_index)
    sheet.clear()  # 清除现有数据

    # 将数据写入 Google Sheets
    values = [df.columns.values.tolist()] + df.values.tolist()
    sheet.update('A1', values)  # 从单元格 A1 开始写入数据

    print(f"============== Updated {len(values)} Rows to Google Sheets Sheet {sheet_index}.")

    # 更新工作表索引
    sheet_index += 1  # 下一个工作表

# 关闭浏览器
print("===================")
driver.quit()
