import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import re


def transition_time(date_str):
    from datetime import datetime
    date_object = datetime.strptime(date_str, "%Y/%m/%d %H:%M")

    # 將 datetime 物件格式化為新的日期字串
    formatted_date_str = date_object.strftime("%m/%d")
    return formatted_date_str

def transition_delivery_way(text):
    if text == '蝦皮店到店':
        return '蝦皮'
    elif text == '7-ELEVEN':
        return '7-11'
    elif text == 'OK MART':
        return 'OK'
    return text

def filter_emoji(desstr, restr=''):
    try:
        co = re.compile(u'[\U00010000-\U0010ffff]')
    except re.error:
        co = re.compile(u'[\uD800-\uDBFF][\uDC00-\uDFFF]')
    return co.sub(restr, desstr)

def process_product_name(text):
    product_name=''
    color=''
    size=''
    text = filter_emoji(text)
    if "規格:" in text:
        match = re.search(r'(.+?)\n規格: (.+)', text)
        if match:
            product_name = match.group(1)
            specifications = match.group(2)
            if ',' in specifications:
                color, size = specifications.split(',')
    else:
        product_name = text
    return product_name, color, size

def open_and_login(driver):
    #options = Options()
    #options.add_experimental_option("debuggerAddress", "127.0.0.1:9527")
    time.sleep(2)

    login_account = driver.find_element_by_name("loginKey")
    login_password = driver.find_element_by_name("password")
    login_account.send_keys('') # your account
    login_password.send_keys('') # your password

    login_button = driver.find_element_by_class_name("wyhvVD._1EApiB.hq6WM5.L-VL8Q.cepDQ1._7w24N1")
    login_button.click()

def write_order(datas):
    try:
        # workbook = openpyxl.Workbook()
        workbook = openpyxl.load_workbook('test2.xlsx')
        # sheet = workbook['正在進行']
        sheet  =workbook.active
        sheet.title = "order"
        sheet.font = Font(name=u'微軟正黑體')
        sheet.alignment = Alignment(vertical='center', horizontal='center')
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    # 將數據寫入一行
    for data in datas:
        sheet.append(data)
    current_row = sheet.max_row
    print(current_row)
    for i in range(1,7):
        sheet.merge_cells(start_row=max(current_row - len(datas)+1,1), start_column=i, end_row=current_row, end_column=i)
    sheet.merge_cells(start_row=max(1,current_row - len(datas)+1), start_column=10, end_row=current_row, end_column=10)
    # sheet.merge_cells(start_row=1, start_column=11, end_row=1+len(datas), end_column=11)
    workbook.save('test2.xlsx')


driver = webdriver.Chrome(executable_path='chromedriver/chromedriver.exe')
url = 'https://seller.shopee.tw/portal/notification/order-updates/'
driver.get(url)
open_and_login(driver)

# 設定等待時間為10秒
wait = WebDriverWait(driver, 10)

try:
    element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'notification-list')))
except Exception as e:
    # 如果發生異常，印出錯誤信息
    print(f"找不到元素: {e}")

# element = driver.find_elements_by_class_name('notification-list')
#print(element)
# date, account_id, name, delivery, shop, itmes, color, size,  total = [], [], [], [], [], [], [], [], []
next_page_button = driver.find_element_by_class_name('shopee-button.shopee-button--small.shopee-button--frameless.shopee-button--block.shopee-pager__button-next')
next_page_button.click()
time.sleep(2)


while True:
    notification_items = driver.find_elements(By.CLASS_NAME, 'notification-item')
    for item in reversed(notification_items):
        print("into loop")
        #badge_sup = item.find_elements(By.CSS_SELECTOR, '.shopee-badge-x__sup shopee-badge-x__sup--dot.shopee-badge-x__sup--fixed')
        #if badge_sup and badge_sup[0].is_displayed(): #代表是新的
        time.sleep(2)
        try:
            item_button = WebDriverWait(item, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'container-item.clickable'))
            )
            # item_button = item.find_element_by_class_name('container-item.clickable')
            if "你有一筆貨到付款的訂單" in item_button.find_element_by_class_name('title').text:
                item_button.click()
                time.sleep(5)
                driver.switch_to.window(driver.window_handles[1])
                date = transition_time(driver.find_element_by_class_name('time').text)
                # print(driver.find_element_by_class_name('time').text)
                account_id = driver.find_element_by_class_name('username.text-overflow').text
                # print(driver.find_element_by_class_name('username.text-overflow').text) #account id
                name = ""
                info = driver.find_elements_by_class_name('card-style.shopee-card')
                info = info[2].text
                pattern = re.compile(r'買家收件地址\n(.+?)\n')
                match = pattern.search(info)
                if match:
                    buyer_address = match.group(1)
                    name = buyer_address.split(',')[0]

                delivery_way = driver.find_element_by_class_name('carrier').text
                transition_delivery_way(delivery_way)
                delivery_shop = ""
                phone=""
                total = driver.find_element_by_class_name('amount').text
                total = total[3:] # reove NT$
                total_money = int(total.replace(',', '')) # replace 1,386 to 1386

                # print(driver.find_element_by_class_name('shopee-card__content').find_elements_by_class_name('body').text)
                product_list = driver.find_element_by_class_name('product-list')
                products = product_list.find_elements_by_class_name('product-item.product')
                data=[]
                for product in products:
                    product_name, color, size = process_product_name(product.text)

                    data.append([date, account_id, name, phone, delivery_way, delivery_shop, product_name, color, size, total_money])
                    # write_order(time, account_id, name, phone, delivery_way, delivery_shop, product_name, color, size, total_money)

                print(data)
                write_order(data)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

        except NoSuchElementException:
            print("元素未找到，請處理...")
        time.sleep(1.5)
    prev_page_button = driver.find_element(By.CLASS_NAME, 'shopee-pager__button-prev')

    # 获取元素的class name
    class_name = prev_page_button.get_attribute('class')

    # 判断是否包含 'disabled'
    if 'disabled' in class_name: #這是第一頁了
        break
    else: #
        prev_page_button.click()
        time.sleep(3)

print("test over")
driver.quit()
