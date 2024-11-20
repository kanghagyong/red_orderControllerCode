import time
import json
import pandas as pd
import matplotlib.pyplot as plt

from selenium import webdriver
# import requests
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
# 웹드라이버 생성
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
now = datetime.now()
# 디스플레이 옵션 설정
pd.set_option('display.width', 320)
pd.set_option('display.max_columns', 20)
def restart_browser(driver):
    driver.quit()  # 기존 브라우저 종료
    new_driver = webdriver.Chrome()  # 새 브라우저 시작
    return new_driver

options = ChromeOptions()
options.add_argument('--blink-settings=imagesEnabled=false')
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(3)
itemUrl = 'https://www.redprinting.co.kr/ko'
driver.get(itemUrl)
# start_filename = 'data/case_item_card.xlsx'
# start_filename = 'data/case_item_sticker.xlsx' #'data/case_item_sticker2.xlsx'
start_filename = 'data/case_item_stationery_new.xlsx' #'data/case_item_stationery.xlsx', 'data/case_item_stationery3.xlsx'
option_filename = 'list/stationery_option.xlsx' #'data/case_item_stationery.xlsx', 'data/case_item_stationery3.xlsx'

df_item = pd.read_excel(start_filename, sheet_name = 0, engine='openpyxl')
totalList = len(df_item)
totalcolCount = len(df_item.columns)
# 상품코드 가져오기

# def login_proc(driver):
userid = 'red_openmarket'  # red_openmarket, redprinting
userpw = 'red4874#'  # red4874#, redprinting#1234
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, '#header > div.main-home-header-warp > div > div.header-right > ul > li.select-my > a').click()

# 로그인화면 아이디 및 패스워드 입력
WebDriverWait(driver, 10).until(
    EC.invisibility_of_element_located((By.ID, 'overlay'))
)
driver.find_element(By.ID, 'mb_id').click()
driver.find_element(By.ID, 'mb_id').send_keys(userid)

WebDriverWait(driver, 10).until(
    EC.invisibility_of_element_located((By.ID, 'overlay'))
)
driver.find_element(By.ID, 'mb_password').click()
driver.find_element(By.ID, 'mb_password').send_keys(userpw)

WebDriverWait(driver, 10).until(
    EC.invisibility_of_element_located((By.ID, 'overlay'))
)
driver.find_element(By.ID, 'btnLogin').click()
WebDriverWait(driver, 10).until(
    EC.invisibility_of_element_located((By.ID, 'overlay'))
)
time.sleep(1)


for item in range(totalList):
    time.sleep(2)
    Gc = df_item.iloc[item, 0]
    ic = df_item.iloc[item, 1]
    ti = str(df_item.iloc[item, 2])
    print(ic, ti)
    itemUrl = 'https://www.redprinting.co.kr/ko/product/item/'+Gc+'/'+ic+'/detail/'+ti
    driver.get(itemUrl)
    item_Name = driver.find_element(By.CLASS_NAME, 'pdt_cod_nm').text
    item_Name = item_Name.replace('/', '-').replace('(', '').replace(')', '').replace('~', '').replace('+', '').replace(':', '')
    opt_cnt = str(df_item.iloc[item, 3])
    df_option = pd.read_excel(option_filename, sheet_name=0, engine='openpyxl')
    if opt_cnt == '1':
        time.sleep(3)
        a = 0
        totalSelect = Select(driver.find_element(By.ID, 'select_sub_mtrl'))
        if ic == 'GSSBSTP':
            totalSelect = Select(driver.find_element(By.ID, 'select_sub_mtrl_1'))

        option_first = []
        values = [option.text for option in totalSelect.options if not option.get_attribute("disabled")]
        for v in values:
            if v != '선택해주세요':
                option_first.append(v)
        # print(option_first)
        for opval in option_first:
            time.sleep(1)
            totalS = Select(driver.find_element(By.ID, 'select_sub_mtrl'))
            if ic == 'GSSBSTP':
                totalS = Select(driver.find_element(By.ID, 'select_sub_mtrl_1'))
            totalS.select_by_visible_text(opval)
            time.sleep(0.5)
            driver.find_element(By.ID, 'select_option_btn').click()
            time.sleep(0.5)
            total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
            amount = driver.find_element(By.XPATH, '//*[@id="select_option_item"]/div/div/div[1]/input').get_attribute('value')
            time.sleep(1)
            driver.find_element(By.ID, 'direct_order_btn').click()
            time.sleep(1)
            al = Alert(driver)
            al.accept()
            time.sleep(0.5)
            imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
            df_option.loc[a, 'ProductName'] = item_Name
            df_option.loc[a, 'ItemCode'] = ic
            df_option.loc[a, 'ItemName'] = opval
            df_option.loc[a, 'TmplIndex'] = ti
            df_option.loc[a, 'Amount'] = amount
            df_option.loc[a, 'OrderCode'] = imsiordernum
            df_option.loc[a, 'Price'] = total_price
            # df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
            # print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            print(a, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            a = a+1
            # time.sleep(0.5)
            # driver.find_element(By.CLASS_NAME, 'removeX').click()
            time.sleep(5)
        nowtime = str(now.year) + str(now.month) + str(now.day) + '_' + str(now.hour) + str(now.minute) + str(now.second)
        new_filename = 'reData/' + item_Name + '-'+ic + '_'+ti+'_' + nowtime + '.xlsx'
        print(new_filename)
        df_option.to_excel(new_filename, sheet_name=ic, index=False)
        time.sleep(1)

    elif opt_cnt == '2':
        time.sleep(3)
        # continue
        b = 0
        totalSelect1 = Select(driver.find_element(By.ID, 'select_sub_mtrl_1'))
        option_first = []
        values_f = [option.text for option in totalSelect1.options if not option.get_attribute("disabled")]
        for v in values_f:
            if v != '선택해주세요':
                option_first.append(v)
        # print(option_first)
        #첫번째 옵션 시작
        for of in option_first:
            time.sleep(2)
            totalS1 = Select(driver.find_element(By.ID, 'select_sub_mtrl_1'))
            totalS1.select_by_visible_text(of)
            time.sleep(0.5)
            sub_mtrlval = ""
            for op1 in totalS1.options:
                if op1.text == of:
                    sub_mtrlval = op1.get_attribute('data-type')
            if ti in ['19', '61']:
                totalSelect2 = Select(driver.find_element(By.ID, 'select_sub_mtrl'))
            else:
                if of == '충전용 무지패드 샤이니':
                    totalSelect2 = Select(driver.find_element(By.ID, 'select_sub_mtrl_3'))
                else:
                    totalSelect2 = Select(driver.find_element(By.ID, 'select_sub_mtrl_2'))
            option_second = []
            values_s = [option.text for option in totalSelect2.options if not option.get_attribute("disabled")]
            for v in values_s:
                if v != '선택해주세요':
                    option_second.append(v)
            # print(option_second)
            # print(sub_mtrlval)
            # 두번째 옵션 시작
            for os in option_second:
                time.sleep(2)
                totalS12 = Select(driver.find_element(By.ID, 'select_sub_mtrl_1'))
                totalS12.select_by_visible_text(of)
                if ti in ['19', '61']:
                    totalS22 = Select(driver.find_element(By.ID, 'select_sub_mtrl'))
                else:
                    if of == '충전용 무지패드 샤이니':
                        totalS22 = Select(driver.find_element(By.ID, 'select_sub_mtrl_3'))
                    else:
                        totalS22 = Select(driver.find_element(By.ID, 'select_sub_mtrl_2'))

                time.sleep(1)
                if ic == 'GSSBMTL' and ti in ['37', '38']:
                    totalS22.select_by_visible_text(os)
                else:
                    y=0
                    sub_mtrlval_json = json.loads(sub_mtrlval)
                    for op2 in totalS22.options:
                        if (op2.text == os and sub_mtrlval_json['MTRL_GRP_GB'] in op2.get_attribute('data-type')):
                            totalS22.select_by_index(y)
                        y=y+1
                time.sleep(0.5)
                driver.find_element(By.ID, 'select_option_btn').click()
                time.sleep(0.5)
                total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
                amount = driver.find_element(By.XPATH, '//*[@id="select_option_item"]/div/div/div[1]/input').get_attribute('value')
                time.sleep(1)
                driver.find_element(By.ID, 'direct_order_btn').click()
                time.sleep(1)
                al = Alert(driver)
                al.accept()
                time.sleep(0.5)
                imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
                df_option.loc[b, 'ProductName'] = item_Name
                df_option.loc[b, 'ItemCode'] = ic
                df_option.loc[b, 'ItemName'] = of+'_'+os
                df_option.loc[b, 'TmplIndex'] = ti
                df_option.loc[b, 'Amount'] = amount
                df_option.loc[b, 'OrderCode'] = imsiordernum
                df_option.loc[b, 'Price'] = total_price
                # df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
                # print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
                print(b, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
                b = b+1
                # time.sleep(0.5)
                # removeX_btn = driver.find_element(By.ID, 'select_option_item').find_elements(By.CLASS_NAME, 'removeX')
                # for removeX in removeX_btn:
                #     try:
                #         removeX.click()
                #     except Exception as e:
                #         print(e)

                # driver.find_element(By.CLASS_NAME, 'removeX').click()
                time.sleep(5)
        nowtime = str(now.year) + str(now.month) + str(now.day) + '_' + str(now.hour) + str(now.minute) + str(now.second)
        new_filename = 'reData/'+item_Name+'-' + ic + '_'+ti+'_' + nowtime + '.xlsx'
        print(new_filename)
        df_option.to_excel(new_filename, sheet_name=ic, index=False)
        time.sleep(1)
    elif opt_cnt == '3':
        time.sleep(3)
        c = 0
        if ic == 'GSSBSTP' and ti == '2':
            submtrl1 = 'select_sub_mtrl_1'
            submtrl2 = 'select_sub_mtrl_2'
            submtrl3 = 'select_sub_mtrl_3'
        elif ic == 'GSSBSTP' and ti == '8':
            submtrl1 = 'select_sub_mtrl_2'
            submtrl2 = 'select_sub_mtrl_3'
            submtrl3 = 'select_sub_mtrl_4'
        totalSelect1 = Select(driver.find_element(By.ID, submtrl1))
        option_first = []
        values_f = [option.text for option in totalSelect1.options if not option.get_attribute("disabled")]
        for v in values_f:
            if v != '선택해주세요': option_first.append(v)
        totalSelect2 = Select(driver.find_element(By.ID, submtrl2))
        option_second = []
        values_s = [option.text for option in totalSelect2.options if not option.get_attribute("disabled")]
        for v in values_s:
            if v != '선택해주세요': option_second.append(v)
        totalSelect3 = Select(driver.find_element(By.ID, submtrl3))
        option_third = []
        values_t = [option.text for option in totalSelect3.options if not option.get_attribute("disabled")]
        for v in values_t:
            if v != '선택해주세요': option_third.append(v)
        #첫번째 옵션 시작
        for of in option_first:
            time.sleep(1)
            totalSelect1 = Select(driver.find_element(By.ID, submtrl1))
            time.sleep(0.5)
            totalSelect1.select_by_visible_text(of)
            time.sleep(0.5)
            driver.find_element(By.ID, 'select_option_btn').click()
            time.sleep(0.5)
            total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
            amount = driver.find_element(By.XPATH, '//*[@id="select_option_item"]/div/div/div[1]/input').get_attribute('value')
            time.sleep(1)
            driver.find_element(By.ID, 'direct_order_btn').click()
            time.sleep(1)
            al = Alert(driver)
            al.accept()
            time.sleep(0.5)
            imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
            df_option.loc[c, 'ProductName'] = item_Name
            df_option.loc[c, 'ItemCode'] = ic
            df_option.loc[c, 'ItemName'] = of
            df_option.loc[c, 'TmplIndex'] = ti
            df_option.loc[c, 'Amount'] = amount
            df_option.loc[c, 'OrderCode'] = imsiordernum
            df_option.loc[c, 'Price'] = total_price

            # df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
            # print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            print(c, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            c = c+1
            # time.sleep(0.5)
            # driver.find_element(By.CLASS_NAME, 'removeX').click()
            time.sleep(5)

        for os in option_second:
            time.sleep(1)
            totalSelect2 = Select(driver.find_element(By.ID, submtrl2))
            time.sleep(0.5)
            totalSelect2.select_by_visible_text(os)
            time.sleep(0.5)
            driver.find_element(By.ID, 'select_option_btn').click()
            time.sleep(0.5)
            total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
            amount = driver.find_element(By.XPATH, '//*[@id="select_option_item"]/div/div/div[1]/input').get_attribute('value')
            time.sleep(1)
            driver.find_element(By.ID, 'direct_order_btn').click()
            time.sleep(1)
            al = Alert(driver)
            al.accept()
            time.sleep(0.5)
            imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
            df_option.loc[c, 'ProductName'] = item_Name
            df_option.loc[c, 'ItemCode'] = ic
            df_option.loc[c, 'ItemName'] = os
            df_option.loc[c, 'TmplIndex'] = ti
            df_option.loc[c, 'Amount'] = amount
            df_option.loc[c, 'OrderCode'] = imsiordernum
            df_option.loc[c, 'Price'] = total_price

            # df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
            # print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            print(c, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            c = c+1
            # time.sleep(0.5)
            # driver.find_element(By.CLASS_NAME, 'removeX').click()
            time.sleep(5)

        for ot in option_third:
            time.sleep(1)
            totalSelect3 = Select(driver.find_element(By.ID, submtrl3))
            time.sleep(0.5)
            totalSelect3.select_by_visible_text(ot)
            time.sleep(0.5)
            driver.find_element(By.ID, 'select_option_btn').click()
            time.sleep(0.5)
            total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
            amount = driver.find_element(By.XPATH, '//*[@id="select_option_item"]/div/div/div[1]/input').get_attribute('value')
            time.sleep(1)
            driver.find_element(By.ID, 'direct_order_btn').click()
            time.sleep(1)
            al = Alert(driver)
            al.accept()
            time.sleep(0.5)
            imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
            df_option.loc[c, 'ProductName'] = item_Name
            df_option.loc[c, 'ItemCode'] = ic
            df_option.loc[c, 'ItemName'] = ot
            df_option.loc[c, 'TmplIndex'] = ti
            df_option.loc[c, 'Amount'] = amount
            df_option.loc[c, 'OrderCode'] = imsiordernum
            df_option.loc[c, 'Price'] = total_price

            # df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
            # print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            print(c, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", str(imsiordernum))
            c = c+1
            # time.sleep(0.5)
            # driver.find_element(By.CLASS_NAME, 'removeX').click()
            time.sleep(5)
        nowtime = str(now.year) + str(now.month) + str(now.day) + '_' + str(now.hour) + str(now.minute) + str(now.second)
        new_filename = 'reData/'+item_Name+'-' + ic + '_'+ti+'_' + nowtime + '.xlsx'
        print(new_filename)
        df_option.to_excel(new_filename, sheet_name=ic, index=False)
        time.sleep(1)
        #로그아웃

