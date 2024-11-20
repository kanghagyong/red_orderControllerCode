import time
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

itemList_card = ['BCSPDFT']
itemList_sticker = ['STCUXXX']
itemList_stationery = ['GSSBMTL', 'GSSBACM', 'GSSBSTP']
options = ChromeOptions()
options.add_argument('--blink-settings=imagesEnabled=false')
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(3)
# start_filename = 'data/case_item_card.xlsx'
# start_filename = 'data/case_item_sticker.xlsx' #'data/case_item_sticker2.xlsx'
start_filename = 'data/STCUXXX_redopenmarket.xlsx' #'data/case_item_stationery.xlsx', 'data/case_item_stationery3.xlsx'
# csv 파일 읽어와서 dataframe으로 저장
# df_item = pd.read_csv('data/inpress_241010.csv')
# df_item = pd.read_csv('data/case_item.csv')
df_item_config = pd.read_excel(start_filename, sheet_name = 0, engine='openpyxl', nrows=1)
df_item = pd.read_excel(start_filename, sheet_name = 0, engine='openpyxl', skiprows = 7)
totalList = len(df_item)
totalcolCount = len(df_item.columns)
# 상품코드 가져오기
itemUrl = df_item_config.iloc[0,0]
itemCode = df_item_config.iloc[0,1]
userid = df_item_config.iloc[0,2]
userpw = df_item_config.iloc[0,3]
if userid == "x":
    userid = 'redprinting'
if userpw == "x":
    userpw = 'redprinting#1234'

time.sleep(1)
driver.get(itemUrl)

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
driver.execute_script('window.scrollTo(0, 200)')
# try:
# 각 상품코드 별 어떤 형태인지 체크
if itemCode in itemList_card:
    for i in range(totalList):
        paper = df_item.iloc[i, 1]
        wgtcod = df_item.iloc[i, 2]
        dosu = df_item.iloc[i, 3]
        size = df_item.iloc[i, 4]
        amount = df_item.iloc[i, 5]
        apcs = df_item.iloc[i, 6]
        sizesplit = size.split("*")
        wid_size = sizesplit[0]
        hei_size = sizesplit[1]

        if driver.find_element(By.ID, 'paperSelectBoxItText').text != paper:
            # 상품 페이지 용지선택
            time.sleep(2)
            driver.find_element(By.ID, 'paperSelectBoxItContainer').click()
            time.sleep(1)
            driver.find_element(By.LINK_TEXT, paper).click()

        time.sleep(1)
        if wgtcod != "":
            # 상품 페이지 G수 선택
            if driver.find_element(By.ID, 'paper_sub_selectSelectBoxItText').text != str(wgtcod):
                driver.find_element(By.ID, 'paper_sub_selectSelectBoxItContainer').click()
                time.sleep(1)
                driver.find_element(By.LINK_TEXT, str(wgtcod)).click()

        time.sleep(1)
        if dosu != "":
            # 상품 페이지 인쇄도수 선택
            if driver.find_element(By.ID, 'soduSelectBoxItText').text != dosu:
                driver.find_element(By.ID, 'soduSelectBoxItContainer').click()
                time.sleep(1)
                driver.find_element(By.LINK_TEXT, dosu).click()

        time.sleep(1)
        if driver.find_element(By.ID, 'sizeSelectBoxItText').text != '사이즈직접입력':
            # 사이즈 직접입력 선택
            driver.find_element(By.ID, 'sizeSelectBoxItContainer').click()
            time.sleep(1)
            driver.find_element(By.LINK_TEXT, '사이즈직접입력').click()

        time.sleep(0.5)
        if driver.find_element(By.ID, 'CUT_WDT').get_attribute('value') != wid_size:
            # 상품 페이지 사이즈 직접입력 사이즈 입력
            driver.find_element(By.ID, 'CUT_WDT').click()
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(wid_size)

        time.sleep(0.5)
        if driver.find_element(By.ID, 'CUT_HGH').get_attribute('value') != hei_size:
            driver.find_element(By.ID, 'CUT_HGH').click()
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(hei_size)

        time.sleep(0.5)
        driver.find_element(By.ID, 'WRK_HGH').click()

        time.sleep(0.5)
        if amount != "":
            if driver.find_element(By.ID, 'number1_selSelectBoxItText').text != str(amount):
                # 상품 페이지 인쇄도수 선택
                driver.find_element(By.ID, 'number1_selSelectBoxItContainer').click()
                time.sleep(1)
                driver.find_element(By.LINK_TEXT, str(amount)).click()

        time.sleep(0.5)
        apcs_nowstatus = driver.find_element(By.ID, 'opt_string').get_attribute('value')
        if apcs == "무광":
            if 'COT_DFT' in apcs_nowstatus:
                time.sleep(0.5)
            else:
                # 상품 페이지 코팅 선택
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');")
            time.sleep(0.5)
            driver.execute_script("productOrder.opt_select('COT_DFT','MA');")# 상품 페이지 무광코팅 선택
        elif apcs == "유광":
            if 'COT_DFT' in apcs_nowstatus:
                time.sleep(0.5)
            else:
                # 상품 페이지 코팅 선택
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');")
            time.sleep(0.5)
            driver.execute_script("productOrder.opt_select('COT_DFT','GL');")# 상품 페이지 유광코팅 선택
        else:
            if 'COT_DFT' in apcs_nowstatus:
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');") #코팅 후가공 다시 선택 시 선택해제됨.
            else:
                time.sleep(0.5)

        # driver.find_element(By.XPATH, '/html/body/div[9]/div/aside/div[5]/div[2]/div[3]/span[1]/label').click()
        # if driver.find_element(By.XPATH, '/html/body/div[11]/div/aside/div[5]/div[2]/div[3]/span[1]/input').get_attribute('checked'):
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, 'body > div.page_order > div > aside > div.search_tool.search_tool_update > div.confirmbox-wrap > div.radio-group-big > span.radio-group-inner > label').click()

        time.sleep(0.5)
        driver.execute_script('window.scrollTo(0, 300)')
        time.sleep(0.5)
        total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
        df_item.loc[i, 'Price'] = total_price
        time.sleep(0.5)
        driver.find_element(By.ID, 'direct_order_btn').click()
        time.sleep(1)
        al = Alert(driver)
        al.accept()
        time.sleep(0.5)
        imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
        df_item.loc[i, 'OrderCode'] = imsiordernum #주문관리코드생성후추가
        print(i, "번째 TOTAL_PRICE : ", total_price, "pot_tmp_cod : ", imsiordernum)
        time.sleep(5)

elif itemCode in itemList_sticker:
    for i in range(totalList):
        paper = df_item.iloc[i, 1]
        wgtcod = df_item.iloc[i, 2]
        dosu = df_item.iloc[i, 3]
        size = df_item.iloc[i, 4]
        amount = df_item.iloc[i, 5]
        apcs1 = df_item.iloc[i, 6]
        apcs2 = df_item.iloc[i, 7]
        apcs3 = df_item.iloc[i, 8]
        apcs4 = df_item.iloc[i, 9]
        apcs5 = df_item.iloc[i, 10]
        sizesplit = size.split("*")
        wid_size = sizesplit[0]
        hei_size = sizesplit[1]
        time.sleep(2)
        paper_text = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "paperSelectBoxItText"))
        )
        if paper_text.text != paper:
            # 상품 페이지 용지선택
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, 'overlay'))
            )
            driver.find_element(By.ID, 'paperSelectBoxItContainer').click()
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, 'overlay'))
            )
            driver.find_element(By.LINK_TEXT, paper).click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        if wgtcod != "":
            # 상품 페이지 G수 선택
            wgt_text = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.ID, "paper_sub_selectSelectBoxItText"))
            )
            if wgt_text.text != str(wgtcod):
                driver.find_element(By.ID, 'paper_sub_selectSelectBoxItContainer').click()
                WebDriverWait(driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, 'overlay'))
                )
                driver.find_element(By.LINK_TEXT, str(wgtcod)).click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        if dosu != "":
            # 상품 페이지 인쇄도수 선택
            docu_text = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.ID, "soduSelectBoxItText"))
            )
            if docu_text.text != dosu:
                driver.find_element(By.ID, 'soduSelectBoxItContainer').click()
                WebDriverWait(driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, 'overlay'))
                )
                driver.find_element(By.LINK_TEXT, dosu).click()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        size_text = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "sizeSelectBoxItText"))
        )
        if size_text.text != '사이즈직접입력':
            # 사이즈 직접입력 선택
            driver.find_element(By.ID, 'sizeSelectBoxItContainer').click()
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, 'overlay'))
            )
            driver.find_element(By.LINK_TEXT, '사이즈직접입력').click()

        CUT_WDT_text = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "CUT_WDT"))
        )
        if CUT_WDT_text.get_attribute('value') != wid_size:
            # 상품 페이지 사이즈 직접입력 사이즈 입력
            driver.find_element(By.ID, 'CUT_WDT').click()
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(Keys.BACK_SPACE)
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_WDT').send_keys(wid_size)

        # WebDriverWait(driver, 10).until(
        #     EC.invisibility_of_element_located((By.ID, 'overlay'))
        # )
        CUT_HGH_text = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "CUT_HGH"))
        )
        if CUT_HGH_text.get_attribute('value') != hei_size:
            driver.find_element(By.ID, 'CUT_HGH').click()
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(Keys.BACK_SPACE)
            time.sleep(0.5)
            driver.find_element(By.ID, 'CUT_HGH').send_keys(hei_size)

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        driver.find_element(By.XPATH, '//*[@id="WRK_HGH"]').click()
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        driver.execute_script("productOrder.check_PRN_CNT();")

        number1_text = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.ID, "number1"))
        )
        if number1_text.get_attribute('value') != str(amount):
            time.sleep(0.5)
            driver.find_element(By.ID, 'number1').click()
            driver.find_element(By.ID, 'number1').send_keys(Keys.DELETE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.DELETE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.DELETE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.DELETE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.BACK_SPACE)
            driver.find_element(By.ID, 'number1').send_keys(Keys.BACK_SPACE)
            time.sleep(0.5)
            driver.find_element(By.ID, 'number1').send_keys(str(amount))

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )
        apcs_nowstatus = driver.find_element(By.ID, 'opt_string').get_attribute('value')
        if apcs1 == "무광":
            if 'COT_DFT' in apcs_nowstatus:
                time.sleep(0.2)
            else:
                # 상품 페이지 코팅 선택
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');")
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, 'overlay'))
            )
            driver.execute_script("productOrder.opt_select('COT_DFT','MA');")# 상품 페이지 무광코팅 선택
        elif apcs1 == "유광":
            if 'COT_DFT' in apcs_nowstatus:
                time.sleep(0.2)
            else:
                # 상품 페이지 코팅 선택
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');")
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, 'overlay'))
            )
            driver.execute_script("productOrder.opt_select('COT_DFT','GL');")# 상품 페이지 유광코팅 선택
        else:
            if 'COT_DFT' in apcs_nowstatus:
                driver.execute_script("productOrder.opt_use_yn('COT_DFT', 'SID_S');") #코팅 후가공 다시 선택 시 선택해제됨.
            else:
                time.sleep(0.2)
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, 'overlay'))
        )

        if apcs4 == "묶음재단":
            driver.execute_script("productOrder.opt_checked('CUT_DFT', 'DFXXX');")  # 상품 페이지 묶음재단 선택
        elif apcs4 == "개별재단":
            driver.execute_script("productOrder.opt_checked('CUT_DFT', 'DFITM');")  # 상품 페이지 개별재단 선택
        else:
            time.sleep(0.2)

        WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )

        driver.find_element(By.XPATH, '//*[@id="WRK_HGH"]').click()
        total_price = driver.find_element(By.ID, 'TOTAL_PRICE').text
        df_item.loc[i, 'Price'] = total_price
        WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )
        time.sleep(1)
        driver.find_element(By.ID, 'direct_order_btn').click()
        time.sleep(1)
        al = Alert(driver)
        al.accept()
        time.sleep(0.5)
        imsiordernum = driver.find_element(By.ID, 'pot_tmp_cod').get_attribute('value')
        df_item.loc[i, 'OrderCode'] = imsiordernum
        print(i, "번째 TOTAL_PRICE : ", str(total_price), "pot_tmp_cod : ", imsiordernum)
        WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )
        # driver.refresh()
        time.sleep(5)
else:
    print('error')

nowtime=str(now.year)+'_'+str(now.month)+'_'+str(now.day)+'_'+str(now.hour)+'_'+str(now.minute)
new_filename = 'reData/STCUXXX_redopenmarket'+'_' + nowtime + '.xlsx'
df_item.to_excel(new_filename, sheet_name=itemCode, index=False)

# finally:
#     driver.quit()
