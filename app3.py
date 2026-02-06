import pyautogui as pyautogui
from flask import Flask, send_from_directory, render_template_string, render_template_string, request
import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException, UnexpectedAlertPresentException, NoAlertPresentException

now = datetime.now()

pd.set_option('display.width', 320)
pd.set_option('display.max_columns', 20)
import tempfile

app3 = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
LIST_FOLDER = 'list'
DATA_FOLDER = 'data'
REDATA_FOLDER = 'reData'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app3.route('/')
def index():
    xlsx_files1 = [f for f in os.listdir(REDATA_FOLDER) if f.endswith('.txt')]
    file_links1 = ''.join(f'<li><a href="/download2/{f}">{f}</a></li>' for f in xlsx_files1)
    html = '''
        <div style="width:100%;float:left;">
            <h1>코드생성 파일업</h1>

            <form action="/upload2" method="post">

                <h3>색상 선택</h3>
                {% for c in colors %}
                    <label>
                        <input type="radio" name="color" value="{{ c }}" style="width:20px;height:20px;"> {{ c }}
                    </label>
                {% endfor %}

                <h3>사이즈 선택</h3>
                <span>2:S, 3:M, 4:L, 5:XL, 6:XXL</span>
                <select name="type" style="width:100px;height:50px;">
                    {% for s in types %}
                        <option value="{{ s }}">{{ s }}</option>
                    {% endfor %}
                </select>

                <br><br>
                <input type="submit" value="파일생성">
            </form>

            <h1>다운로드 목록</h1>
            <ul>
                {{ files1|safe }}
            </ul>
        </div>
        '''

    return render_template_string(
        html,
        files1=file_links1,
        colors=['화이트', '모쿠그레이', '블랙', '라이트블루', '라이트핑크'],
        types=['adult', 'child']
    )


@app3.route('/upload2', methods=['POST'])
def upload2_file():
    selected_color = request.form.get('color')   # radio
    selected_type = request.form.get('type')     # select

    print("선택 색상:", selected_color)
    print("선택 사이즈:", selected_type)

    uploadfile_ordernum_creating(selected_color, selected_type)
    time.sleep(3)

    return '파일이 업로드되었습니다!<br><a href="/">목록</a>'

# 로그인 체크 프로세스
def login_check_proc(userid, userpw, itemUrl, driver, itemCode):
    time.sleep(1)
    driver.get(itemUrl)

    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, '#header > div.main-home-header-warp > div > div.header-right > ul > li.select-my > a').click()

    # 로그인화면 아이디 및 패스워드 입력
    WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )
    driver.find_element(By.ID, 'mb_id').click()
    driver.find_element(By.ID, 'mb_id').send_keys(userid)
    WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )
    driver.find_element(By.ID, 'mb_password').click()
    driver.find_element(By.ID, 'mb_password').send_keys(userpw)
    WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )
    driver.find_element(By.ID, 'btnLogin').click()
    WebDriverWait(driver, 10).until( EC.invisibility_of_element_located((By.ID, 'overlay')) )

    # driver.execute_script('window.scrollTo(0, 200)')

def create_driver():
    options = ChromeOptions()
    options.add_argument('--blink-settings=imagesEnabled=false')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    # options.add_argument('--headless=new')  # 안정화 후 적용

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(3)
    return driver

# 일반적인 상품 주문관리 코드 생성 로직
def uploadfile_ordernum_creating(color, type):

    driver = create_driver()

    # 상품코드 가져오기
    itemUrl = 'https://www.redprinting.co.kr/ko/product/item/CL/CLTMSHS'
    itemCode = 'CLTMSHS'
    userid = 'red_openmarket' #red_openmarket, #redprinting
    userpw = 'guest1004!' #red4874# , #redprinting#1234

    login_check_proc(userid, userpw, itemUrl, driver, itemCode)
    try:
        driver.find_element(By.XPATH, '//*[@id="widget"]/div/article[4]/div[2]/div/div/button').click()
        time.sleep(0.5)
        # colorList = ['스포츠그레이', '블랙', '네이비', '마룬', '포레스트그린']
        sizeList = ['1', '2', '3', '4', '5']
        countList = [10, 50, 100, 200, 500]
        select_color(driver, color)
        # select_size(driver, size)
        select_type(driver, type)

        if type == 'child':
            sizeList = ['1', '2', '3', '4']

        printArea = [
            '좌측가슴,x,x'
            ,'x,x,뒷면'
            ,'좌측가슴,x,x'
            ,'x,앞면,x'
            ,'x,x,뒷면'
        ]


        pojanglist = ['N', 'Y']
        results = []
        # for color in colorList:
        #     select_color(driver, color)
        #
        for size in sizeList:
            select_size(driver, size)

            # for pojang in pojanglist:
            #     select_pojang(driver, pojang)

            for cnt in countList:
                select_count(driver, cnt)

                i=0
                for area in printArea:
                    i = i + 1
                    select_print_area(driver, area)
                    #print(area)
                    time.sleep(1.5)  # 가격 DOM 갱신 대기
                    if i == 6 :
                        print("인쇄영역 초기화")
                    else:
                        print("인쇄영역 변경.")
                        try:
                            driver.execute_script("fnPreOrderPot('pot_create', event);")
                            time.sleep(6)
                        except Exception as e:
                            print("JS 실행 오류:", e)

        print('생성완료.')
        # driver.execute_script('window.location.reload();')
        time.sleep(2)

    except WebDriverException as e:
        print(f"WebDriver 오류 발생: {e}")

    finally:
        print(f"finally 끝:")
        try:
            driver.quit()
        except:
            pass

selector_map = {
    "color": ".cloth-color",
    "size": ".button-wrap",
    "area": ".print-area"
}

def select_color(driver, color):
    container = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector_map["color"]))
    )
    el = container.find_element(
        By.XPATH, f".//div[contains(normalize-space(text()), '{color}')]"
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)
    driver.execute_script("arguments[0].click();", el)
    time.sleep(0.5)


def select_size(driver, size):
    driver.find_element(By.XPATH, '//*[@id="widget"]/div/article[3]/div[3]/button['+size+']').click()
    time.sleep(0.5)

def select_pojang(driver, po):
    driver.find_element(By.ID, po).click()
    time.sleep(0.5)

def select_type(driver, type):
    driver.find_element(By.ID, type).click()
    time.sleep(1)

def select_count(driver, cnt):
    qty_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "PRN_CNT"))
    )
    driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
        """, qty_input, cnt)

def select_print_area(driver, area_code):
    position_map = {
        0: '좌측가슴',
        1: '앞면',
        2: '뒷면'
    }

    container = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector_map["area"]))
    )
    # 예: A,x,C,x,x,x → 분해
    areas = area_code.split(",")
    print("areas", areas)
    for idx, code in enumerate(areas):
        if code == "x":
            continue

        alt_text = position_map[idx]

        img = container.find_element(
            By.XPATH, f".//img[@alt='{alt_text}']"
        )

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", img)
        time.sleep(0.15)
        driver.execute_script("arguments[0].click();", img)
    time.sleep(0.5)

def get_price(driver):
    price_el = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".price strong"))
    )
    return price_el.text.replace(',', '').replace('원', '').strip()

if __name__ == '__main__':
    app3.run(debug=True, host='0.0.0.0', port=5003)
