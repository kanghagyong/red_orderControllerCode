from flask import Flask, send_from_directory, render_template_string
import os
import time
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
# import requests
from selenium.webdriver.common.alert import Alert
from typing import Optional
from selenium.webdriver.common.by import By
# 웹드라이버 생성
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.common.exceptions import WebDriverException, UnexpectedAlertPresentException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
now = datetime.now()

pd.set_option('display.width', 320)
pd.set_option('display.max_columns', 20)
import tempfile

bizhouse = Flask(__name__)

BIZ_FOLDER = 'bizhouse_file'
os.makedirs(BIZ_FOLDER, exist_ok=True)

@bizhouse.route('/')
def index():
    files = [f for f in os.listdir(BIZ_FOLDER) if f.endswith('.xlsx')]
    file_links = ''.join(f'<li><a href="/download/{f}">{f}</a></li>' for f in files)
    first_fileupload =  render_template_string('''
        <div style="width:30%;float:left;">
            <h1>다운로드 파일</h1>
            <h1>다운로드 목록</h1>
            <ul>
                {{ files|safe }}
            </ul>
        </div>
    ''', files=file_links)



    return first_fileupload

if __name__ == '__main__':
    bizhouse.run(debug=True, host='0.0.0.0', port=5005)
