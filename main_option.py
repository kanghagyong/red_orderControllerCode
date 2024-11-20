import time
import pandas as pd
import matplotlib.pyplot as plt

from selenium import webdriver
import requests
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
# 웹드라이버 생성
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime

now = datetime.now()
# 디스플레이 옵션 설정
pd.set_option('display.width', 320)
pd.set_option('display.max_columns', 20)

itemList_stationery = ['GSSBMTL', 'GSSBACM', 'GSSBSTP']
start_filename = 'list/card_option.xlsx'
df_item_config = pd.read_excel(start_filename, sheet_name = 0, engine='openpyxl', skiprows = 1)
totalList = len(df_item_config)
nowtime = str(now.year) + '' + str(now.month) + '' + str(now.day) + '_' + str(now.hour) + '' + str(
    now.minute) + '' + str(now.second) + '' + str(now.microsecond)

for item in range(totalList):
    # 상품코드 가져오기
    itemCode = df_item_config.iloc[item, 0]
    if itemCode in itemList_stationery:
        sampleOption_filename = 'list/option_sample2.xlsx'
        df_item = pd.read_excel(sampleOption_filename, sheet_name=0, engine='openpyxl', skiprows=1)

    else:
        sampleOption_filename = 'list/option_sample.xlsx'
        df_item = pd.read_excel(sampleOption_filename, sheet_name=0, engine='openpyxl', skiprows=1)
        if "," in df_item_config.iloc[item, 1]:
            papers_wgtlist = df_item_config.iloc[item, 1].split(",")
        else:
            papers_wgtlist = []
            papers_wgtlist.append(df_item_config.iloc[item, 1])
        if "," in df_item_config.iloc[item, 2]:
            dosulist = df_item_config.iloc[item, 2].split(",")
        else:
            dosulist = []
            dosulist.append(df_item_config.iloc[item, 2])
        if "," in df_item_config.iloc[item, 3]:
            sizelist = df_item_config.iloc[item, 3].split(",")
        else:
            sizelist = []
            sizelist.append(df_item_config.iloc[item, 3])
        if "," in str(df_item_config.iloc[item, 4]):
            amountlist = df_item_config.iloc[item, 4].split(",")
        else:
            amountlist = []
            amountlist.append(str(df_item_config.iloc[item, 4]))
        if df_item_config.iloc[item, 5] == 'x':
            afterpcs01_list = ['x']
        else:
            afterpcs01_list = df_item_config.iloc[item, 5].split(",")

        if df_item_config.iloc[item, 6] == 'x':
            afterpcs02_list = ['x']
        else:
            afterpcs02_list = df_item_config.iloc[item, 6].split(",")

        if df_item_config.iloc[item, 7] == 'x':
            afterpcs03_list = ['x']
        else:
            afterpcs03_list = df_item_config.iloc[item, 7].split(",")

        if df_item_config.iloc[item, 8] == 'x':
            afterpcs04_list = ['x']
        else:
            afterpcs04_list = df_item_config.iloc[item, 8].split(",")

        if df_item_config.iloc[item, 9] == 'x':
            afterpcs05_list = ['x']
        else:
            afterpcs05_list = df_item_config.iloc[item, 9].split(",")

        # print(papers_wgtlist)
        # print(dosulist)
        # print(sizelist)
        # print(amountlist)
        # print(afterpcs01_list)
        # print(afterpcs02_list)
        # print(afterpcs03_list)
        # print(afterpcs04_list)
        # print(afterpcs05_list)

        papers = []  # 용지리스트
        wgt = []  # 용지무게 리스트
        for pw in papers_wgtlist:
            pwdata = pw.split("_")
            if pwdata[0] in papers:
                time.sleep(0.2)
            else:
                papers.append(pwdata)

        i=0
        for p in papers:
            for d in dosulist:
                for s in sizelist:
                    for a in amountlist:
                        for ap1 in afterpcs01_list:
                            # if ap1 == '유광':
                            for ap2 in afterpcs02_list:
                                for ap3 in afterpcs03_list:
                                    for ap4 in afterpcs04_list:
                                        for ap5 in afterpcs05_list:
                                            df_item.loc[i, 'ItemCode'] = itemCode
                                            df_item.loc[i, 'Papers'] = p[0]
                                            df_item.loc[i, 'WgtCod'] = p[1]
                                            df_item.loc[i, 'Dosu'] = d
                                            df_item.loc[i, 'Sizes'] = s
                                            df_item.loc[i, 'Amount'] = a
                                            df_item.loc[i, 'AfterPcs01'] = ap1
                                            df_item.loc[i, 'AfterPcs02'] = ap2
                                            df_item.loc[i, 'AfterPcs03'] = ap3
                                            df_item.loc[i, 'AfterPcs04'] = ap4
                                            df_item.loc[i, 'AfterPcs05'] = ap5
                                            df_item.loc[i, 'OrderCode'] = ""
                                            df_item.loc[i, 'Price'] = ""
                                            i = i + 1

        print(i)
        new_filename = 'data/' + itemCode + '_' + nowtime + '.xlsx'
        df_item.to_excel(new_filename, sheet_name=itemCode, index=False, startrow=7)

# ItemCode	Papers	WgtCod	Dosu	Sizes	Amount	AfterPcs01	AfterPcs02	AfterPcs03	OrderCode	Price
# df_item.loc[i, 'ItemCode'] = ""

