#coding=utf-8
from platform import python_version
import os, time, socket,bs4,requests,string
import numpy as np, pandas as pd
from pandas import ExcelWriter
from bs4 import BeautifulSoup
import selenium
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import *

def investor(stockcode):
    url=f"https://goodinfo.tw/tw/ShowBuySaleChart.asp?STOCK_ID={stockcode}&CHT_CAT=DATE"   #法人持股

    chrome=webdriver.Chrome(executable_path=r'D:\大三\資料工程\chromedriver.exe')

    chrome.get(url)

    chrome.find_element_by_css_selector("#divBuySaleDetailData > table > tbody > tr > td > table > tbody > tr > td:nth-child(1) > nobr:nth-child(1) > select > option:nth-child(2)").click()
    time.sleep(5)
    chrome.find_element_by_css_selector("#divBuySaleDetailData > table > tbody > tr > td > table > tbody > tr > td:nth-child(2) > nobr:nth-child(1) > select > option:nth-child(3)").click()
    time.sleep(5)
    result=pd.read_html(chrome.page_source)

    data=result[25]
    chrome.quit()
    
    return data