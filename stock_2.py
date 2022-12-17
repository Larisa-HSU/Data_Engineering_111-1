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

#webdriver路徑記得要改

def investor(stockcode):
    url=f"https://goodinfo.tw/tw/ShowBuySaleChart.asp?STOCK_ID={stockcode}&CHT_CAT=DATE"   #法人持股:一年

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

    
def kchart(stockcode):
    url=f"https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID={stockcode}&CHT_CAT2=DATE"   #k線圖:一年

    chrome=webdriver.Chrome(executable_path=r'D:\大三\資料工程\chromedriver.exe')

    chrome.get(url)

    chrome.find_element_by_css_selector("#selK_ChartPeriod > option:nth-child(3)").click()
    time.sleep(2)
    result=pd.read_html(chrome.page_source)

    data=result[25]
    chrome.quit()
    
    return data

def shareholder(stockcode):
    url=f"https://goodinfo.tw/tw/EquityDistributionCatHis.asp?STOCK_ID={stockcode}"   #股東結構:比例

    chrome=webdriver.Chrome(executable_path=r'D:\大三\資料工程\chromedriver.exe')

    chrome.get(url)

    result=pd.read_html(chrome.page_source)

    data=result[17]
    chrome.quit()
    
    return data


def cash(stockcode):
    url=f"https://goodinfo.tw/tw/StockFinDetail.asp?RPT_CAT=CF_M_QUAR_ACC&STOCK_ID={stockcode}"   #現金流量

    chrome=webdriver.Chrome(executable_path=r'D:\大三\資料工程\chromedriver.exe')

    chrome.get(url)

    chrome.find_element_by_css_selector("#RPT_CAT > option:nth-child(1)").click()
    time.sleep(2)
    result=pd.read_html(chrome.page_source)

    data=pd.DataFrame(result[15].iloc[73,1:]).T
    chrome.quit()
    
    return data
