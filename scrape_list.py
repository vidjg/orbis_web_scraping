# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 16:32:46 2018

@author: sqian
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup as soup
from bs4 import SoupStrainer
from lxml import html
import requests
import numpy as np
import time
import pandas as pd
import win32com.client as win32

browser = webdriver.Chrome()

# Test for once
page = 0
url = "https://datacatalog.worldbank.org/search/indicators?search_api_views_fulltext_op=AND&query=&nid=&sort_by=changed&sort_order=DESC&0=type%3Aindicators&page=0%2C{0}&f[0]=field_wbddh_data_type%3A293&f[1]=type%3Aindicators".format(page)
browser.get(url)
time.sleep(5)
innerHTML = browser.execute_script("return document.body.innerHTML")
tree = html.fromstring(innerHTML)


for page in range(0, 1606):
    url = "https://datacatalog.worldbank.org/search/indicators?search_api_views_fulltext_op=AND&query=&nid=&sort_by=changed&sort_order=DESC&0=type%3Aindicators&page=0%2C{0}&f[0]=field_wbddh_data_type%3A293&f[1]=type%3Aindicators".format(page)
    browser.get(url)
    innerHTML = browser.execute_script("return document.body.innerHTML")
    tree = html.fromstring(innerHTML)
    
    data_temp = pd.DataFrame(columns=['title','period','database','source','last_update'])
    
    row_num = min(11, 16056 - 10 * page + 1)
    for i in range(1,row_num):
        data_dict = {}
        try:
            data_dict['title'] = tree.xpath('//div[contains(@class,"views-row-{0} ")]/div[starts-with(@id,"node"),1]/h2/a/text()'.format(i))
        except:
            data_dict['title'] = ''
        try:
            data_dict['period'] = tree.xpath('//div[contains(@class,"views-row-{0} ")]/div[starts-with(@id,"node")]/div[2]/span/span[text()=" Periodicity:"]/following-sibling::b/text()'.format(i))
        except:
            data_dict['period'] = ''
        try:
            data_dict['database'] = [x.strip() for x in tree.xpath('//div[contains(@class,"views-row-{0} ")]/div[starts-with(@id,"node")]/div[2]/span/span[text()="Dataset:"]/following-sibling::b/text()'.format(i))]
        except:
            data_dict['database'] = ''
        try:
            data_dict['source'] = [x.strip() for x in tree.xpath('//div[contains(@class,"views-row-{0} ")]/div[starts-with(@id,"node")]/div[2]/span/span[text()="Source:"]/following-sibling::b/text()'.format(i))]
        except:
            data_dict['source'] = ''
        try:
            data_dict['last_update'] = tree.xpath('//div[contains(@class,"views-row-{0} ")]/div[starts-with(@id,"node")]/div[2]/span/span[text()="Last Updated:"]/following-sibling::b/text()'.format(i))
        except:
            data_dict['last_update'] = ''
#        data_dict['title'] = tree.xpath('//*[starts-with(@id,"node"),1]/h2/a/text()')
#        data_dict['period'] = tree.xpath('//*[starts-with(@id,"node")]/div[2]/span/span[text()=" Periodicity:"]/following-sibling::b/text()')
#        data_dict['database'] = [x.strip() for x in tree.xpath('//*[starts-with(@id,"node")]/div[2]/span/span[text()="Dataset:"]/following-sibling::b/text()')]
#        # data_dict['source'] = [x.strip() for x in tree.xpath('//*[starts-with(@id,"node")]/div[2]/span/span[text()="Source:"]/following-sibling::b/text()')]
#        data_dict['last_update'] = tree.xpath('//*[starts-with(@id,"node")]/div[2]/span/span[text()="Last Updated:"]/following-sibling::b/text()')
        
    data_temp = data_temp.append(data_dict, ignore_index=True)
    
    if page == 0:
        data_temp.to_csv('indicators_WBG.csv', index=False, header=True)
    else:
        data_temp.to_csv('indicators_WBG.csv', mode='a', index=False, header=False)
    print('Page {0} Done! {1:.2f}% Finished!'.format(page+1, (page+1)/1605*100))