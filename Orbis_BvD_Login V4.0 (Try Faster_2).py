# -*- coding: utf-8 -*-
"""
Created on Thu May 24 13:34:26 2018
@author: Shuai
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup as soup
from bs4 import SoupStrainer
from lxml import html
import numpy as np
import time
import pandas as pd
import win32com.client as win32

login_url = "https://orbis4.bvdinfo.com/"
form_data = {"user": "WBG_IFC", "pw": "Global Markets"}

browser = webdriver.Chrome()
browser.get(login_url)
username = browser.find_element_by_name("user")
password = browser.find_element_by_name("pw")
username.send_keys(form_data["user"])
password.send_keys(form_data["pw"])

login_button = browser.find_element_by_class_name("ok")
login_button.click()

def login_orbis(browser,year):

    try:
        search_button = browser.find_element_by_xpath("//a[@data-paramid='AllCompanies']")
        search_button.click()
    except:
        restart_button = browser.find_element_by_xpath("//input[@class='button ok']")
        restart_button.click()
        search_button = browser.find_element_by_xpath("//a[@data-paramid='AllCompanies']")
        search_button.click()

    load_button = browser.find_element_by_css_selector('#search-toolbar > div > div.menu-container > ul > li[data-vs-value=load-search-section]')
    load_button.click()
    time.sleep(3)
    while 1:
        try:
            search_input = browser.find_element_by_css_selector('#tooltabSectionload-search-section > div:nth-child(1) > div.toolbar-tabs-zone-header > div.criterion-search > input.toolbar-find-criterion')
            search_input.clear()
            search_input.send_keys(str(year))
            search_input.send_keys(Keys.RETURN)
            time.sleep(2)
            search_set = browser.find_element_by_css_selector('ul.user-data-item-folder.search-result > li > a > span.name.clickable')
            search_set.click()
            time.sleep(0.5)
            ok_button = browser.find_element_by_css_selector('a.button.submit.ok')
            ok_button.click()
            break
        except:
            continue
        
    data_url = "https://orbis4.bvdinfo.com/version-201866/orbis/1/Companies/List"
    browser.get(data_url)
    time.sleep(3)
    while 1:
        try:
            select_view = browser.find_element_by_css_selector('div.menuViewContainer > div.menuView > ul > li > a')
            select_view.click()  
            break
        except:
            continue
    while 1:
        try:
            view_year = browser.find_element_by_css_selector('span.name.clickable[title="All Columns {0}"]'.format(year))
            view_year.click()
            break
        except:
            continue

def hard_refresh(browser,year,start_page):
    browser.close()
    
    browser = webdriver.Chrome()
    browser.get(login_url)
    username = browser.find_element_by_name("user")
    password = browser.find_element_by_name("pw")
    username.send_keys(form_data["user"])
    password.send_keys(form_data["pw"])
    while 1:
        try:
            login_button = browser.find_element_by_class_name("ok")
            login_button.click()
            break
        except:
            browser.close()
            browser = webdriver.Chrome()
            browser.get(login_url)
    
    login_orbis(browser,year)
    while 1:
        try:
            page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
            page_input.clear()
            page_input.send_keys(str(start_page))
            page_input.send_keys(Keys.RETURN)
            time.sleep(3)
            break
        except:
            continue
    return browser
 
######### Settings for scraping #####################
year_to_get = list(range(2015,2007,-1))
year_to_get = [2014,2013]
start_page = 35661
#####################################################

login_orbis(browser,year_to_get[0])
# =============================================================================
# if_start = input('Start scraping data? (Y/n)\n')
# if if_start != 'Y':
#     exit
# =============================================================================

start_time = time.time()
start_datetime = time.ctime()
big_round_time = start_time
hard_refresh_times = 0
# Starting from the first year

for year in year_to_get:
    
    page_input = browser.find_elements_by_css_selector("ul.navigation > li > input")[0]
    page_input.clear()
    page_input.send_keys(str(start_page))
    page_input.send_keys(Keys.RETURN)
    time.sleep(3)
    
    # html retrieving
    innerHTML = browser.execute_script("return document.body.innerHTML")
    page_soup = soup(innerHTML, "html.parser")
    columns = page_soup.find('table', {'class': 'scroll-header'}).find_next('tr')
    label_info = columns.find_all('div', class_='column-label')
    column_names = []
    for x in label_info:
        try:
            column_names.append(x.span['data-fulllabel'] + ' <' + x.find_all('span')[1]['data-full-configuration'] + '>')
        except:
            column_names.append(x.span['data-fulllabel'])
    column_names = ['company_name'] + column_names
    column_num = len(column_names)
    company_data = pd.DataFrame()
    company_names = []
    for x in column_names:
        company_data[x] = []
    
    page_num = int(page_soup.find('input', {'title': 'Number of page'})['value'])
    
    per_page = 100
    total_companies = int(page_soup.find('td',{'class':'grand-total'}).text.replace(',',''))
    grand_total_pages =  total_companies // per_page + 1 # Number of pages of data to retrieve
    total_pages = grand_total_pages
    per_round = 20
    pages = (start_page // per_round + 1) * (per_round)
    each_time = 20
    pages_each_time = each_time
    
    print("Start retrieving data of year {0} at Page {1}!".format(year, start_page))
    
    strainer_a = SoupStrainer('a', {'data-action': "reporttransfer"})
    strainer_td = SoupStrainer('td', {'class': 'scroll-data'})
    strainer_td = SoupStrainer('td', {'class': 'scroll-data'})
    strainer_input = SoupStrainer('input', {'title': 'Number of page'})
    strainer_tbody = SoupStrainer(id='resultsTable')
    
    stuck_times = 0
    stuck = 0
    teleport = 0
    innerHTML = []
    page_done = 0
    fastest_time = 0
    while page_num <= total_pages:
        teleport = 0
        stuck_times = 0
        has_too_fast = 0
        stopwatch = time.time()
        while 1:
            innerHTML = browser.execute_script("return document.body.innerHTML")
            tree = html.fromstring(innerHTML)
            page_num = int(tree.cssselect('input[type="number"]')[0].attrib['value'])
            if page_num - pages + per_round - 1 == page_done:
                print("Page {0} retrieved!".format(page_num))
                stuck = 0
                page_soup = soup(innerHTML,"lxml").select_one('#resultsTable')
#                try:
#                    company_names += [x.text for x in page_soup.find_all('a',{'data-action': "reporttransfer"})][:100]
#                except:
#                    company_names = [x.text for x in page_soup.find_all('a',{'data-action': "reporttransfer"})][:100]
                try:
                    company_names += [x.text for x in page_soup.select('td.fixed-data > div.fixed-data > table > tbody > tr > td.columnAlignLeft > span > a[href=#]')]
                except:
                    company_names = [x.text for x in page_soup.select('td.fixed-data > div.fixed-data > table > tbody > tr > td.columnAlignLeft > span > a[href=#]')]

                data_points = page_soup.select('td.scroll-data > div > table > tbody > tr > td')
                if page_num == total_pages:
                    num_on_page = total_companies - (page_done+pages-per_round)*per_page
                else:
                    num_on_page = per_page
                data = np.array_split([x.text for x in data_points], num_on_page)
                company_data = company_data.drop("company_name",axis=1)
                company_data = pd.concat([company_data,pd.DataFrame(data,columns=column_names[1:])])
                company_data.insert(0,"company_name",company_names)
                page_done += 1
                print("Page {0} finished!".format(page_num))
                if page_num % pages_each_time == 1:
                    num_to_roll = min(pages_each_time - page_num % pages_each_time + 1, total_pages - page_num)
                    rolled = 0
                if page_done + pages - per_round == total_pages or page_done == per_round:
                    break
                elif page_done % pages_each_time - 1 == rolled or pages_each_time == 1 and teleport == 0:
                    next_button = browser.find_element_by_xpath("//img[@data-action='next']")                 
                    try:
                        for m in range(rolled, num_to_roll):
                            next_button.click()
                        rolled = m
                        print('Rolling!')
                    except:
                        rolled = m
                elif teleport == 1:
                    next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                    num_to_roll = min(pages_each_time - page_num % pages_each_time + 1, total_pages - page_num)
                    rolled = 0
                    while 1:
                        try:
                            for m in range(rolled, num_to_roll):
                                next_button.click()
                            print('Too fast recovered!')
                            break
                        except:
                            rolled = m
                            next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                            continue
                    teleport = 0
                
            elif page_num - pages + per_round - 1 > page_done:
                stuck += 1
                try:
                    if teleport == 0:
                        browser.refresh()
                        page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
                        page_input.clear()
                        page_to_go = page_done + pages - per_round + 1
                        page_input.send_keys(str(page_to_go))
                        page_input.send_keys(Keys.RETURN)
                        print('Too fast. Teleport to Page {0}'.format(page_done + pages - per_round + 1))
                        stuck = 0
                        teleport = 1
                        has_too_fast += 1
                    elif teleport == 1 and stuck >= 3000:
                        while 1:
                            try:
                                page_to_go = page_done + pages - per_round + 1
                                browser = hard_refresh(browser, year, page_to_go)
                                innerHTML = browser.execute_script("return document.body.innerHTML")
                                print('Hard Refreshed!')
                                break
                            except:
                                continue
                        stuck = 0
                        teleport = 1
                        has_too_fast += 1
                except:
                    continue
            else:
                stuck += 1
                if stuck >= 50 and stuck_times < 1:
                    stuck = 0
                    stuck_times += 1
                    next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                    num_to_roll = min(pages_each_time - page_num % pages_each_time + 1, total_pages - page_num)
                    rolled = 0
                    while 1:
                        try:
                            for m in range(rolled, num_to_roll):
                                next_button.click()
                            print('Recovered from stuck!')
                            break
                        except:
                            rolled = m
                            next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                            continue
                elif stuck >= 50 and stuck_times >= 5:
                    try:
                        stuck_times = 0
                        while 1:
                            try:
                                page_to_go = page_done + pages - per_round + 1
                                browser = hard_refresh(browser, year, page_to_go)
                                innerHTML = browser.execute_script("return document.body.innerHTML")
                                print('Hard Refreshed!')
                                break
                            except:
                                continue
                        stuck = 0
                        teleport = 1
                    except:
                        continue
                elif stuck >= 50 and stuck_times >= 1:
                    try:
                        stuck_times += 1
                        page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
                        page_input.clear()
                        page_to_go = page_done + pages - per_round + 1
                        page_input.send_keys(str(page_to_go))
                        page_input.send_keys(Keys.RETURN)
                        print('Too slow. Teleport to Page {0}'.format(page_done + pages - per_round + 1))
                        stuck = 0
                        teleport = 1
                    except:
                        continue                    
        # Output result datatable to csv
        if pages == per_round:
            company_data.to_csv('All_columns-{0}.txt'.format(year), mode='a', sep='|', index=False)
        else:
            company_data.to_csv('All_columns-{0}.txt'.format(year), mode='a', sep='|', index=False, header=False)
        company_data = company_data.iloc[0:0]
        page_done = 0
        company_names = []
        if time.time() - stopwatch < fastest_time or fastest_time == 0:
            fastest_time = time.time() - stopwatch
        print('{0} to {1} pages output! Time cost:{2:.2f}s'.format(pages - per_round + 1, page_num, time.time() - stopwatch))
        
        if page_num == total_pages:
            print('Data of Year {0} is finished!'.format(year))
            start_page = 1
            browser = hard_refresh(browser, year-1, start_page)
            break
        
        # Report each 2ooo pages
        avg_time = (time.time()-start_time-15)/(pages-start_page)*1000
        round_time_spent = (time.time() - big_round_time)/(pages - max(start_page,pages-2000))*1000
        big_round_time = time.time()
        fastest_time = 0
        hard_refresh_times = 0
        try:
            if pages % 2000 == 0:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = 'shuai.qian@outlook.com'
                mail.Subject = 'Orbis Web Scraping System update'
                mail.Body = """System is running.\
                Current Time: {2}.\
                Start Time: {4}.\
                Current Year of Data: {0}.\
                Current Page: {1}. Start Page: {3}. Total Pages: {5}.\
                Average Time per 1000 pages: {6:.2f}s.\
                Time spent on the last 1000 pages: {8:.2f}s\
                Number of Hard Refresh in the last 1000 pages: {9}.\
                Approximately another {7:.2f} hours to finish this year of data.\
                Reported by Orbis Data Scraping System
                """.format(year,pages,time.ctime(),start_page,start_datetime,grand_total_pages,avg_time,(grand_total_pages-pages)/1000*avg_time/3600,round_time_spent,hard_refresh)
                mail.Display()
                mail.Save()
                mail.Close(0)
        except:
            pass
        
        # Turn to the next round
        pages += per_round
        
        # Test if it's time to do hard refresh
        if time.time() - stopwatch >= 1.5*fastest_time and stuck_times <= 1 and has_too_fast < 1 or has_too_fast > 1:
            browser = hard_refresh(browser, year, pages - per_round + 1)
            print('Hard Refreshed!')
            hard_refresh_times += 1
        
print('Successfully output to csv file!')