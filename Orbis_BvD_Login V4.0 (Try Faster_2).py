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
            search_input.send_keys(str(year))
            break
        except:
            continue
    while 1:
        try:
            search_set = browser.find_element_by_css_selector('#tooltabSectionload-search-section > div:nth-child(1) > div.user-data-item-folder.searches > div:nth-child(7) > div > div.no-scroll-bar > ul.user-data-item-folder.search-result > li > a > span.name.clickable')
            search_set.click()
            break
        except:
            continue
    while 1:
        try:
            ok_button = browser.find_element_by_css_selector('a.button.submit.ok')
            ok_button.click()
            break
        except:
            continue
        
    data_url = "https://orbis4.bvdinfo.com/version-201866/orbis/1/Companies/List"
    browser.get(data_url)
    select_view = browser.find_element_by_css_selector('div.menuViewContainer > div.menuView > ul > li > a')
    while 1:
        try:
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
    
    login_button = browser.find_element_by_class_name("ok")
    login_button.click()
    
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
 
def save_draft(year, pages):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'shuai.qian@outlook.com'
    mail.Subject = 'Orbis Web Scraping System update'
    mail.Body = 'System is running. \nCurrent Time: {2}. Current Year: {0}. Current Page: {1}'.format(year,pages,time.ctime())
    mail.Display()
    mail.Save()
    mail.Close(0)

######### Settings for scraping #####################
year_to_get = list(range(2015,2007,-1))
year_to_get = [2015,2014]
start_page = 72941
#####################################################


login_orbis(browser,year_to_get[0])
if_start = input('Start scraping data? (Y/n)\n')
if if_start != 'Y':
    exit

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
    strainer_input = SoupStrainer('input', {'title': 'Number of page'})
    
    stuck_times = 0
    stuck = 0
    teleport = 0
    innerHTML = []
    while page_num <= total_pages:
        teleport = 0
        stuck_times = 0
        stopwatch = time.time()
        while 1:
            current = browser.execute_script("return document.body.innerHTML")
            soup_input = soup(current, "lxml", parse_only=strainer_input)
            page_num = int(soup_input.input['value'])
            if page_num - pages + per_round - 1 == len(innerHTML):
                innerHTML.append(current)
                print("Page {0} retrieved!".format(page_num))
                stuck = 0
                page_soup = soup(innerHTML[-1],"lxml")         
                try:
                    company_names += [x.text for x in page_soup.find_all('a',{'data-action': "reporttransfer"})][:100]
                except:
                    company_names = [x.text for x in page_soup.find_all('a',{'data-action': "reporttransfer"})][:100]
                data_points = page_soup.find('td', {'class': 'scroll-data'}).find_all('tr')
                company_data = company_data.drop("company_name",axis=1)
                for i in range(min(per_page,total_companies-(page_num-1)*per_page)):
                    company_data.loc[i + len(innerHTML) * per_page] = [x.text for x in data_points[i].find_all('td')]
                company_data.insert(0,"company_name",company_names)
                print("Page {0} finished!".format(page_num))
                if len(innerHTML) + pages - per_round == total_pages or len(innerHTML) == per_round:
                    break
                elif page_num % pages_each_time == 1 or pages_each_time == 1 and teleport == 0:
                    next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                    num_to_roll = min(pages_each_time - page_num % pages_each_time + 1, total_pages - page_num)
                    rolled = 0
                    while 1:
                        try:
                            for m in range(rolled, num_to_roll):
                                next_button.click()
                            print('Rolling!')
                            break
                        except:
                            rolled = m
                            next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                            continue
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
                
            elif page_num - pages + per_round - 1 > len(innerHTML):
                try:
                    if teleport == 0:
                        browser.refresh()
                        page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
                        page_input.clear()
                        page_to_go = len(innerHTML) + pages - per_round + 1
                        page_input.send_keys(str(page_to_go))
                        page_input.send_keys(Keys.RETURN)
                        print('Too fast. Teleport to Page {0}'.format(len(innerHTML) + pages - per_round + 1))
                        stuck = 0
                        teleport = 1
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
                elif stuck >= 50 and stuck_times == 1:
                    try:
                        stuck_times += 1
                        browser.refresh()
                        page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
                        page_input.clear()
                        page_to_go = len(innerHTML) + pages - per_round + 1
                        page_input.send_keys(str(page_to_go))
                        page_input.send_keys(Keys.RETURN)
                        print('Too slow. Teleport to Page {0}'.format(len(innerHTML) + pages - per_round + 1))
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
        innerHTML = []
        company_names = []
        print('{0} to {1} pages output! Time cost:{2:.2f}s'.format(pages - per_round + 1, page_num, time.time() - stopwatch))
        
        if page_num == total_pages:
            print('Data of Year {0} is finished!'.format(year))
            start_page = 1
            browser = hard_refresh(browser, year-1, start_page)
            break
        
        # Report each 2ooo pages
        try:
            if pages % 2000 == 0:
                save_draft(year, pages)
        except:
            pass
        
        # Turn to the next round
        pages += per_round
        
        # Test if it's time to do hard refresh
        if time.time() - stopwatch >= 55*per_round/20:
            browser = hard_refresh(browser, year, pages - per_round + 1)
        
        
print('Successfully output to csv file!')