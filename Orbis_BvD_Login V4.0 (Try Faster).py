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

try:
    search_button = browser.find_element_by_xpath("//a[@data-paramid='AllCompanies']")
    search_button.click()
except:
    restart_button = browser.find_element_by_xpath("//input[@class='button ok']")
    restart_button.click()
    search_button = browser.find_element_by_xpath("//a[@data-paramid='AllCompanies']")
    search_button.click()

# Ask if filters are all set
while 1:
    if_ready = input('Are filters all set? (Y/n)\n')
    if if_ready == 'Y':
        break

indicator_url = "https://orbis4.bvdinfo.com/version-2018523/orbis/1/Companies/ListEdition"
browser.get(indicator_url)

# Ask if indicators (columns) are all set
while 1:
    if_ready = input(
        "Are indicators (columns) all set and ready to retrieve data? (Y/n)\n(Make sure you've clicked 'Apply' and no repetitive indicator exists)\n")
    if if_ready == 'Y':
        break

start_page = 14541

# Starting from the 1st Page
#page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
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

print("Start retrieving data!")

strainer_a = SoupStrainer('a', {'data-action': "reporttransfer"})
strainer_td = SoupStrainer('td', {'class': 'scroll-data'})
strainer_input = SoupStrainer('input', {'title': 'Number of page'})

stuck = 0
teleport = 0
innerHTML = []
while page_num <= total_pages:
    teleport = 0
    while 1:
        current = browser.execute_script("return document.body.innerHTML")
        
        soup_input = soup(current, "lxml", parse_only=strainer_input)
        page_num = int(soup_input.input['value'])
        if page_num - pages + per_round - 1 == len(innerHTML):
            innerHTML.append(current)
            print("Page {0} retrieved!".format(page_num))
            stuck = 0
            teleport = 0
            if len(innerHTML) + pages - per_round == total_pages or len(innerHTML) == per_round:
                break
            elif pages_each_time == 1 or page_num % pages_each_time == 1:
                next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                try:
                    for m in range(min(pages_each_time, per_round - page_num % per_round + 1)):
                        next_button.click()
                    pages_each_time = each_time
                except:
                    continue
            
        elif page_num - pages + per_round - 1 > len(innerHTML):
            try:
                if teleport == 0:
                    print('Too fast. Teleport to Page {0}'.format(len(innerHTML) + pages - per_round + 1))
                    page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
                    page_input.clear()
                    page_to_go = len(innerHTML) + pages - per_round + 1
                    page_input.send_keys(str(page_to_go))
                    page_input.send_keys(Keys.RETURN)
                    next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                    try:
                        for m in range(min(pages_each_time, per_round - page_num % per_round + 1)):
                            next_button.click()
                        pages_each_time = each_time
                    except:
                        continue
                    stuck = 0
                    teleport = 1
            except:
                continue
        else:
            stuck += 1
            if stuck >= 50:
                print('stuck!')
                pages_each_time = pages - page_num + 1
                stuck = 0
                next_button = browser.find_element_by_xpath("//img[@data-action='next']")
                try:
                    for m in range(min(pages_each_time, per_round - page_num % per_round + 1)):
                        next_button.click()
                    pages_each_time = each_time
                except:
                    continue
    
    if page_num > pages + 1:
        page_input = browser.find_element_by_xpath("//input[@title='Number of page']")
        page_input.clear()
        page_input.send_keys(str(page_num + 1))
        page_input.send_keys(Keys.RETURN)

    # Parsing HTML
    for page_done in range(per_round):
        page_soup = soup(innerHTML[page_done],"lxml")
        company_names = [x.text for x in page_soup.find_all('a',{'data-action': "reporttransfer"})]
# =============================================================================
#         soup_a = soup(innerHTML[page_done], "lxml", parse_only=strainer_a)
#         company_names = [x.text for x in soup_a.find_all()]
#         soup_td = soup(innerHTML[page_done], "lxml", parse_only=strainer_td)
#         data_points = soup_td.find_all('tr')        
# =============================================================================
#        tds = soup_td.td.find_all('td')
        data_points = page_soup.find('td', {'class': 'scroll-data'}).find_all('tr')
        for i in range(0, per_page):
            company_data.loc[i + (page_done) * per_page] = (
                        [company_names[i]] + [x.text for x in data_points[i].find_all('td')])
# =============================================================================
#             company_data.loc[i + (page_done) * per_page] = (
#                         [company_names[i]] + [x.text for x in tds[(column_num-1)*i:(i+1)*(column_num-1)]])
# =============================================================================

            if page_done + 1 + pages - per_round == total_pages and i == total_pages * per_page - total_companies:
                break

        print("Page {0} finished!".format(page_done + 1 + pages - per_round))
        if page_done + 1 + pages - per_round == total_pages:
            break
        else:
            page_done += 1
    
    if pages == per_round:
        company_data.to_csv('company_data_50.txt', mode='a', sep='|', index=False)
    else:
        company_data.to_csv('company_data_50.txt', mode='a', sep='|', index=False, header=False)
    company_data = company_data.iloc[0:0]
    innerHTML = []
    print('{0} to {1} pages output!'.format(pages - per_round + 1, page_num))
    if page_num == total_pages:
        break
    pages += per_round
    
print('Successfully output to csv file!')