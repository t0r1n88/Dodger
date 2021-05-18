from selenium import webdriver
from bs4 import BeautifulSoup
from requests_html import HTML
import requests

URL = 'https://www.list-org.com/'
# browser = webdriver.Chrome()
# browser.implicitly_wait(10)
# test_lst = ['0326013559', '0323339496', '	0326481349']
# browser.get('https://www.list-org.com/?search=inn')
#
# # Находим элемент ИНН
# input_inn = browser.find_element_by_class_name('search_input')
# input_inn.send_keys('0326013559')
# button_find = browser.find_element_by_css_selector('div.main:nth-child(1) div.content:nth-child(2) div:nth-child(1) form.bord.frm1 div.bord.search_div:nth-child(5) > button.search_btn.btn.btn-default')
# button_find.click()
#
# button_name = browser.find_element_by_css_selector('div.main:nth-child(1) div.content div.org_list:nth-child(2) p:nth-child(1) label:nth-child(1) > a:nth-child(2)').click()
#
# Для отработки используем заранее скачаную старницу
html = open('data/html_1.html', encoding='utf8').read()

soup = BeautifulSoup(html, 'lxml')


item = soup.find(('div', {'class': 'org_list'}))
url = item.findAll('a')
for i in url:
    print(i)
