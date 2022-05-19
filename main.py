import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from random import randrange
from fake_useragent import UserAgent

# from random import randrange
ua = UserAgent()
ua = ua.random

import requests
from bs4 import BeautifulSoup
import lxml

select_www = int(input('Select: 1 - "lae-show-results", 2 - "show-results" '))

if select_www == 1:
    url = 'https://www.nchacutting.com/ncha-shows/world-standings/lae-show-results'
    save_path = './out/1/txt_'
elif select_www == 2:
    url = 'https://www.nchacutting.com/ncha-shows/world-standings/show-results'
    save_path = './out/2/txt_'
else:
    print('No!!!!!!!!!!!!!!!!')

headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "User-Agent": f'{ua}'  # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,
    # like Gecko) Chrome/96.0.4664.45 Safari/537.36"
}

print('start...')

options = webdriver.FirefoxOptions()
options.set_preference("general.useragent.override", f"{ua}")

s = Service('geckodriver.exe')

driver = webdriver.Firefox(service=s, options=options)

driver.implicitly_wait(1.5)
driver.get(url)

time.sleep(5)
source_html = driver.page_source

# # with requests.Session() as session:
# #     response = session.get(url=url, headers=headers)

# # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
# with open('index.html', 'w', encoding='utf-8') as file:
#     file.write(source_html)


# 2
#
# with open("index.html", "r", encoding='utf-8') as f:
#     source_html = f.read()

soup = BeautifulSoup(source_html, 'lxml')

# source_html = driver.page_source

# # # with requests.Session() as session:
# # #     response = session.get(url=url, headers=headers)

# # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
# with open('index2.html', 'w', encoding='utf-8') as file:
#     file.write(source_html)

tab_ = soup.find_all('div', class_='list-group-item small')


def try_div():
    global aaa, iii

    btn_login.click()
    time.sleep(1)

    source_html = driver.page_source

    # with open("index2.html", "r", encoding='utf-8') as f:
    #     source_html = f.read()

    soup = BeautifulSoup(source_html, 'lxml')

    yyy = soup.find_all('a', class_='list-group-item small')

    for mmm in yyy:
        with open(f'{save_path}{f_name}.txt', 'a', encoding='utf-8') as file:
            file.write(str(mmm))
            file.close()

    if select_www == 1:
        xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[3]/div/div/div/div/button'
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
        except:
            xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
    else:
        #xp_close = '//*[@id="main"]/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
        xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
        except:
            xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[3]/div/div/div/div/button'
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
    time.sleep(1)
    # time.sleep(randrange(1, 2))
    aaa += 1
    print(f'Processed: {f_name}    {aaa} / {ttt_count}')
    iii += 1


aaa = 0
iii = 1
xp_iter = 2
ttt_count = len(tab_)

for ttt in tab_:
    f_name = f'{(ttt.text.split("Results")[0].split("#")[-1]).replace(")", "").strip()}'

    try:
        btn_login = driver.find_element(By.XPATH, f'//*[@id="showResultsWidget"]/div[2]/div[{xp_iter}]/div/div[{iii}]/a')
        try_div()
    except:
        iii = 1
        xp_iter += 2
        btn_login = driver.find_element(By.XPATH, f'//*[@id="showResultsWidget"]/div[2]/div[{xp_iter}]/div/div[{iii}]/a')
        try_div()

time.sleep(5)
driver.close()
driver.quit()
