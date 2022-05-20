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
# from selenium.Java
import time
from random import randrange
from fake_useragent import UserAgent

# from random import randrange
ua = UserAgent()
ua = ua.random

import requests
from bs4 import BeautifulSoup
import lxml
import json
import time

import xlsxwriter

workbook = xlsxwriter.Workbook('out.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True, 'font_color': 'red'})
bold.set_align('center')

bold_1 = workbook.add_format({'bold': True, 'font_color': 'black'})
bold_1.set_align('center')

bold_2 = workbook.add_format({'bold': True, 'font_color': 'blue'})
bold_2.set_align('center')

bold_3 = workbook.add_format({'bold': True, 'font_color': 'black'})
bold_3 = workbook.add_format({'bg_color': '#b4b4b4'})
bold_3.set_align('center')

data_format1 = workbook.add_format({'bg_color': '#b4b4b4'})
data_format1.set_align('center')
#
# =========================================================

# Format the first column
worksheet.set_column('A:A', 25, data_format1)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 25)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25)

worksheet.set_default_row(25)

worksheet.write('A1', 'Company name', bold_3)
worksheet.write('B1', 'Description', bold_1)
worksheet.write('C1', 'Sector', bold_1)
worksheet.write('D1', 'E-Mail', bold_1)
worksheet.write('E1', 'Phone', bold_1)
worksheet.write('F1', 'Link', bold_1)



url = 'https://www.pollutec.com/fr-fr/liste-exposants.html'


headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "User-Agent": f'{ua}'  # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,
    # like Gecko) Chrome/96.0.4664.45 Safari/537.36"
}

print('start...')

# # 1
# #
# options = webdriver.FirefoxOptions()
# options.set_preference("general.useragent.override", f"{ua}")
#
# s = Service('geckodriver.exe')
#
# driver = webdriver.Firefox(service=s, options=options)
#
# driver.implicitly_wait(1.5)
# driver.get(url)
#
# time.sleep(5)
# source_html = driver.page_source
#
#
# #//*[@id="exhibitor-directory"]/div/div/div/div[2]/div[3]/div/ul/div[103]/div/div[2]/div/div[1]/div[1]/div[1]/a/h3
#
# #WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
#
#
#
# # # '//*[@id="exhibitor-directory"]/div/div/div/div[2]/div[3]/div/ul/div[1]/div/div[2]/div/div[1]/div[1]/div[1]/a'
# # # '//*[@id="exhibitor-directory"]/div/div/div/div[2]/div[3]/div/ul/div[47]/div/div[2]/div/div[1]/div[1]/div[2]/span/a'
#
#
# for i in range(1, 1501):
#     print(f'TRY: {i} X-PATH')
#     for j in range(10):  # adjust integer value for need
#         try:
#             tit_ = driver.find_element(By.XPATH, f'//*[@id="exhibitor-directory"]/div/div/div/div[2]/div[3]/div/ul/div[{i}]/div/div[2]/div/div[1]/div[1]/div[1]/a/h3')  # )).click()
#             break
#         except:
#             driver.execute_script("window.scrollBy(0, 20000)")
#         print(j)
#         time.sleep(1)
#
#     # time.sleep(0.5)
#
#     a_s = f'#exhibitor-directory > div > div > div > div.filter-results > div:nth-child(3) > div > ul > div:nth-child({i}) > div > div.flexible-content > div > div.description-container.col-md-8.col-xs-12 > div.company-info > div.tags.single-line-ellipsis > span > a'
#     a_ = driver.find_elements(By.CSS_SELECTOR, a_s) #a_ = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, a_x)))
#
#     if not a_:
#         a_s = f'#exhibitor-directory > div > div > div > div.filter-results > div:nth-child(3) > div > ul > div:nth-child({i}) > div > div.flexible-content > div > div.description-container.col-md-8.col-xs-12 > div.company-info > div:nth-child(1) > a'
#         a_ = driver.find_elements(By.CSS_SELECTOR, a_s)
#
#     link = [elem.get_attribute('href') for elem in a_]
#
#     print(tit_.text)
#     print(link)
#
#     # запись ссылок из СПИСКА в файл
#     with open('urls.txt', 'a', encoding='utf-8') as file:
#         for url in link:
#             file.write(f'{url}\n')
#
# driver.close()
# driver.quit()
#
# END of 1 ===========================================================================================================

# 2
#
# читаю ССЫЛКИ из ранее созданного файла
# !!! ОБРЕЗАЮ СИМВОЛ ПЕРЕНОСА СТРОКИ !!!
with open('urls.txt') as file:
    url_list = [line.strip() for line in file.readlines()]

# СЧЁТЧИК количества ЛОТОВ(ссылок)
url_count = len(url_list)
print(url_count)

with requests.Session() as session:
    # d_t = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    # d_t_now = datetime.strptime(d_t, '%d-%m-%Y %H:%M:%S')

    # d_t = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # если нужно обработать ТОЛЬКО пять ССЫЛОК
    # for url in url_list[:5]:
    #
    # чтобы не городить СЧЁТЧИК по типу: Х++,
    # оберну ... в "enumerate()" в результате получаю
    # и ССЫЛКУ и ИНДЕКС места, на котором она находится!!!(кортеж)
    # for url in enumerate(url_list[:5]):
    #
    # !!! 777 in to_do !!!
    #
    error_ = []
    row = 2
    for i, url in enumerate(url_list[0:121]):
        # for url in url_list:
        # print(url)

        options = webdriver.FirefoxOptions()
        options.set_preference("general.useragent.override", f"{ua}")

        s = Service('geckodriver.exe')

        driver = webdriver.Firefox(service=s, options=options)

        driver.implicitly_wait(1.5)
        driver.get(url)

        # time.sleep(3)
        click_xp = '//*[@id="exhibitor_details_phone"]/p/a'

        start_time = time.time()
        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, click_xp)))
        except:
            error_.append(url)
        finish_time = time.time() - start_time

        source_html = driver.page_source

        driver.close()
        driver.quit()

        soup = BeautifulSoup(source_html, 'lxml')



        # #************************************************
        # #
        # response = session.get(url=url, headers=headers)
        # soup = BeautifulSoup(response.text, 'lxml')
        # #
        # #************************************************





        # # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
        # source_html = driver.page_source
        # with open('index.html', 'w', encoding='utf-8') as file:
        #     file.write(soup)

        # content > div:nth-child(4) > div.box_title.auction_page_title

        # !!! СТРАНИЦА ЛОТА !!!
        #
        # название
        try:
            title_ = soup.find('h1', class_='wrap-word').text.strip()
        except:
            title_ = 'NONE'

        try:
            description_ = soup.find('div', class_='form-group-view-mode wrap-word exhibitor-details-description').text.strip()
            description_ = description_[11:]
        except:
            description_ = 'NONE'

        try:
            sec_ = soup.find("div", {"data-dtm-category-id": "8564"}).find_all('span', class_='label label-default label-in-list tag-item')
            secteurs = ';'.join([str(elem.text) for elem in sec_])
        except:
            secteurs = 'NONE'

        try:
            e_mail = soup.find("div", {"id": 'exhibitor_details_email'}).find('a').text
        except:
            e_mail = 'NONE'

        try:
            tel_ = soup.find("div", {"id": 'exhibitor_details_phone'}).find('a').text
        except:
            tel_ = 'NONE'






        # print(title_)
        # print(description_)
        # print(f'link: {url}')
        # print(f'e-mail: {e_mail}')
        # print(f'tel: {tel_}')
        print(f'#{row - 1} ---> {finish_time} ---> {title_}')


        worksheet.write(f'A{row}', title_, bold_1) # worksheet.write_url(f'F{row}', url, string=f'{lot_num}')
        worksheet.write(f'B{row}', description_, bold_1)
        worksheet.write(f'C{row}', secteurs, bold_1)
        worksheet.write(f'D{row}', e_mail, bold_1)
        worksheet.write(f'E{row}', tel_, bold_1)
        worksheet.write(f'F{row}', url, bold_2)

        #     worksheet.write(f'C{row}', volume)
        #     worksheet.write(f'D{row}', d24h_)
        #     worksheet.write(f'E{row}', d7d)
        #     worksheet.write(f'F{row}', floor_price)
        #     worksheet.write(f'G{row}', num_owners)
        #     worksheet.write(f'H{row}', items)
        #
        row = row + 1
        #
    workbook.close()

with open('error_.json', 'w', encoding='utf-8') as file:
   json.dump(error_, file, indent=4, ensure_ascii=False)


        #
        # # ССЫЛКА на картинку
        # lot_img = soup.find('div', class_='ad-page-image-holder').find('img').get('src')





#
# END of 2 ===========================================================================================================












# # breakpoint()
# #
# #
# #
# # # # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
# # # with open('index.html', 'w', encoding='utf-8') as file:
# # #     file.write(source_html)
# #
# # # driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
# # for i in range(10):  # adjust integer value for need
# #     # you can change right side number for scroll convenience or destination
# #     driver.execute_script("window.scrollBy(0, 5000)")
# #     # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
# #     with open(f'index{i}.html', 'w', encoding='utf-8') as file:
# #         file.write(source_html)
# #     # you can change time integer to float or remove
# #     time.sleep(1)
#
# # time.sleep(5)
# driver.close()
# driver.quit()
#
#
# # # 2
# # #
# # with open("index.html", "r", encoding='utf-8') as f:
# #     source_html = f.read()
#
# soup = BeautifulSoup(source_html, 'lxml')
#
# # source_html = driver.page_source
#
# # # # with requests.Session() as session:
# # # #     response = session.get(url=url, headers=headers)
#
# # # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
# # with open('index2.html', 'w', encoding='utf-8') as file:
# #     file.write(source_html)
#
# tab_ = soup.find_all('div', class_='description-container col-md-8 col-xs-12')
#
# test = []
#
# for i in tab_:
#     title_ = i.find('h3', class_='text-center-mobile wrap-word')
#     test.append(i)
#
# print(len(test))
#
# # with open('test.json', 'w', encoding='utf-8') as file:
# #     json.dump(test, file, indent=4, ensure_ascii=False)
#
# breakpoint()
#
#
# def try_div():
#     global aaa, iii
#
#     btn_login.click()
#     time.sleep(1)
#
#     source_html = driver.page_source
#
#     # with open("index2.html", "r", encoding='utf-8') as f:
#     #     source_html = f.read()
#
#     soup = BeautifulSoup(source_html, 'lxml')
#
#     yyy = soup.find_all('a', class_='list-group-item small')
#
#     for mmm in yyy:
#         with open(f'{save_path}{f_name}.txt', 'a', encoding='utf-8') as file:
#             file.write(str(mmm))
#             file.close()
#
#     if select_www == 1:
#         xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[3]/div/div/div/div/button'
#         try:
#             WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
#         except:
#             xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
#             WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
#     else:
#         #xp_close = '//*[@id="main"]/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
#         xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[4]/div/div/div/div/button'
#         try:
#             WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
#         except:
#             xp_close = '/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[3]/div/div/div/div/button'
#             WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp_close))).click()
#     time.sleep(1)
#     # time.sleep(randrange(1, 2))
#     aaa += 1
#     print(f'Processed: {f_name}    {aaa} / {ttt_count}')
#     iii += 1
#
#
# aaa = 0
# iii = 1
# xp_iter = 2
# ttt_count = len(tab_)
#
# for ttt in tab_:
#     f_name = f'{(ttt.text.split("Results")[0].split("#")[-1]).replace(")", "").strip()}'
#
#     try:
#         btn_login = driver.find_element(By.XPATH, f'//*[@id="showResultsWidget"]/div[2]/div[{xp_iter}]/div/div[{iii}]/a')
#         try_div()
#     except:
#         iii = 1
#         xp_iter += 2
#         btn_login = driver.find_element(By.XPATH, f'//*[@id="showResultsWidget"]/div[2]/div[{xp_iter}]/div/div[{iii}]/a')
#         try_div()
#
# time.sleep(5)
# driver.close()
# driver.quit()
