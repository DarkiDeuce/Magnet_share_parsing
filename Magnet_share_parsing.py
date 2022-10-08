import requests
import telebot
import openpyxl
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

file = openpyxl.Workbook()
file['Sheet'].title = 'Акции'
work = file['Акции']

work['A1'] = 'Имя товара'
work['B1'] = 'Категория акции'
work['C1'] = 'Старая цена'
work['D1'] = 'Новая цена'
work['E1'] = 'Скидка'

bot = telebot.TeleBot(" ")
file_path = ''

strings = 1

def last_element():
    url = "https://magnit.ru/promo/"
    s = Service('C:/Users/User/Desktop/All/Activity/Python/Задачи/chromedriver.exe')

    driver = webdriver.Chrome(service=s)

    try:
        driver.get(url=url)
        driver.set_window_size(1920, 1080)
        time.sleep(2)

        # pickle.dump(driver.get_cookies(), open('cookies', 'wb'))
        # for cookie in pickle.load(open('cookies2', 'rb')):
        #     driver.add_cookie(cookie)

        button_city = driver.find_element(By.CLASS_NAME, 'js-geo-another').click()
        time.sleep(2)

        # button_city_list = driver.find_element(By.CLASS_NAME, 'header__contacts-text').click()
        # time.sleep(5)

        input_city = driver.find_element(By.NAME, "citySearch")
        input_city.clear()
        input_city.send_keys("Таштагол")
        time.sleep(5)

        city = driver.find_element(By.CLASS_NAME, 'city-search__link ').click()
        time.sleep(1)

        button = driver.find_element(By.XPATH, '//button[text()="Да"]').click()
        time.sleep(1)

        category_sale = driver.find_element(By.CLASS_NAME, 'select2-selection__rendered').click()
        time.sleep(1)

        option = driver.find_element(By.XPATH, '//option[text()="Цена: по убыванию"]').click()
        time.sleep(1)

        response_last_position = driver.page_source
    finally:
        driver.close()
        driver.quit()

    soup_last_position = BeautifulSoup(response_last_position, 'lxml')

    name_last_position = soup_last_position.find("a", class_="card-sale_catalogue").find("div", class_="card-sale__title").text

    return name_last_position

def pars(page, strings, name_last_element , name_product = ''):
    headers = { 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest'}

    cookies = {'mg_geo_id': '13265'}
    data = {'page': str(page)}

    response = requests.post('https://magnit.ru/promo/', cookies=cookies, headers=headers, data=data)

    soup = BeautifulSoup(response.text, 'lxml')

    card = soup.find_all('a', class_='card-sale_catalogue')
    for product in card:
        try:
            strings = strings + 1
            name_product = product.find('div', class_='card-sale__title').text
            category = product.find("div", class_="card-sale__header").text
            old_cost = product.find("div", class_="label__price label__price_old").find("span", class_="label__price-integer").text
            new_cost = product.find("div", class_="label__price label__price_new").find("span", class_="label__price-integer").text
            sale = product.find("div", class_="card-sale__discount").text

            work[f'A{strings}'] = name_product
            work[f'B{strings}'] = category
            work[f'C{strings}'] = old_cost
            work[f'D{strings}'] = new_cost
            work[f'E{strings}'] = sale

            if name_product == name_last_element:
                break

        except AttributeError:
            continue

    if name_product == name_last_element:
        print("Конец")
        file.save(file_path)
    else:
        pars(page+1, strings, name_last_element)

@bot.message_handler(commands=["stock"])
def stock(message):
    bot.send_message(message.chat.id, "Ожидание...")

    name_last_element = last_element()
    pars(1, strings, name_last_element)

    bot.send_document(message.chat.id, open(file_path, 'rb'))

    os.remove(file_path)

bot.polling(none_stop=True)
