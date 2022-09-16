import time
import xlsxwriter
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


def Artikul(art):
    inp = art + '.txt'
    st = open(inp)
    st = st.read()
    lst = st.split('\n')
    return lst


def Open_Excel(exc):
    global wb, ws
    inp = exc + '.xlsx'
    wb = xlsxwriter.Workbook(inp)
    ws = wb.add_worksheet()
    ws.write('A1', 'Name')
    ws.write('B1', 'Price')
    ws.write('C1', 'Info')


def Find_Artikul():
    time.sleep(1)
    browser.find_element_by_class_name('search-field__input').clear()
    browser.find_element_by_class_name('search-field__input').send_keys(artikul)
    time.sleep(1)
    browser.find_element_by_class_name('search-field__input').send_keys(Keys.ENTER)
    time.sleep(1)
    browser.find_element_by_class_name('range-revamp-product-compact__image-wrapper').click()


def Input_Excel():
    name = browser.find_element_by_xpath('//*[@id="content"]/div/div[1]/div/div[2]/div[3]/div/div[1]/div/div[1]/h1/div[1]').text
    ws.write('A' + str(i+2), name)
    price = browser.find_element_by_xpath('//*[@id="content"]/div/div[1]/div/div[2]/div[3]/div/div[1]/div/div[2]/div/span/span[1]').text
    ws.write('B' + str(i + 2), price)
    info = browser.find_element_by_xpath('//*[@id="content"]/div/div[1]/div/div[2]/div[3]/div/div[1]/div/div[1]/h1/div[2]/span').text
    ws.write('C' + str(i + 2), info)


def Image():
    pic = requests.get(browser.find_element_by_xpath('//*[@id="content"]/div/div/div/div[2]/div[1]/div/div[2]/div[1]/span/img'))
    out = open("D:\\Project\\ikea\\Pic\\" + "img" + ".jpg", "wb")
    out.write(pic.content)
    out.close()


print('Введите названия txt файла.')
print('Пример: Artikul')
print("Расширение файла писать не надо!")
art = input()
exc = input("Введите название xlsx файла.\n")
# Настройки
options = webdriver.ChromeOptions()
#options.add_argument('--headless')
browser = webdriver.Chrome(executable_path='D:\\Project\\Browser\\chromedriver.exe', options=options)
artikul_list = Artikul(art)
Open_Excel(exc)
try:
    for i in range(0, len(Artikul(art))):
        browser.get(url='https://www.ikea.com/ru/ru/')  # Заходит на сайт
        artikuls = Artikul(art)
        artikul = artikuls[i]
        Find_Artikul()
        Image()
        Input_Excel()
        time.sleep(5)
except Exception as ex:
    print(ex)
finally:
    browser.close()
    browser.quit()
    wb.close()
