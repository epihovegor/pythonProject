from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import openpyxl
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains

driver_path = "C:\\chromedriver\\chromedriver.exe"
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = "C:\\Users\\Egor_Epikhov\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe"
driver_service = webdriver.chrome.service.Service(driver_path)
driver_service.start()
driver = webdriver.Chrome(service=driver_service, options=chrome_options)
driver.set_window_size(1800, 1000)
driver.get('https://fcgie.rospotrebnadzor.ru/poisoning-notifications/poisoning?page=1&size=10&sort=-created_at')

driver.find_element(By.ID, "login").send_keys("eochernykh")
driver.find_element(By.ID, "password").send_keys("66RF1reU1quk")
submit_button = driver.find_element(By.ID, "submit-button")
submit_button.click()
time.sleep(3)

wb = openpyxl.load_workbook('C:\\Users\\Egor_Epikhov\\Desktop\\1.xlsx')

sheet = wb['Отчет ЭИОО для ЕИАС']
data = sheet['B3'].value
surname = sheet['D3'].value
name = sheet['E3'].value
surnames = sheet['F3'].value
age = sheet['G3'].value
gender = sheet['I3'].value
social_status = sheet['J3'].value
address = sheet['K3'].value
mp = sheet['N3'].value
dataP = str(sheet['Q3'].value)[:5]
dataFP = sheet['R3'].value.strftime('%d.%m')
pd = sheet['S3'].value
dataY = sheet['U3'].value
hot = sheet['AD3'].value
circum_pois = sheet['AF3'].value
poison_place = sheet['AG3'].value

time.sleep(2)
# # Работает, заполняем Дата и время заполнения извещения МУ
input_element_data = driver.find_element(By.XPATH, '(//span[@role = "button"])[2]').click()
time.sleep(3)
input_element_data = driver.find_element(By.XPATH, '(//input[@class = "ant-input"])[2]')
input_element_data.send_keys(data)
#
# # Работает, заполняем Фамилию
input_element_surname = driver.find_element(By.XPATH, '(//input[@class = "ant-input"])[6]')
input_element_surname.send_keys(surname)

#
# # Работает, заполняем Имя
input_element_name = driver.find_element(By.XPATH, '(//input[@class = "ant-input"])[7]')
input_element_name.send_keys(name)

#
# # Работает, заполняем Отчество
input_element_surnames = driver.find_element(By.XPATH, '(//input[@class = "ant-input"])[8]')
input_element_surnames.send_keys(surnames)

# # Работает, заполняем Возраст
if age == 0:
    checkbox = driver.find_element(By.XPATH, '(//input[@class="ant-checkbox-input"])[1]')
    checkbox.click()
else:
    input_element_name = driver.find_element(By.XPATH, '(//input[@class = "ant-input-number-input"])[1]')
    input_element_name.send_keys(age)

# # Работает, выбираем Пол
if gender:
    if gender.lower() == 'мужской':
        checkbox_label = driver.find_element(By.XPATH, '(//input[@class = "ant-checkbox-input"])[2]')
        checkbox = checkbox_label.find_element(By.XPATH, '(//input[@class = "ant-checkbox-input"])[2]')
        checkbox.click()
    elif gender.lower() == 'женский':
        checkbox_label = driver.find_element(By.XPATH, '(//input[@class = "ant-checkbox-input"])[3]')
        checkbox = checkbox_label.find_element(By.XPATH, '(//input[@class = "ant-checkbox-input"])[3]')
        checkbox.click()

# # Работает, но не выбирается т.к. иерархический список
input_element_social_status = driver.find_element(By.XPATH, '(//input[@class = "ant-select-selection-search-input"])[1]')
input_element_social_status.send_keys(social_status)
time.sleep(2)
# pyautogui.press('Enter')
input_element_social_status.send_keys(Keys.RETURN)

# Работает, социальный статус.
input_element_social_status = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '(//input[@class = "ant-select-selection-search-input"])[1]')))
input_element_social_status.clear()
input_element_social_status.send_keys(social_status)
time.sleep(2)
option_with_text = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{social_status}')]")))
driver.execute_script("arguments[0].scrollIntoView();", option_with_text)
time.sleep(2)
option_with_text.click()

# # Работает, заполняем Адрес места происшествия (в файле должен быть без района)
input_element_address = driver.find_element(By.XPATH, '(//input[@class="search-input"])[2]')
input_element_address.send_keys(address)
time.sleep(2)
input_element_address.send_keys(Keys.RETURN)
time.sleep(2)
#
# #Работает, заполняем Место происшествия
input_element_mp = driver.find_element(By.ID, 'poisoning_place_poisoning_fields_colection_27')
input_element_mp.send_keys(mp)
time.sleep(2)
input_element_mp.send_keys(Keys.RETURN)
time.sleep(2)

# Работает, ставим Дата отравления
input_element_dataP = driver.find_element(By.ID, 'poisoning_dt_poisoning_fields_colection_31')
input_element_dataP.clear()
for char in dataP:
    input_element_dataP.send_keys(char)
    time.sleep(0.5)
time.sleep(2)
# # Работает, ставим Дата первичного обращения
# input_element_dataFP = driver.find_element(By.ID, 'circulation_dt_poisoning_fields_colection_32')
# input_element_dataFP_click = driver.find_element(By.XPATH, '(//span[@role = "button"])[17]')
# input_element_dataFP_click.click()
# time.sleep(2)
# input_element_dataFP = driver.find_element(By.ID, 'circulation_dt_poisoning_fields_colection_32')
# input_element_dataFP.clear()
# for char in dataFP:
#     input_element_dataFP.send_keys(char)
#     time.sleep(0.5)

# Работает, Предварительный диагноз
input_element_pd = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'provisionalDisease_poisoning_fields_colection_1[0]')))
input_element_pd.clear()
input_element_pd.send_keys(pd)
time.sleep(2)
option_with_text = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{pd}')]")))
driver.execute_script("arguments[0].scrollIntoView();", option_with_text)
time.sleep(2)
option_with_text.click()

# Дата, Установлен

# Дата, Подтвержден

# # Работает, заполняем Характер отравления
input_element_hot = driver.find_element(By.ID, 'poisoning_nature_poisoning_fields_colection_40')
input_element_hot.send_keys(hot)
time.sleep(2)
pyautogui.press('Enter')
time.sleep(2)
#
# # Работает, заполняем Обстоятельства отравления
input_element_hot = driver.find_element(By.ID, 'poisoning_circumstances_poisoning_fields_colection_43')
input_element_hot.send_keys(circum_pois)
time.sleep(2)
pyautogui.press('Enter')
time.sleep(2)
#
# # Работает, заполняем Место приобретения яда/химического вещества
input_element_hot = driver.find_element(By.ID, 'purchase_place__poisoning_fields_colection_45')
input_element_hot.send_keys(poison_place)
time.sleep(2)
pyautogui.press('Enter')

time.sleep(100)