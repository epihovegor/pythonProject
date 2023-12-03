from datetime import datetime
import openpyxl
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

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
dP = str(sheet['Q3'].value)[:5]
dFP = sheet['R3'].value.strftime('%d.%m')
medorg = sheet['W3'].value
dY = str(sheet['U3'].value)[:5]
dPP = str(sheet['V3'].value)[:5]

# # Работает, ставим Дата первичного обращения
input_element_dFP = driver.find_element(By.ID, 'circulation_dt_poisoning_fields_colection_32')
input_element_dFP.clear()
for char in dFP:
    input_element_dFP.send_keys(char)
    time.sleep(0.5)

# Работает, ставим Дата отравления
input_element_dP = driver.find_element(By.ID, 'poisoning_dt_poisoning_fields_colection_31')
input_element_dP.clear()
for char in dP:
    input_element_dP.send_keys(char)
    time.sleep(0.5)

# Дата, Установлен
input_element_dY = driver.find_element(By.XPATH, '(//span[@role = "button"])[19]').click()
time.sleep(3)
input_element_dY = driver.find_element(By.ID, 'determined_dt_poisoning_fields_colection_3[0]')
input_element_dY.clear()
for char in dY:
    input_element_dY.send_keys(char)
    time.sleep(0.5)

# Дата, Подтвержден
input_element_dPP = driver.find_element(By.XPATH, '(//span[@role = "button"])[21]').click()
time.sleep(3)
input_element_dPP = driver.find_element(By.ID, 'confirmed_dt_poisoning_fields_colection_4[0]')
input_element_dPP.clear()
for char in dPP:
    input_element_dPP.send_keys(char)
    time.sleep(0.5)

time.sleep(10)