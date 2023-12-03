# Хватает первую ячейку A1 из excel заносит в поисковую строку ya.ru находит, очищает поисковую строку, вставляет A2
# и т.д. пока не наткнется на пустую ячейку, всё что находит, выводит в терминал
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


class WebDriverAutomation:
    def __init__(self, driver_path, chrome_binary_location, window_size=(1800, 1000)):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.binary_location = chrome_binary_location
        driver_service = webdriver.chrome.service.Service(driver_path)
        driver_service.start()
        self.driver = webdriver.Chrome(service=driver_service, options=chrome_options)
        self.driver.set_window_size(*window_size)

    def navigate_to_url(self, url):
        self.driver.get(url)

    def load_excel_data(self, file_path, sheet_name='Sheet1'):
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]
        return sheet

    def perform_search(self, search_query):
        input_element = self.driver.find_element(By.XPATH, '//input[@name = "text"]')
        input_element.send_keys(search_query)
        pyautogui.press('Enter')


# Пример использования класса
if __name__ == "__main__":
    driver_path = "C:\\chromedriver\\chromedriver.exe"
    chrome_binary_location = "C:\\Users\\Egor_Epikhov\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe"
    automation = WebDriverAutomation(driver_path, chrome_binary_location)

    # Навигация по URL
    automation.navigate_to_url('https://ya.ru')

    # Загрузка данных из Excel
    sheet = automation.load_excel_data('C:\\Users\\Egor_Epikhov\\Desktop\\2.xlsx')
    column_a = sheet['A']

    # Цикл по всем ячейкам в столбце A
    for cell in column_a:
        search_query = cell.value
        if search_query is not None:
            # Выполнение поиска
            automation.perform_search(search_query)

            # Добавьте здесь необходимые действия после успешного поиска
            print(f"Найдено: {search_query}")

            # Очистка поля поиска для следующей итерации
            input_element = automation.driver.find_element(By.XPATH, '//input[@name = "text"]')
            input_element.clear()

    # Закрытие веб-драйвера
    automation.driver.quit()