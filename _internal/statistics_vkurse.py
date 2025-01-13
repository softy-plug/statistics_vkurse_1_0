import os

# Установка необходимых библиотек
os.system("pip install openpyxl")
os.system("pip install selenium")
os.system("pip install webdriver-manager")

import time
import glob
import shutil  # Импортируем shutil для перемещения файлов
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager

input("Для запуска программы нажмите Enter")

# Путь к Вашему Excel-файлу
exl_path = r"vkurse.xlsx"
exl = load_workbook(exl_path)

# Работа с листом data
sheet_data = exl["data"]

login_exl = sheet_data.cell(row=2, column=1).value  # Логин
password_exl = sheet_data.cell(row=2, column=2).value  # Пароль

# Работа с листом vkurse_stats
sheet_vk_stats = exl["vkurse_stats"]

# Работа с листом vkurse_chat
# sheet_vk_chat = exl["vkurse_chat"]

# Работа с листом vkurse_names
sheet_vk_names = exl["vkurse_names"]

# Инициализация WebDriver
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Установка пути для загрузки файлов
download_dir = os.path.join(os.getcwd(), "downloads")  # Текущая папка для загрузок
os.makedirs(download_dir, exist_ok=True)  # Создаем папку, если она не существует

# Настройка ChromeOptions для изменения папки загрузки
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,  # Папка для загрузок
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Инициализация WebDriver с опциями
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Вход на сайт
driver.get("https://edu.vkurse.ru/v2/login")

# Кнопка Входа Synergy ID
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "Войти через Synergy ID"))
)
login_button.click()

time.sleep(4)  # Даем время странице загрузиться после нажатия на кнопку входа

# Ввод логина
username_field = driver.find_element(By.NAME, 'username')  # Updated to correct ID
username_field.send_keys(login_exl)

time.sleep(2)  # Задержка перед вводом пароля (опционально)

# Ввод пароля
password_field = driver.find_element(By.NAME, 'password')  # Updated to correct ID
password_field.send_keys(password_exl)

# Кнопка Входа
login_button2 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'kc-login'))  # Corrected to use EC.element_to_be_clickable
)
login_button2.click()

# Пропустить адрес после входа
driver.get("https://edu.vkurse.ru/v2/iva/home/chats?conferenceSessionOpenMode=PIP&currentChatId=21f6df6a-c56d-3d18-a630-e2afa279831f")

# Перебор всех строк на листе vkurse_stats (новый интерфейс)
for row in range(2, sheet_vk_stats.max_row + 1):  # Начнем со строки 2
    vkurse_stats_exl = sheet_vk_stats.cell(row=row, column=1).value  # Чтение ссылки
    driver.get(f"{vkurse_stats_exl}")

    # Скачивание файлов
    try:
        time.sleep(10)  # Добавлено ожидание 5 секунд перед выполнением блока try

        # Кнопка ОК
        ok_button = driver.find_element(By.XPATH, "//button[contains(@class, 'iva-button') and contains(., 'OK')]")
        ok_button.click()

        # Кнопка ОК
        # ok_button = driver.find_element(By.XPATH, "//button[contains(@class, 'iva-button')]")
        # ok_button.click()

        # Обновленный поиск кнопки для скачивания
        download_button_section0 = driver.find_element(By.CLASS_NAME, "iva-icon-button.relative.ng-star-inserted")
        download_button_section0.click()  # Переход к скачиванию

        # Обновленный поиск кнопки для экспорта отчета
        download_button_section = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Экспортировать отчёт об участниках')]"))
        )
        download_button_section.click()  # Переход к скачиванию

        # Ожидание завершения загрузки
        time.sleep(5)  # Задержка для завершения загрузки (можно улучшить с помощью проверки)

        # Получение имени файла .csv
        original_file_name = glob.glob(os.path.join(download_dir, "*.csv"))[0]  # Получаем первый файл .csv
        new_file_name = os.path.join(os.getcwd(),
                                     sheet_vk_names.cell(row=row, column=1).value + ".csv")  # Новое имя файла

        # Переименование и перемещение файла
        shutil.move(original_file_name, new_file_name)

    except Exception as e:
        print(f"Не удалось получить XPATH для задачи {vkurse_stats_exl}: {e}")

# Закрыть браузер
input("Скачивание файлов .csv завершено. Нажмите Enter для закрытия окна")
driver.quit()

# softy_plug