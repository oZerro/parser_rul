import openpyxl
import time
from openpyxl import Workbook
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC



def get_table(path_table):
    workbok = openpyxl.load_workbook(path_table)
    sheet = workbok.active

    numbers = sheet['A']
    seria = sheet['B']
    date = sheet['C']

    ishod = []

    for num, cell in enumerate(seria):
        rez = []

        if cell.value is None:
            break
        if num > 0:
            if numbers[num].value is not None:
                rez.append(int(numbers[num].value))
            if seria[num].value is not None:
                rez.append(int(seria[num].value))
            if date[num].value is not None:
                dat = str(date[num].value)
                dat = dat.split()[0].split('-')[::-1]
                dat[2] = dat[2][2:]
                dat = '.'.join(dat)
                rez.append(dat)
            ishod.append(rez)

    return ishod


def check_value_date(date, in_date):
    date = date.split('.')
    date[2] = f'20{date[2]}'
    date = '.'.join(date)

    if date == in_date:
        return True
    
    return False


def split_list(lst, num_parts):
    avg = len(lst) // num_parts
    remainder = len(lst) % num_parts
    result = []
    start = 0

    for i in range(num_parts):
        end = start + avg + (1 if i < remainder else 0)
        result.append(lst[start:end])
        start = end

    return result


def check_element_for_page(driver, css_selector):
    wait = WebDriverWait(driver, 10)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
        return True
    except TimeoutException:
        return False
    

def check_element_for_tag(driver, tag):
    try:
        driver.find_element(By.TAG_NAME, tag)
        return True
    except NoSuchElementException:
        return False
    

def check_style(driver, css_selector, styles_to_check:dict):
    #css_selector = 'your_css_selector'
    #styles_to_check = {'color': 'red', 'font-size': '16px'}
    # Формирование JavaScript-кода для проверки стилей
    script = f"""
    var element = document.querySelector('{css_selector}');
    var styles = window.getComputedStyle(element);
    """

    # Добавление условий для каждого стиля
    for style, value in styles_to_check.items():
        script += f"if (styles.{style} !== '{value}') return false;"

    # Возвращение true, если все стили соответствуют
    script += "return true;"

    # Выполнение JavaScript-кода
    result = driver.execute_script(script)

    # Проверка результата
    return result


def go_home_page(driver, home_page_url):
    while True:
        driver.get(home_page_url)
        time.sleep(2)
        if driver.current_url == home_page_url:
            print("Мы на главной странице")
            break
        continue


