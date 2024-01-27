import time
import traceback
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from PIL import Image
from dotenv import load_dotenv
from twocaptcha import TwoCaptcha
from openpyxl import Workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from test import *
from multiprocessing import Pool





def main(table):
    load_dotenv()
    options = Options()
    options.add_argument("--headless")
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("--mute-audio")

    token = os.getenv("TOKEN")
    solver = TwoCaptcha(token)
    url = "https://xn--90adear.xn--p1ai/check/driver#+"
    driver = webdriver.Chrome(options=options)
    driver.get(url)
    time.sleep(2)
    home_page = driver.current_url
    num_prav = 0

    number_exept = 0
    num_prav = 0
    arr = []
    while num_prav < len(table):
        wait = WebDriverWait(driver, 10)
        try: 
            dan = table[num_prav]

            rez = []
            for d in dan:
                rez.append(d)

            seria = dan[1]
            date = dan[2]

            
            while True:
                print("Указываю серию")
                input_num = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="num"]')))
                input_num.clear()
                input_num.send_keys(seria)
                value_seria = input_num.get_attribute('value')

                if str(value_seria) == str(seria):
                    print("Указал серию")
                    break
                else:
                    continue
            
            
            while True:
                print("Указываю дату")
                date_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="date"]')))
                date_input.clear()
                date_input.send_keys(date)

                value_date = date_input.get_attribute('value')
                
                if check_value_date(date, value_date):
                    print("Указал дату")
                    break
                
                date_input.clear()
                continue


            while True:
                print("Начинаю решение капчи")
                check_but = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[class="checker"]')))
                check_but.click()

                capcha_pic = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[id="captchaDialog"]')))
                img_capcha = capcha_pic.find_element(By.TAG_NAME, 'img')

                input_num_catcha = capcha_pic.find_element(By.CSS_SELECTOR, 'input[name="captcha_num"]')
                count_time = 0
                while True:
                    if count_time > 60:
                        break
                    print("Ищу картинку капчи")
                    src_img = img_capcha.get_attribute('src')
                    if 'jpeg' in src_img:
                        img_capcha.screenshot('cropped_screenshot.png')
                        print("Нашел картинку")
                        break
                    else:
                        time.sleep(1)
                        count_time += 1
                        continue


                file_path = "cropped_screenshot.png"

                with open(file_path, 'rb') as image_file:
                    # отправляем запрос на решение капчи
                    result = solver.normal(file_path)

                input_num_catcha.clear()
                input_num_catcha.send_keys(result['code'])
                time.sleep(1)

                if check_style(driver, '#captchaFade', {'display': 'none'}):
                    print("Прошел капчу")
                    time.sleep(2)
                    if check_element_for_page(driver, '.timestamp'):
                        break
                    else:
                        driver.refresh()
                        go_home_page(driver, home_page)
                        time.sleep(2)
                else:
                    capcha_pic.find_element(By.CSS_SELECTOR, '#captchaCancel').click()
                    print("Не верно решил пробую еще раз")
                    continue
                
                    


            if check_element_for_page(driver, '.adds-modal'):
                if check_element_for_tag(driver, 'video'):
                    while True:
                        time.sleep(5)
                        if check_style(driver, '.close_modal_window', {'display': 'block'}):
                            driver.find_element(By.CSS_SELECTOR, '.close_modal_window').click()
                            time.sleep(2)
                            break
                        else:
                            continue
                else:
                    time.sleep(10)


            if check_element_for_page(driver, '.decis-item'):
                decis = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.decis-item')))
                fields = decis.find_elements(By.CLASS_NAME, 'field')
                rez.append(fields[2].text.split()[0])
                rez.append(fields[1].text)
                print(num_prav + 1, "Лишен, собрал данные")
            else:
                rez.append("-")
                rez.append("-")
                print(num_prav + 1, "Не лишен, собрал данные")


            arr.append(rez)
            print(number_exept, "ошибок")

            

            
            driver.refresh()
            print(num_prav + 1, "Успешно прошел, иду дальше")
            print()

            go_home_page(driver, home_page)
            num_prav += 1
            if num_prav == len(table):
                break
        except Exception as ex:
            traceback.print_exc()
            print("Что-то пошло не так, иду на повторный круг")
            number_exept += 1
            driver.refresh()
            go_home_page(driver, home_page)
            continue

        

    for i in arr:
        print(i)

    driver.close()
    driver.quit()

    return arr





if __name__ == "__main__":

    wb = Workbook()
    ws = wb.active
    ws.append(['Номер заявки', 'Номер ВУ', 'Дата выдачи', 'Срок лишения', 'Дата лишения'])

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25

    num_process = int(input("Укажите количество процессов: "))

    p = Pool(processes=num_process)

    start_time = time.time()
    
    table = get_table('ishodnik.xlsx')
    tables = split_list(table, num_process)

    rezults = p.map(main, tables)
    
    for i in rezults:
        print(i)

    for i in range(len(rezults)):
        for k in range(len(rezults[i])):
            ws.append(rezults[i][k])
            print(rezults[i][k])
        
    wb.save('rezult.xlsx')

    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Время выполнения программы: {execution_time} секунд")

