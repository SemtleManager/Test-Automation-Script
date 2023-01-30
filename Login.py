from urllib.request import urlretrieve
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

import openpyxl
from datetime import datetime

scene_list = ['정상 시나리오', 'ID 입력 오류', 'PW 입력 오류']  # 시나리오 리스트
inform_list = ['단계', '시나리오', '날짜', '합격 여부', '오류 코드']  # 각 시나리오에 대한 테스트케이스 수행 정보
exl_file = openpyxl.Workbook()  # 엑셀 보고서 작성
exl_sheet = exl_file.active
exl_sheet.append(inform_list)
exl_sheet.column_dimensions['B'].width = 20
exl_sheet.column_dimensions['C'].width = 40
exl_sheet.column_dimensions['E'].width = 105


def login_normal(scene: int):  # 구분 값
    now = datetime.now()
    inform = [scene, scene_list[scene], now, str(False)]  # 테스트 수행 정보
    chromedriver = 'C:\devpython\Webdriver\chromedriver.exe'
    tester = webdriver.Chrome(service=Service(chromedriver))  # 브라우저 : Chrome

    try:
        tester.get('http://semtle.catholic.ac.kr')  # 해당 사이트 테스트

        tester.maximize_window()  # 화면 크기 최대
        tester.implicitly_wait(10)

        login_btn = tester.find_element(By.XPATH, "//div[@class='member-header']/h3/a")  # 로그인 버튼
        login_btn.click()

        tester.implicitly_wait(10)
        time.sleep(1)
        id_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.idform > input')
        pw_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.pwform > input')
        rlogin_btn = tester.find_element(By.CSS_SELECTOR, '#loginform > div:nth-child(5) > input')

        id_input.send_keys('201500011')
        pw_input.send_keys('t1234')
        time.sleep(1)

        rlogin_btn.click()

        tester.implicitly_wait(10)
        time.sleep(2)
        inform[-1] = str(True)

    except Exception as e:
        inform.append(e)

    exl_sheet.append(inform)  # 해당 시나리오 수행 정보 저장
    tester.close()


def login_emptyID(scene: int):  # 구분 값
    now = datetime.now()
    inform = [scene, scene_list[scene], now, str(False)]  # 테스트 수행 정보
    chromedriver = 'C:\devpython\Webdriver\chromedriver.exe'
    tester = webdriver.Chrome(service=Service(chromedriver))  # 브라우저 : Chrome

    try:
        tester.get('http://semtle.catholic.ac.kr')  # 해당 사이트 테스트

        tester.maximize_window()  # 화면 크기 최대
        tester.implicitly_wait(10)

        login_btn = tester.find_element(By.XPATH, "//div[@class='member-header']/h3/a")  # 로그인 버튼
        login_btn.click()

        tester.implicitly_wait(10)
        time.sleep(1)
        id_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.idform > input')
        pw_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.pwform > input')
        rlogin_btn = tester.find_element(By.CSS_SELECTOR, '#loginform > div:nth-child(5) > input')

        id_input.send_keys('')
        pw_input.send_keys('t1234')
        time.sleep(1)

        rlogin_btn.click()
        try:
            alert = tester.switch_to.alert
            time.sleep(1)
            alert.accept()
        except Exception as e:
            inform.append(e)
            scene_info_save(exl_sheet, tester, inform)
            return

        tester.implicitly_wait(10)
        print(tester.current_url)
        time.sleep(2)
        inform[-1] = str(True)

    except Exception as e:
        inform.append(e)

    scene_info_save(exl_sheet, tester, inform) #시나리오 저장

def login_emptyPW(scene: int):  # 구분 값
    now = datetime.now()
    inform = [scene, scene_list[scene], now, str(False)]  # 테스트 수행 정보
    chromedriver = 'C:\devpython\Webdriver\chromedriver.exe'
    tester = webdriver.Chrome(service=Service(chromedriver))  # 브라우저 : Chrome

    try:
        tester.get('http://semtle.catholic.ac.kr')  # 해당 사이트 테스트

        tester.maximize_window()  # 화면 크기 최대
        tester.implicitly_wait(10)

        login_btn = tester.find_element(By.XPATH, "//div[@class='member-header']/h3/a")  # 로그인 버튼
        login_btn.click()

        tester.implicitly_wait(10)
        time.sleep(1)
        id_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.idform > input')
        pw_input = tester.find_element(By.CSS_SELECTOR, '#loginform > div.pwform > input')
        rlogin_btn = tester.find_element(By.CSS_SELECTOR, '#loginform > div:nth-child(5) > input')

        id_input.send_keys('201500011')
        pw_input.send_keys('')
        time.sleep(1)

        rlogin_btn.click()
        try:
            alert = tester.switch_to.alert
            time.sleep(1)
            alert.accept()
        except Exception as e:
            inform.append(e)
            scene_info_save(exl_sheet, tester, inform)
            return

        tester.implicitly_wait(10)
        print(tester.current_url)
        time.sleep(2)
        inform[-1] = str(True)

    except Exception as e:
        inform.append(e)

    scene_info_save(exl_sheet, tester, inform) #시나리오 저장


def scene_info_save(sheet, tester, info):
    sheet.append(info)
    tester.quit()


def create_report():
    print('create report...')
    exl_file.save('semtle_login_test.xlsx')
    exl_file.close()
    print('complete!')


testcases = [login_normal, login_emptyID, login_emptyPW]

for num, case in enumerate(testcases):
    case(num)
    time.sleep(1)

create_report()
