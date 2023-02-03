from urllib.request import urlretrieve
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

import openpyxl
from datetime import datetime

scene_list = ['정상 시나리오(대출, 반납)']  # 시나리오 리스트
inform_list = ['단계', '테스트 케이스', '입력 값','기대 값','날짜', '합격 여부', '오류 코드','비고']  # 각 시나리오에 대한 테스트케이스 수행 정보
exl_file = openpyxl.Workbook()  # 엑셀 보고서 작성
exl_sheet = exl_file.active
exl_sheet.append(inform_list)
exl_sheet.column_dimensions['B'].width = 20
exl_sheet.column_dimensions['C'].width = 40
exl_sheet.column_dimensions['D'].width = 40
exl_sheet.column_dimensions['E'].width = 20
exl_sheet.column_dimensions['G'].width = 105


def booking_normal(scene: int,data_num : int):  # 구분 값
    now = datetime.now()
    expected_val = ['대출 및 반납 성공','대출 및 반납 성공','대출 및 반납 성공','내역 없음']
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
        time.sleep(1)

        bookbtn=tester.find_element(By.CSS_SELECTOR,'body > section.home > header > nav > a:nth-child(4)')
        bookbtn.click()
        bookinput=tester.find_element(By.CSS_SELECTOR,'#bookSearch > input')

        search_data=['C','자료','','Empty']
        inform = [data_num, scene_list[scene], search_data[data_num], expected_val[data_num], now, str(False)]  # 테스트 수행 정보
        bookinput.send_keys(search_data[data_num])
        bookinput.send_keys(Keys.RETURN)
        sections=tester.find_elements(By.CSS_SELECTOR,'section.takeoutbox div.takeoutbox-right div.takeoutbox-btn button.takeout_btn')
        if len(sections)==0:
            pass
        else:
            time.sleep(1)
            sections[0].click()
            try:
                alert_box = tester.switch_to.alert
                time.sleep(1)
                alert_box.accept()
            except Exception as e:
                inform.append(e)
                exl_sheet.append(inform)  # 해당 시나리오 수행 정보 저장
                tester.close()
                return
            tester.implicitly_wait(10)
            time.sleep(1)
            bookinput = tester.find_element(By.CSS_SELECTOR, '#bookSearch > input')
            bookinput.send_keys(search_data[data_num])
            time.sleep(1)
            bookinput.send_keys(Keys.RETURN)
            sections = tester.find_elements(By.CSS_SELECTOR,
                                          'section.takeoutbox div.takeoutbox-right div.takeoutbox-btn button.takeout_btn')
            sections[0].click()
            try:
                alert_box = tester.switch_to.alert
                time.sleep(1)
                alert_box.accept()
            except Exception as e:
                inform.append(e)
                exl_sheet.append(inform)  # 해당 시나리오 수행 정보 저장
                tester.close()
                return

        inform[-1] = str(True)

    except Exception as e:
        inform.append(e)

    tester.implicitly_wait(10)
    exl_sheet.append(inform)  # 해당 시나리오 수행 정보 저장
    tester.close()


def scene_info_save(sheet, tester, info):
    sheet.append(info)
    tester.quit()


def create_report():
    print('create report...')
    exl_file.save('semtle_book_test.xlsx')
    exl_file.close()
    print('complete!')


testcases = [booking_normal]

for num, case in enumerate(testcases):
    for data in range(4):
        case(num,data)

create_report()