from urllib.request import urlretrieve
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import openpyxl
from datetime import datetime

scene_list=['정상 시나리오','아이디 입력란 공백','PW 입력란 공백'] #시나리오 리스트
inform_list=['단계','시나리오','날짜','합격 여부','오류 코드'] #각 시나리오에 대한 테스트케이스 수행 정보
exl_file = openpyxl.Workbook()  # 엑셀 보고서 작성
exl_sheet = exl_file.active
exl_sheet.append(inform_list)
exl_sheet.column_dimensions['B'].width=20
exl_sheet.column_dimensions['C'].width=40
exl_sheet.column_dimensions['E'].width=105

chromedriver='C:\devpython\Webdriver\chromedriver.exe'
tester = webdriver.Chrome(service=Service(chromedriver)) #브라우저 : Chrome

def signup_normal(scene : int): #구분 값
    now = datetime.now()
    inform=[scene,scene_list[scene],now,str(False)] #테스트 수행 정보

    try:
        tester.get('http://semtle.catholic.ac.kr') #해당 사이트 테스트

        tester.maximize_window() #화면 크기 최대
        tester.implicitly_wait(10)

        login_btn=tester.find_element(By.XPATH,"//div[@class='member-header']/h3/a") #로그인 버튼
        login_btn.click()
        tester.implicitly_wait(10)

        signup_btn=tester.find_element(By.CSS_SELECTOR,'#loginform > div.signupform > a')
        signup_btn.click()
        tester.implicitly_wait(10)

        #데이터 입력
        test_data=['201500011','t1234','t1234','Gorani','HongGilDong','201500011','test@naver.com','01012341234']
        inputs=tester.find_elements(By.CSS_SELECTOR,'div.join > form > div.join_input > input')
        dropdown=Select(tester.find_element(By.CSS_SELECTOR,'div.join > form > div.join_input > select'))

        for idx,input_box in enumerate(inputs):
            input_box.send_keys(test_data[idx])
            time.sleep(2)
        dropdown.select_by_value('4')
        time.sleep(2)

        tester.execute_script('window.scrollTo(0,document.body.scrollHeight)')

        chkbox=tester.find_element(By.CSS_SELECTOR,'#marketing')
        chkbox.click()
        time.sleep(2)

        rsign_btn=tester.find_element(By.CSS_SELECTOR,'#sign > div.join_input.join_btnbox > button')
        rsign_btn.click()

        try:
            alert_box=tester.switch_to.alert
            time.sleep(1)
            alert_box.accept()
        except:
            pass

        tester.implicitly_wait(10)
        time.sleep(2)

        inform[-1]=str(True) #

    except Exception as e:
        inform.append(e)

    exl_sheet.append(inform) #해당 시나리오 수행 정보 저장
    tester.quit()

def create_report():
    print('create report...')
    exl_file.save('semtle_signup_test.xlsx')
    exl_file.close()
    print('complete!')


signup_normal(0)
create_report()