if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.gui_dto import GUIDto

from common.utils import global_log_append
from common.chrome import open_browser, get_chrome_driver
from common.selenium_activities import close_new_tabs, alert_ok_try
from common.account_file import AccountFile

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains

from datetime import timedelta, datetime
import time

import pandas as pd


class EzadminCrawlerProcess:
    def __init__(self):
        self.default_wait = 10
        open_browser()
        self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        self.driver.implicitly_wait(self.default_wait)
        self.driver.maximize_window()

    def get_dict_account(self):
        df_accounts = AccountFile(self.guiDto.account_file).df_account
        df_accounts = df_accounts.fillna("")
        dict_accounts = {}
        for index, row in df_accounts.iterrows():
            channel = str(row["채널명"])
            domain = str(row["도메인"])
            account_id = str(row["ID"])
            account_pw = str(row["PW"])
            url = str(row["URL"])
            dict_accounts[channel] = {"도메인": domain, "ID": account_id, "PW": account_pw, "URL": url}
        return dict_accounts

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg = log_msg

    def login(self, user_id: str, user_pw: str):
        driver = self.driver
        driver.get(f"https://ylkorea1.cafe24.com/member/login.html")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "기존 회원 로그인")]'))
        )
        time.sleep(0.2)

        input_id = driver.find_element(By.CSS_SELECTOR, 'div[class="inputBox"] input[id="member_id"]')
        input_id.send_keys(user_id)
        time.sleep(0.2)

        input_pw = driver.find_element(By.CSS_SELECTOR, 'div[class="inputBox"] input[id="member_passwd"]')
        input_pw.send_keys(user_pw)
        time.sleep(0.2)

        driver.find_element(By.XPATH, '//button[contains(text(), "기존 회원 로그인")]').click()
        time.sleep(0.2)

        # alert 발생
        # 장바구니에 품목을 추가 하세요
        alert_msg = ""
        try:
            driver.implicitly_wait(1)
            alert = driver.switch_to.alert
            alert_msg = alert.text
            alert.accept()
        except Exception as e:
            print(f"no login alert")
            pass
        finally:
            driver.implicitly_wait(self.default_wait)

        if alert_msg.find("아이디 또는 비밀번호가 일치하지 않습니다") > -1:
            print(f"alert_msg: {alert_msg}")
            self.log_msg.emit(f"{user_id}, {user_pw} 로그인에 실패했습니다. 사유: {alert_msg}")
            raise Exception(f"{user_id}, {user_pw} 로그인에 실패했습니다. 사유: {alert_msg}")

    # 전체작업 시작
    def work_start(self):
        print(f"process: work_start")

        try:
            # 계정 엑셀파일 검사
            self.dict_accounts = self.get_dict_account()

        except Exception as e:
            print(e)
            if str(e).find("채널명") > -1:
                self.log_msg.emit(f"계정 엑셀 파일 양식이 아닙니다.")

        finally:
            self.driver.close()
            time.sleep(0.2)


if __name__ == "__main__":
    process = EzadminCrawlerProcess()
    process.work_start()
