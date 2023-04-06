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
from common.store_column_enum import CommonStoreEnum, Cafe24Enum, ElevenStreetEnum

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains

from datetime import timedelta, datetime
import time

import pandas as pd
from openpyxl import load_workbook

import re


class EzadminCrawlerProcess:
    def __init__(self):
        open_browser()
        self.default_wait = 10
        self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        self.driver.implicitly_wait(self.default_wait)
        self.driver.maximize_window()

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg = log_msg

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

    def get_store_list(self):
        # header 옵션에 2를 넣어서 3행부터 읽기 시작합니다.
        # 반드시 3행에 쇼핑몰의 이름이 위치해야 합니다.
        df_stats: pd.DataFrame = pd.read_excel(
            self.guiDto.stats_file, sheet_name=self.guiDto.sheet_name, keep_default_na="", header=2
        )
        store_list = []
        for column in df_stats.columns:
            if column.find("Unnamed") <= -1 and column.find("\n") <= -1 and column.find("합계") <= -1:
                store_list.append(column)
        print(store_list)
        self.log_msg.emit(f"{store_list}가 발견되었습니다.")
        return store_list

    def get_store_column_range(self, store_name):
        merged_cells = self.sheet.merged_cells

        for merged_cell in merged_cells:
            if merged_cell.start_cell.internal_value == store_name:
                print(merged_cell.start_cell.internal_value)
                # store_column_range = merged_cell.coord
                store_min_col = merged_cell.min_col
                store_max_col = merged_cell.max_col
                store_column_range = [store_min_col, store_max_col]
                break

        # store_column_range = re.sub(r"\d+", "", store_column_range)
        print(store_column_range)
        return store_column_range

    def get_target_date_row(self, target_date):
        # 순회할 셀 범위 지정
        min_row, max_row = 1, self.sheet.max_row
        min_col, max_col = 1, self.sheet.max_column

        for row in self.sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                if target_date in str(cell.value):
                    print(f"'{target_date}'이 포함된 셀 위치: ({cell.row}, {cell.column})")
                    return cell.row

    def ezadmin_login(self):
        driver = self.driver
        self.driver.get(self.dict_accounts["이지어드민"]["URL"])
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//body[@class="ezadmin-main-body"]'))
        )
        time.sleep(0.2)

        login_domain = self.dict_accounts["이지어드민"]["도메인"]
        login_id = self.dict_accounts["이지어드민"]["ID"]
        login_pw = self.dict_accounts["이지어드민"]["PW"]

        # 로그인 시도
        # 이 행위 중 하나라도 실패한다면 로그인 실패
        try:
            driver.implicitly_wait(2)

            open_button = driver.find_element(By.XPATH, '//a[./span[@class="img_login"]][contains(text(), "로그인")]')
            driver.execute_script("arguments[0].click();", open_button)
            time.sleep(0.2)

            domain_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-domain"]')
            domain_input.clear()
            domain_input.send_keys(login_domain)
            time.sleep(0.2)

            id_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-id"]')
            id_input.clear()
            id_input.send_keys(login_id)
            time.sleep(0.2)

            pwd_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-pwd"]')
            pwd_input.clear()
            pwd_input.send_keys(login_pw)
            time.sleep(0.2)

            save_domain = driver.find_element(By.XPATH, '//input[@id="savedomain"]')
            driver.execute_script("arguments[0].click();", save_domain)
            time.sleep(0.2)

            login_button = driver.find_element(By.XPATH, '//input[@class="login-btn" and @value="로그인"]')
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(0.2)

            # 로그인 성공 시 나오는 화면
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//body[@class="bgline"]')))
            time.sleep(0.2)

            self.close_ezadmin_notice_popups()

        except Exception as e:
            print(e)
            raise Exception(f"이지어드민 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

    # 이지어드민 로그인 시 발생하는 팝업창을 모두 닫습니다.
    def close_ezadmin_notice_popups(self):
        driver = self.driver
        try:
            driver.implicitly_wait(1)

            driver.execute_script("hide_board('internal_board');")
            time.sleep(0.2)

            driver.execute_script("hide_board('sys_notice_board');")
            time.sleep(0.2)

            # $x('//a[contains(text(), "팝업 전체 닫기")]')
            close_all_popups = driver.find_element(By.XPATH, '//a[contains(text(), "팝업 전체 닫기")]')
            driver.execute_script("arguments[0].click();", close_all_popups)
            time.sleep(0.2)

        except Exception as e:
            print(e)

        finally:
            driver.implicitly_wait(self.default_wait)

    # 정산통계 -> 판매처별정산통계 화면으로 이동합니다.
    def go_store_stats_menu(self):
        driver = self.driver
        driver.get("https://ga20.ezadmin.co.kr/template35.htm?template=F308")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "판매처별정산통계")]'))
        )
        time.sleep(0.1)

    # 전체작업 시작
    def work_start(self):
        print(f"process: work_start")

        try:
            # 계정 엑셀 파일
            self.dict_accounts = self.get_dict_account()

            # 통계 엑셀 파일에서 상점 이름을 추출합니다.
            store_list = self.get_store_list()

            self.workbook = load_workbook(self.guiDto.stats_file)

            self.sheet = self.workbook[self.guiDto.sheet_name]

            target_date_row = self.get_target_date_row(self.guiDto.target_date)

            self.ezadmin_login()

            for store_name in store_list:
                print(f"{store_name} 작업 시작")
                self.log_msg.emit(f"{store_name} 작업 시작")

                try:
                    store_column_range = self.get_store_column_range(store_name)

                    self.go_store_stats_menu()

                    print()

                except Exception as e:
                    print(str(e))
                    self.log_msg.emit(f"{store_name} 작업 실패")
                    global_log_append(str(e))
                    continue

        except Exception as e:
            print(str(e))
            if str(e).find("채널명") > -1:
                self.log_msg.emit(f"계정 엑셀 파일 양식이 아닙니다.")
            else:
                self.log_msg.emit(f"{str(e)}")

        finally:
            # self.driver.close()
            self.workbook.close()
            time.sleep(0.2)


if __name__ == "__main__":
    process = EzadminCrawlerProcess()
    process.work_start()
