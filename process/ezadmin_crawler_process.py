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
from openpyxl import load_workbook


class EzadminCrawlerProcess:
    def __init__(self):
        self.default_wait = 10
        # open_browser()
        # self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        # self.driver.implicitly_wait(self.default_wait)
        # self.driver.maximize_window()

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

    def get_df_stats(self):
        # header 옵션에 2를 넣어서 3행부터 시작합니다.
        # 3행에는 각 쇼핑몰의 이름이 들어있습니다.
        df_stats: pd.DataFrame = pd.read_excel(
            self.guiDto.stats_file, sheet_name=self.guiDto.sheet_name, keep_default_na="", header=2
        )
        store_list = []
        for column in df_stats.columns:
            if column.find("Unnamed") <= -1 and column.find("\n") <= -1 and column.find("합계") <= -1:
                store_list.append(column)
        print(store_list)

        workbook = load_workbook(self.guiDto.stats_file)

        sheet = workbook[self.guiDto.sheet_name]

        merged_cells = sheet.merged_cells
        # print(merged_cells)

        value = store_list[0]

        # 병합된 셀의 범위를 순회하며 입력된 문자(value)를 입력
        for merged_cell in merged_cells:
            if merged_cell.start_cell.internal_value == value:
                print(merged_cell.start_cell.internal_value)
                store_range = merged_cell.coord
                break

        print(store_range)
        print()

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg = log_msg

    def login(self, user_id: str, user_pw: str):
        driver = self.driver
        driver.get(f"https://ylkorea1.cafe24.com/member/login.html")
        time.sleep(0.2)

    # 전체작업 시작
    def work_start(self):
        print(f"process: work_start")

        try:
            # 계정 엑셀 파일
            self.dict_accounts = self.get_dict_account()

            # 통계 엑셀 파일
            self.df_stats = self.get_df_stats()

        except Exception as e:
            print(e)
            if str(e).find("채널명") > -1:
                self.log_msg.emit(f"계정 엑셀 파일 양식이 아닙니다.")

        finally:
            # self.driver.close()
            time.sleep(0.2)


if __name__ == "__main__":
    process = EzadminCrawlerProcess()
    process.work_start()
