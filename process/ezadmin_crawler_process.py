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

            # 통계 엑셀 파일에서 상점 이름을 추출합니다.
            store_list = self.get_store_list()

            self.workbook = load_workbook(self.guiDto.stats_file)

            self.sheet = self.workbook[self.guiDto.sheet_name]

            for store_name in store_list:
                print(store_name)
                merged_cells = self.sheet.merged_cells

                for merged_cell in merged_cells:
                    if merged_cell.start_cell.internal_value == store_name:
                        print(merged_cell.start_cell.internal_value)
                        store_column_range = merged_cell.coord
                        break

                print(store_column_range)
                print()

        except Exception as e:
            print(e)
            if str(e).find("채널명") > -1:
                self.log_msg.emit(f"계정 엑셀 파일 양식이 아닙니다.")

        finally:
            # self.driver.close()
            self.workbook.close()
            time.sleep(0.2)


if __name__ == "__main__":
    process = EzadminCrawlerProcess()
    process.work_start()
