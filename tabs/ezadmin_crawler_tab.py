import sys
import warnings

warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from datetime import *

from threads.ezadmin_crawler_thread import EzadminCrawlerThread
from dtos.gui_dto import GUIDto
from common.utils import *

from common.account_file import AccountFile
import pandas as pd

from configs.ezadmin_crawler_config import EzadminCrawlerConfig as Config
from configs.ezadmin_crawler_config import EzadminCrawlerData as ConfigData

from common.chrome import open_browser, get_chrome_driver
from selenium import webdriver


class EzadminCrawlerTab(QWidget):
    # 초기화
    def __init__(self):
        self.config = Config()
        __saved_data = self.config.get_data()
        self.saved_data = self.config.dict_to_data(__saved_data)

        super().__init__()
        self.initUI()

    # 로그 작성
    @pyqtSlot(str)
    def log_append(self, text):
        today = str(datetime.now())[0:10]
        now = str(datetime.now())[0:-7]
        self.browser.append(f"[{now}] {str(text)}")
        global_log_append(text)

    # 크롬 브라우저
    def open_chrome_browser(self):
        open_browser()

    # 시작 클릭
    def crawler_start_button_clicked(self):
        if self.account_file.text() == "":
            QMessageBox.information(self, "작업 시작", f"계정 엑셀 파일을 선택해주세요.")
            return
        else:
            account_file = self.account_file.text()

        if not os.path.isfile(account_file):
            QMessageBox.information(self, "작업 시작", f"엑셀 경로가 잘못되었습니다.")
            return

        if self.stats_file.text() == "":
            QMessageBox.information(self, "작업 시작", f"통계 엑셀 파일을 선택해주세요.")
            return
        else:
            stats_file = self.stats_file.text()

        if not os.path.isfile(stats_file):
            QMessageBox.information(self, "작업 시작", f"엑셀 경로가 잘못되었습니다.")
            return

        if self.sheet_combobox.currentText() == "":
            print(f"시트를 선택해주세요.")
            QMessageBox.information(self, "작업 시작", f"시트를 선택해주세요.")
            self.log_append(f"시트를 선택해주세요.")
            return

        selected_date_list = []
        self.date_listwidget.selectAll()
        if len(self.date_listwidget.selectedItems()) <= 0:
            print(f"선택된 날짜가 없습니다.")
            QMessageBox.information(self, "날짜 선택", f"선택된 날짜가 없습니다.")
            return

        else:
            date_list_items = self.date_listwidget.selectedItems()
            for date_item in date_list_items:
                selected_date_list.append(date_item.text())

        guiDto = GUIDto()
        guiDto.account_file = account_file
        guiDto.stats_file = stats_file
        guiDto.sheet_name = self.sheet_combobox.currentText()
        guiDto.target_date_list = selected_date_list

        self.crawler_thread = EzadminCrawlerThread()
        self.crawler_thread.log_msg.connect(self.log_append)
        self.crawler_thread.crawler_finished.connect(self.crawler_finished)
        self.crawler_thread.setGuiDto(guiDto)

        self.crawler_start_button.setDisabled(True)
        self.crawler_stop_button.setDisabled(False)
        self.crawler_thread.start()

    # 중지 클릭
    @pyqtSlot()
    def crawler_stop_button_clicked(self):
        print(f"search stop clicked")
        self.log_append(f"중지 클릭")
        self.crawler_finished()

    # 작업 종료
    @pyqtSlot()
    def crawler_finished(self):
        print(f"search thread finished")
        self.log_append(f"작업 종료")
        self.crawler_thread.stop()
        self.crawler_start_button.setDisabled(False)
        self.crawler_stop_button.setDisabled(True)
        print(f"thread_is_running: {self.crawler_thread.isRunning()}")

    def crawler_save_button_clicked(self):
        dict_save = {"account_file": self.account_file.text(), "stats_file": self.stats_file.text()}

        question_msg = "현재 상태를 저장하시겠습니까?"
        reply = QMessageBox.question(self, "상태 저장", question_msg, QMessageBox.Yes, QMessageBox.No)

        if reply == QMessageBox.Yes:
            print(f"저장")
            self.config.write_data(dict_save)
        else:
            print(f"저장 취소")

    def account_file_select_button_clicked(self):
        print(f"excel file select")
        file_name = QFileDialog.getOpenFileName(self, "", "", "excel file (*.xlsx)")

        if file_name[0] == "":
            print(f"선택된 파일이 없습니다.")
            return

        print(file_name[0])
        self.account_file.setText(file_name[0])

    def stats_file_select_button_clicked(self):
        print(f"excel file select")
        file_name = QFileDialog.getOpenFileName(self, "", "", "excel file (*.xlsx)")

        if file_name[0] == "":
            print(f"선택된 파일이 없습니다.")
            return

        print(file_name[0])
        self.stats_file.setText(file_name[0])

    def stats_file_textChanged(self):
        self.sheet_combobox.clear()
        try:
            excel_file = pd.ExcelFile(self.stats_file.text())
            sheet_list = excel_file.sheet_names
        except Exception as e:
            sheet_list = []
        self.set_sheet_combobox(sheet_list)

    def set_sheet_combobox(self, sheet_list):
        for sheet in sheet_list:
            self.sheet_combobox.addItem(sheet)

    def add_date_button_clicked(self):
        target_date = self.date_edit.text()
        print(target_date)
        self.date_listwidget.addItem(target_date)

    def remove_date_button_clicked(self):
        print("click")
        selected_items = self.date_listwidget.selectedItems()
        if not selected_items:
            return
        for item in selected_items:
            self.date_listwidget.takeItem(self.date_listwidget.row(item))

    # 메인 UI
    def initUI(self):
        # 대상 시트 선택
        sheet_setting_groupbox = QGroupBox("시트 선택")
        self.sheet_combobox = QComboBox()

        sheet_setting_inner_layout = QHBoxLayout()
        sheet_setting_inner_layout.addWidget(self.sheet_combobox)
        sheet_setting_groupbox.setLayout(sheet_setting_inner_layout)

        # 작업 날짜 선택
        date_edit_groupbox = QGroupBox("날짜 선택")
        self.date_edit = QDateEdit(QDate.currentDate().addDays(-1))
        self.date_edit.setGeometry(100, 100, 150, 40)

        date_edit_inner_layout = QHBoxLayout()
        date_edit_inner_layout.addWidget(self.date_edit)
        date_edit_groupbox.setLayout(date_edit_inner_layout)

        # 날짜 버튼
        date_button_groupbox = QGroupBox("작업 날짜 추가")
        self.add_date_button = QPushButton("날짜 추가")
        self.remove_date_button = QPushButton("날짜 제거")

        self.add_date_button.clicked.connect(self.add_date_button_clicked)
        self.remove_date_button.clicked.connect(self.remove_date_button_clicked)

        date_button_inner_layout = QHBoxLayout()
        date_button_inner_layout.addWidget(self.add_date_button)
        date_button_inner_layout.addWidget(self.remove_date_button)
        date_button_groupbox.setLayout(date_button_inner_layout)

        # 날짜 리스트
        date_list_groupbox = QGroupBox("날짜 목록")
        self.date_listwidget = QListWidget()
        self.date_listwidget.setSelectionMode(QAbstractItemView.MultiSelection)

        date_list_inner_layout = QVBoxLayout()
        date_list_inner_layout.addWidget(self.date_listwidget)
        date_list_groupbox.setLayout(date_list_inner_layout)

        # 계정 엑셀 파일
        account_file_groupbox = QGroupBox("계정 엑셀 파일")
        self.account_file = QLineEdit()
        self.account_file.setText(self.saved_data.account_file)
        self.account_file.setDisabled(True)
        self.account_file_select_button = QPushButton("파일 선택")

        self.account_file_select_button.clicked.connect(self.account_file_select_button_clicked)

        account_file_inner_layout = QHBoxLayout()
        account_file_inner_layout.addWidget(self.account_file)
        account_file_inner_layout.addWidget(self.account_file_select_button)
        account_file_groupbox.setLayout(account_file_inner_layout)

        # 통계 엑셀 파일
        stats_file_groupbox = QGroupBox("통계 엑셀 파일")
        self.stats_file = QLineEdit()
        self.stats_file.textChanged.connect(self.stats_file_textChanged)
        self.stats_file.setText(self.saved_data.stats_file)
        self.stats_file.setDisabled(True)
        self.stats_file_select_button = QPushButton("파일 선택")

        self.stats_file_select_button.clicked.connect(self.stats_file_select_button_clicked)

        stats_file_inner_layout = QHBoxLayout()
        stats_file_inner_layout.addWidget(self.stats_file)
        stats_file_inner_layout.addWidget(self.stats_file_select_button)
        stats_file_groupbox.setLayout(stats_file_inner_layout)

        # 사전 작업용 브라우저
        chrome_browser_groupbox = QGroupBox("브라우저 사전 작업")
        chrome_browser_button = QPushButton("브라우저 열기")

        chrome_browser_button.clicked.connect(self.open_chrome_browser)

        browser_inner_layout = QVBoxLayout()
        browser_inner_layout.addWidget(chrome_browser_button)
        chrome_browser_groupbox.setLayout(browser_inner_layout)

        # 시작 중지
        start_stop_groupbox = QGroupBox("시작 중지")
        self.crawler_save_button = QPushButton("저장")
        self.crawler_start_button = QPushButton("시작")
        self.crawler_stop_button = QPushButton("중지")
        self.crawler_stop_button.setDisabled(True)

        self.crawler_save_button.clicked.connect(self.crawler_save_button_clicked)
        self.crawler_start_button.clicked.connect(self.crawler_start_button_clicked)
        self.crawler_stop_button.clicked.connect(self.crawler_stop_button_clicked)

        start_stop_inner_layout = QHBoxLayout()
        start_stop_inner_layout.addWidget(self.crawler_save_button)
        start_stop_inner_layout.addWidget(self.crawler_start_button)
        start_stop_inner_layout.addWidget(self.crawler_stop_button)
        start_stop_groupbox.setLayout(start_stop_inner_layout)

        # 로그 그룹박스
        log_groupbox = QGroupBox("로그")
        self.browser = QTextBrowser()

        log_inner_layout = QHBoxLayout()
        log_inner_layout.addWidget(self.browser)
        log_groupbox.setLayout(log_inner_layout)

        # 레이아웃 배치
        top_layout = QVBoxLayout()
        top_layout.addWidget(account_file_groupbox)
        top_layout.addWidget(stats_file_groupbox)

        mid_layout = QHBoxLayout()
        mid_layout.addWidget(sheet_setting_groupbox, 5)
        mid_layout.addWidget(date_edit_groupbox, 3)

        date_button_layout = QHBoxLayout()
        date_button_layout.addStretch(5)
        date_button_layout.addWidget(date_button_groupbox, 3)

        date_list_layout = QHBoxLayout()
        date_list_layout.addStretch(5)
        date_list_layout.addWidget(date_list_groupbox, 3)

        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch(5)
        bottom_layout.addWidget(chrome_browser_groupbox, 2)
        bottom_layout.addWidget(start_stop_groupbox, 3)

        log_layout = QVBoxLayout()
        log_layout.addWidget(log_groupbox)

        layout = QVBoxLayout()
        layout.addLayout(top_layout, 1)
        layout.addLayout(mid_layout, 1)
        layout.addLayout(date_button_layout, 1)
        layout.addLayout(date_list_layout, 4)
        layout.addLayout(bottom_layout, 1)
        layout.addLayout(log_layout, 8)

        self.setLayout(layout)
