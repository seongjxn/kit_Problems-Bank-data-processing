from os import path
from time import strftime

# PySide2 관련 패키지
from PySide2.QtWidgets import (QMainWindow, QDoubleSpinBox, QLabel, QCheckBox, QRadioButton,
                               QPushButton, QGroupBox, QLineEdit, QListWidget, QDateEdit, QDateTimeEdit )
from PySide2.QtCore import Qt, QCoreApplication, QEvent, QObject, QThread, Signal, QDate, QDateTime
from PySide2.QtGui import QCursor, QIcon

# UI 클래스
class MainUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.mainUI()

    def mainUI(self):
        # 기본 창 설정
        self.setFixedSize(400, 580)
        self.setWindowTitle("Problems Bank Data Processing")


        # Macro
        self.grp_macro = QGroupBox(self)
        self.grp_macro.setGeometry(20, 30, 361, 81)
        self.grp_macro.setTitle("Macro")

        self.lb_macro_id = QLabel(self.grp_macro)
        self.lb_macro_id.setGeometry(10, 20, 21, 22)
        self.lb_macro_id.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        self.lb_macro_id.setText("ID")

        self.text_macro_id = QLineEdit(self.grp_macro)
        self.text_macro_id.setGeometry(40, 20, 113, 20)

        self.lb_macro_pw = QLabel(self.grp_macro)
        self.lb_macro_pw.setGeometry(10, 50, 21, 22)
        self.lb_macro_pw.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        self.lb_macro_pw.setText("PW")

        self.text_macro_pw = QLineEdit(self.grp_macro)
        self.text_macro_pw.setGeometry(40, 50, 113, 20)
        self.text_macro_pw.setEchoMode(QLineEdit.PasswordEchoOnEdit)

        self.lb_macro_date = QLabel(self.grp_macro)
        self.lb_macro_date.setGeometry(160, 50, 71, 22)
        self.lb_macro_date.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        self.lb_macro_date.setText("예약할 날짜")

        self.date_macro = QDateEdit(self.grp_macro)
        self.date_macro.setGeometry(240, 50, 110, 22)
        self.date_macro.setDate(QDate.currentDate())


        # 로그
        self.lw_log = QListWidget(self)
        self.lw_log.setGeometry(20, 120, 361, 401)
        self.lw_log.addItem(f"{strftime('%H:%M:%S')} - 프로그램 실행")
        self.lw_log.setFocusPolicy(Qt.NoFocus)
        self.lw_log.setCurrentRow(0)


        # 예약 시간
        self.lb_start = QLabel(self)
        self.lb_start.setGeometry(22, 535, 81, 22)
        self.lb_start.setText("예약 시작 일시")

        self.date_start = QDateTimeEdit(self)
        self.date_start.setGeometry(109, 535, 121, 22)
        self.date_start.setDisplayFormat("yyyy-MM-dd hh:mm")
        self.date_start.setDateTime(QDateTime.currentDateTime())


        # 시작 버튼
        self.bt_start = QPushButton(self)
        self.bt_start.setGeometry(306, 535, 75, 23)
        self.bt_start.setText("작업 시작")
        self.bt_start.setCursor(QCursor(Qt.PointingHandCursor))
        self.bt_start.setDefault(True)
