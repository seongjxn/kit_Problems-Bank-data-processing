from os import path
from time import strftime

# PySide2 관련 패키지
from PySide2.QtWidgets import (QMainWindow, QFileDialog, QLabel, QSpinBox,
                               QPushButton, QGroupBox, QLineEdit, QListWidget )
from PySide2.QtCore import Qt
from PySide2.QtGui import QCursor, QIcon

# UI 클래스
class PBDP_mainUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.mainUI()
    
    def mainUI(self):
        # 기본 창 설정
        self.setFixedSize(400, 580)
        self.setWindowTitle("Problems Bank Data Processing")
        
        
        self.grp = QGroupBox(self)
        self.grp.setGeometry(20, 30, 361, 81)
        
        self.lb_filePath = QLabel(self.grp)
        self.lb_filePath.setGeometry(10, 20, 70, 23)
        self.lb_filePath.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        self.lb_filePath.setText("파일 경로")
        
        self.text_filePath = QLineEdit(self.grp)
        self.text_filePath.setGeometry(90, 20, 181, 23)
        
        self.bt_fileBrowse = QPushButton(self.grp)
        self.bt_fileBrowse.setGeometry(275, 20, 75, 23)
        self.bt_fileBrowse.setText("파일 열기")
        self.bt_fileBrowse.setCursor(QCursor(Qt.PointingHandCursor))
        self.bt_fileBrowse.setDefault(True)
        
        self.lb_problemCount = QLabel(self.grp)
        self.lb_problemCount.setGeometry(10, 50, 70, 23)
        self.lb_problemCount.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        self.lb_problemCount.setText("문제 수")

        self.sb_problemCount = QSpinBox(self.grp)
        self.sb_problemCount.setGeometry(90, 50, 51, 23)
        self.sb_problemCount.setMinimum(1)
        self.sb_problemCount.setValue(20)
        
        
        # 로그
        self.lw_log = QListWidget(self)
        self.lw_log.setGeometry(20, 120, 361, 401)
        self.lw_log.addItem(f"{strftime('%H:%M:%S')} - 프로그램 실행 완료.")
        self.lw_log.setFocusPolicy(Qt.NoFocus)
        self.lw_log.setCurrentRow(0)
        
        
        # 시작 버튼
        self.bt_start = QPushButton(self)
        self.bt_start.setGeometry(306, 535, 75, 23)
        self.bt_start.setText("작업 시작")
        self.bt_start.setCursor(QCursor(Qt.PointingHandCursor))
    
    
    def showDialog(self):
        return QFileDialog.getOpenFileName(self, 'Open file', './', "Excel(*.xlsx *xls *csv)")[0]