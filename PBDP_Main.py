# PBDP Source Code
from PBDP_UI import PBDP_mainUI
from PBDP_Excel import *

# PySide2
from PySide2.QtCore import QCoreApplication, QObject, QThread, Signal
from PySide2.QtWidgets import QApplication

# Python Packages
import sys
from typing import Union

from time import ( strftime as time_strftime )

from random import ( randint as random_randint )

class PBDP_main(QObject) :
    def __init__(self, app) :
        super().__init__()
        
        # Main UI
        global mainUI
        mainUI = PBDP_mainUI()
        
        # Key Event Function
        global QKeyThreaed, keyFunc
        QKeyThreaed = QThread()
        QKeyThreaed.start()
        keyFunc = PBDP_keyFunc()
        keyFunc.moveToThread(QKeyThreaed)
        
        mainUI.show()                   # Main UI Show
        
        self.signal()
        
        sys.exit(app.exec_())
    
    
    def exit(self) -> None:
        QKeyThreaed.quit()
        QCoreApplication.instance().quit()
        return
    
    
    def signal(self) -> None:
        QCoreApplication.instance().aboutToQuit.connect(self.exit)
        
        mainUI.bt_fileBrowse.clicked.connect(self.get_open_file)
        mainUI.bt_start.clicked.connect(keyFunc.start)
        
        keyFunc.logging_signal.connect(self.logging)
        
    
    
    def logging(self, message: str) -> None : 
        mainUI.lw_log.addItem(f"{time_strftime('%H:%M:%S')} - {message}")
        mainUI.lw_log.setCurrentRow(mainUI.lw_log.count()-1)
        mainUI.lw_log.scrollToBottom()
        return
    
    
    def get_open_file(self) :
        filename = mainUI.showDialog()
        mainUI.text_filePath.setText(filename)
        return



class PBDP_keyFunc(QObject) :
    logging_signal = Signal(str)

    def start(self) -> None :
        self.logging_signal.emit("작업 시작")
        
        filename = mainUI.text_filePath.text()
        problem_count = mainUI.sb_problemCount.value()

        if filename == "" :
            self.logging_signal.emit("파일 경로를 입력해주세요.")
            return
        
        wb = open_excel_file(filename)
        if wb == None :
            self.logging_signal.emit("파일을 열 수 없습니다.")
            return
        
        problems = read_problems_from_excel_file(wb)
        if problems == None :
            self.logging_signal.emit("문제를 읽을 수 없습니다.")
            return
        
        self.logging_signal.emit(f"문제 {len(problems)}개 중 {problem_count}개를 생성합니다.")
        seeds = self.get_random_seed(problem_count, len(problems))

        selected_problems = self.get_selected_problems(problems, seeds)
        if selected_problems == None :
            self.logging_signal.emit("문제를 선택할 수 없습니다.")
            return
        
        wb = create_excel_file()
        for problem in selected_problems:
            append_problem_to_excel_file(wb, problem)

        save_filename = f"{time_strftime('%Y%m%d_%H%M%S')}_{filename.split('/')[-1].split('.')[0]}_추출.xlsx"
        if save_excel_file(wb, save_filename) :
            self.logging_signal.emit(f"{save_filename} 저장 완료")
            self.logging_signal.emit("작업 완료")

    
    
    def get_random_seed(self, count, range) -> list:
        seeds = list()
        
        while len(seeds) < count:
            seed = random_randint(1, range)
            if seed not in seeds:
                seeds.append(seed)
        
        print(seeds)
        return seeds
    
    def get_selected_problems(self, problems, seeds) -> list:
        try:
            print(problems)
            print(seeds)
            selected_problems = list()

            for problem in problems:
                if problem['order'] in seeds:
                    selected_problems.append(problem)

            print(selected_problems)

            return selected_problems
        except Exception as e:
            print(e)
            return None


def launch() -> None : 
    app = QApplication(sys.argv)
    PBDP_main(app)


if __name__ == "__main__" : 
    launch()