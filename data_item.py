# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QHeaderView,
)
from PyQt5.QtCore import Qt

# 엑셀 파일 이름
excel_filename = "data_item.xlsx"

# 엑셀 파일 생성 (없을 시)
if not os.path.exists(excel_filename):
    df = pd.DataFrame(columns=["항목1", "항목2", "항목3"])  # 예시로 기본 열 이름 사용
    df.to_excel(excel_filename, index=False)

# 엑셀 파일 읽기
df = pd.read_excel(excel_filename)


class StockManagementApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Stock Management Program")

        # 레이아웃 설정
        layout = QVBoxLayout()

        # 첫 번째 섹션: 입력 박스
        self.inputs = {}
        input_layout = QHBoxLayout()

        for column in df.columns:
            label = QLabel(column)
            line_edit = QLineEdit()
            line_edit.textChanged.connect(self.convert_to_uppercase)
            input_layout.addWidget(label)
            input_layout.addWidget(line_edit)
            self.inputs[column] = line_edit

        layout.addLayout(input_layout)

        # 두 번째 섹션: 버튼
        button_layout = QHBoxLayout()

        add_button = QPushButton("추가")
        update_button = QPushButton("수정")
        delete_button = QPushButton("삭제")
        clear_button = QPushButton("초기화")

        add_button.clicked.connect(self.add_entry)
        update_button.clicked.connect(self.update_entry)
        delete_button.clicked.connect(self.delete_entry)
        clear_button.clicked.connect(self.clear_inputs)

        button_layout.addWidget(add_button)
        button_layout.addWidget(update_button)
        button_layout.addWidget(delete_button)
        button_layout.addWidget(clear_button)

        layout.addLayout(button_layout)

        # 세 번째 섹션: 리스트 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.load_data()

        # 테이블에서 셀 클릭 시 입력 양식에 데이터를 로드하는 기능 연결
        self.table.cellClicked.connect(self.load_entry_to_form)

        layout.addWidget(self.table)

        self.setLayout(layout)
        self.resize(800, 600)
        self.show()

    def convert_to_uppercase(self, text):
        """입력된 소문자를 대문자로 변환"""
        sender = self.sender()
        sender.setText(text.upper())

    def load_data(self):
        """엑셀 파일의 데이터를 테이블에 로드"""
        self.table.setRowCount(len(df))
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                self.table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

    def load_entry_to_form(self, row, column):
        """선택된 행의 데이터를 입력 양식으로 로드"""
        for j, col in enumerate(df.columns):
            value = str(df.iloc[row, j])
            self.inputs[col].setText(value)

    def add_entry(self):
        """입력 데이터를 추가"""
        global df  # 함수 맨 앞에 global 선언
        new_data = {col: self.inputs[col].text() for col in df.columns}

        # 입력된 데이터가 없는 경우 빈 값으로 채우기
        new_data = {col: (val if val else "") for col, val in new_data.items()}

        # 적어도 하나의 필드가 입력된 경우에만 추가
        if any(new_data.values()):
            new_row = pd.DataFrame([new_data])
            df = pd.concat([df, new_row], ignore_index=True)
            self.save_data()
            self.load_data()
            self.clear_inputs()
        else:
            QMessageBox.warning(self, "경고", "추가할 데이터를 입력해주세요.")

    def update_entry(self):
        """선택된 행의 데이터를 수정"""
        global df  # 함수 맨 앞에 global 선언
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            for j, col in enumerate(df.columns):
                df.at[selected_row, col] = self.inputs[col].text()
            self.save_data()
            self.load_data()
            self.clear_inputs()
        else:
            QMessageBox.warning(self, "경고", "수정할 행을 선택해주세요.")

    def delete_entry(self):
        """선택된 행의 데이터를 삭제"""
        global df  # 함수 맨 앞에 global 선언
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            df = df.drop(selected_row).reset_index(drop=True)
            self.save_data()
            self.load_data()
            self.clear_inputs()
        else:
            QMessageBox.warning(self, "경고", "삭제할 행을 선택해주세요.")

    def clear_inputs(self):
        """입력란 초기화"""
        for line_edit in self.inputs.values():
            line_edit.clear()

    def save_data(self):
        """데이터를 엑셀 파일에 저장"""
        df.to_excel(excel_filename, index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = StockManagementApp()
    sys.exit(app.exec_())
