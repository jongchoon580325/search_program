import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QDateEdit,
    QComboBox,
    QPushButton,
    QGridLayout,
    QVBoxLayout,
    QHBoxLayout,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
)
from PyQt5.QtCore import QDate, Qt


class StockManagementApp(QWidget):
    def __init__(self):
        super().__init__()

        # 최초 프로그램 실행 시 'data_save.xlsx' 파일 생성
        self.initialize_excel_file()

        # 창 제목 설정
        self.setWindowTitle("주식 매수 관리")

        # 그리드 레이아웃 생성
        grid = QGridLayout()

        # 첫 번째 줄 항목
        grid.addWidget(QLabel("거래일자"), 0, 0)
        self.trade_date = QDateEdit(self)
        self.trade_date.setCalendarPopup(True)
        self.trade_date.setDisplayFormat("yyyy-MM-dd")  # 날짜 형식 지정
        self.trade_date.setDate(QDate.currentDate())  # 현재 날짜로 설정
        grid.addWidget(self.trade_date, 0, 1)

        grid.addWidget(QLabel("국가"), 0, 2)
        self.country = QComboBox(self)
        self.country.addItems(
            ["대한민국", "미국", "일본", "중국"]
        )  # 필요에 따라 항목 추가
        grid.addWidget(self.country, 0, 3)

        grid.addWidget(QLabel("증권사"), 0, 4)
        self.broker = QComboBox(self)
        self.broker.addItems(
            ["삼성증권", "미래에셋", "키움증권"]
        )  # 필요에 따라 항목 추가
        grid.addWidget(self.broker, 0, 5)

        grid.addWidget(QLabel("계좌번호"), 0, 6)
        self.account_number = QComboBox(self)
        self.account_number.addItems(
            ["123-456-789", "987-654-321", "555-666-777"]
        )  # 예시 항목
        grid.addWidget(self.account_number, 0, 7)

        grid.addWidget(QLabel("종목명"), 0, 8)
        self.stock_name = QComboBox(self)
        self.stock_name.addItems(["삼성전자", "애플", "테슬라"])  # 예시 항목
        grid.addWidget(self.stock_name, 0, 9)

        # 두 번째 줄 항목
        grid.addWidget(QLabel("틱커명"), 1, 0)
        self.ticker = QComboBox(self)
        self.ticker.addItems(["005930", "AAPL", "TSLA"])  # 예시 항목
        grid.addWidget(self.ticker, 1, 1)

        grid.addWidget(QLabel("매수단가"), 1, 2)
        self.purchase_price = QLineEdit(self)
        grid.addWidget(self.purchase_price, 1, 3)

        grid.addWidget(QLabel("매수수량"), 1, 4)
        self.purchase_quantity = QLineEdit(self)
        grid.addWidget(self.purchase_quantity, 1, 5)

        grid.addWidget(QLabel("달러매수금"), 1, 6)
        self.purchase_amount_usd = QLineEdit(self)
        grid.addWidget(self.purchase_amount_usd, 1, 7)

        grid.addWidget(QLabel("원화매수금"), 1, 8)
        self.purchase_amount_krw = QLineEdit(self)
        grid.addWidget(self.purchase_amount_krw, 1, 9)

        # 세 번째 줄 버튼들
        btn_layout = QHBoxLayout()

        self.add_button = QPushButton("추가", self)
        self.add_button.setFixedSize(170, 25)
        self.add_button.clicked.connect(self.add_entry)  # 버튼 클릭 시 데이터 추가
        btn_layout.addWidget(self.add_button)

        self.update_button = QPushButton("수정", self)
        self.update_button.setFixedSize(170, 25)
        self.update_button.clicked.connect(
            self.update_entry
        )  # 버튼 클릭 시 데이터 수정
        btn_layout.addWidget(self.update_button)

        self.delete_button = QPushButton("삭제", self)
        self.delete_button.setFixedSize(170, 25)
        self.delete_button.clicked.connect(
            self.delete_entry
        )  # 버튼 클릭 시 데이터 삭제
        btn_layout.addWidget(self.delete_button)

        self.reset_button = QPushButton("초기화", self)
        self.reset_button.setFixedSize(170, 25)
        self.reset_button.clicked.connect(self.reset_form)  # 초기화 버튼 기능
        btn_layout.addWidget(self.reset_button)

        # 최종 레이아웃 구성
        layout = QVBoxLayout()
        layout.addLayout(grid)
        layout.addLayout(btn_layout)

        # 테이블 생성
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            [
                "거래일자",
                "국가",
                "증권사",
                "계좌번호",
                "종목명",
                "틱커명",
                "매수단가",
                "매수수량",
                "달러매수금",
                "원화매수금",
            ]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.cellClicked.connect(self.load_entry)  # 행 선택 시 데이터 불러오기
        layout.addWidget(self.table)

        self.setLayout(layout)

        # 창 크기 조정 및 화면 중앙 배치
        self.resize(1000, 400)
        self.center()

        # 초기 데이터 로드
        self.load_data()

    def center(self):
        qr = self.frameGeometry()
        cp = QApplication.desktop().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initialize_excel_file(self):
        file_path = "data_save.xlsx"
        if not os.path.exists(file_path):
            df = pd.DataFrame(
                columns=[
                    "거래일자",
                    "국가",
                    "증권사",
                    "계좌번호",
                    "종목명",
                    "틱커명",
                    "매수단가",
                    "매수수량",
                    "달러매수금",
                    "원화매수금",
                ]
            )
            df.to_excel(file_path, index=False)

    def load_data(self):
        file_path = "data_save.xlsx"
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
            self.table.setRowCount(len(df))
            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    self.table.setItem(i, j, QTableWidgetItem(str(value)))

    def add_entry(self):
        # 입력된 데이터를 읽어옴
        data = {
            "거래일자": self.trade_date.date().toString("yyyy-MM-dd"),
            "국가": self.country.currentText(),
            "증권사": self.broker.currentText(),
            "계좌번호": self.account_number.currentText(),
            "종목명": self.stock_name.currentText(),
            "틱커명": self.ticker.currentText(),
            "매수단가": self.purchase_price.text(),
            "매수수량": self.purchase_quantity.text(),
            "달러매수금": self.purchase_amount_usd.text(),
            "원화매수금": self.purchase_amount_krw.text(),
        }

        # 테이블에 데이터 추가
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for i, (key, value) in enumerate(data.items()):
            self.table.setItem(row_position, i, QTableWidgetItem(value))

        # 엑셀 파일에 저장
        self.save_to_excel()

        QMessageBox.information(self, "성공", "데이터가 저장되었습니다.")
        self.reset_form()

    def update_entry(self):
        row = self.table.currentRow()
        if row >= 0:
            # 테이블 업데이트
            self.table.setItem(
                row, 0, QTableWidgetItem(self.trade_date.date().toString("yyyy-MM-dd"))
            )
            self.table.setItem(row, 1, QTableWidgetItem(self.country.currentText()))
            self.table.setItem(row, 2, QTableWidgetItem(self.broker.currentText()))
            self.table.setItem(
                row, 3, QTableWidgetItem(self.account_number.currentText())
            )
            self.table.setItem(row, 4, QTableWidgetItem(self.stock_name.currentText()))
            self.table.setItem(row, 5, QTableWidgetItem(self.ticker.currentText()))
            self.table.setItem(row, 6, QTableWidgetItem(self.purchase_price.text()))
            self.table.setItem(row, 7, QTableWidgetItem(self.purchase_quantity.text()))
            self.table.setItem(
                row, 8, QTableWidgetItem(self.purchase_amount_usd.text())
            )
            self.table.setItem(
                row, 9, QTableWidgetItem(self.purchase_amount_krw.text())
            )

            # 엑셀 파일에 저장
            self.save_to_excel()

            QMessageBox.information(self, "성공", "데이터가 수정되었습니다.")
            self.reset_form()
        else:
            QMessageBox.warning(self, "오류", "수정할 데이터를 선택하세요.")

    def delete_entry(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

            # 엑셀 파일에 저장
            self.save_to_excel()

            QMessageBox.information(self, "성공", "데이터가 삭제되었습니다.")
            self.reset_form()
        else:
            QMessageBox.warning(self, "오류", "삭제할 데이터를 선택하세요.")

    def load_entry(self, row, column):
        self.trade_date.setDate(
            QDate.fromString(self.table.item(row, 0).text(), "yyyy-MM-dd")
        )
        self.country.setCurrentText(self.table.item(row, 1).text())
        self.broker.setCurrentText(self.table.item(row, 2).text())
        self.account_number.setCurrentText(self.table.item(row, 3).text())
        self.stock_name.setCurrentText(self.table.item(row, 4).text())
        self.ticker.setCurrentText(self.table.item(row, 5).text())
        self.purchase_price.setText(self.table.item(row, 6).text())
        self.purchase_quantity.setText(self.table.item(row, 7).text())
        self.purchase_amount_usd.setText(self.table.item(row, 8).text())
        self.purchase_amount_krw.setText(self.table.item(row, 9).text())

    def save_to_excel(self):
        file_path = "data_save.xlsx"
        data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for column in range(self.table.columnCount()):
                item = self.table.item(row, column)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        df = pd.DataFrame(
            data,
            columns=[
                "거래일자",
                "국가",
                "증권사",
                "계좌번호",
                "종목명",
                "틱커명",
                "매수단가",
                "매수수량",
                "달러매수금",
                "원화매수금",
            ],
        )
        df.to_excel(file_path, index=False)

    def reset_form(self):
        self.purchase_price.clear()
        self.purchase_quantity.clear()
        self.purchase_amount_usd.clear()
        self.purchase_amount_krw.clear()


if __name__ == "__main__":
    # 필요한 라이브러리 설치 안내
    try:
        import pandas as pd
    except ImportError:
        print(
            "pandas 라이브러리가 필요합니다. 설치하려면 'pip install pandas' 명령어를 사용하세요."
        )

    app = QApplication(sys.argv)
    window = StockManagementApp()
    window.show()
    sys.exit(app.exec_())
