import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,  # 버튼 추가 위한 임포트
)


class StockEntryForm(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("주식 매수 입력")

        layout = QVBoxLayout()  # 전체 레이아웃

        for i in range(0, 9, 3):
            h_layout = QHBoxLayout()

            for j in range(i, i + 3):
                title = [
                    "거래일자",
                    "국가명",
                    "증권사명",
                    "종목명",
                    "틱커명",
                    "매수단가",
                    "매수수량",
                    "달러 매수금",
                    "원화 매수금",
                ][j]

                label = QLabel(title)
                line_edit = QLineEdit()

                h_layout.addWidget(label)
                h_layout.addWidget(line_edit)

            layout.addLayout(h_layout)

        # 버튼 레이아웃 추가:
        button_layout = QHBoxLayout()  # 가로 방향 배치

        addButton = QPushButton("추가")
        modifyButton = QPushButton("수정")
        deleteButton = QPushButton("삭제")
        clearButton = QPushButton("초기화")

        button_layout.addWidget(addButton)
        button_layout.addWidget(modifyButton)
        button_layout.addWidget(deleteButton)
        button_layout.addWidget(clearButton)

        layout.addLayout(button_layout)  # 버튼 레이아웃 추가

        self.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = StockEntryForm()
    form.show()
    sys.exit(app.exec_())
