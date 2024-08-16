import sys
import os
import fnmatch
import platform
import subprocess
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QFileDialog, QProgressBar, QMessageBox, QLabel, QTextEdit, QHeaderView
from PySide6.QtCore import QDateTime, QThread, Signal, Qt
from PySide6.QtGui import QFont, QImage, QPixmap

class FileSearchThread(QThread):
    progress = Signal(int)
    result = Signal(list)

    def __init__(self, directories, search_term):
        super().__init__()
        self.directories = directories
        self.search_term = search_term
        self._is_running = True

    def run(self):
        results = []
        total_files = 0
        processed_files = 0

        # 전체 파일 수 계산
        for directory in self.directories:
            for _, _, files in os.walk(directory):
                total_files += len(files)

        if total_files == 0:
            self.result.emit(results)
            return

        # 검색 수행
        for directory in self.directories:
            for root, dirs, files in os.walk(directory):
                if not self._is_running:
                    break  # _is_running 플래그를 확인하여 중지 여부를 체크
                for file in files:
                    if not self._is_running:
                        break  # 파일 검색 중에도 _is_running을 체크하여 중지 가능
                    if fnmatch.fnmatch(file.lower(), self.search_term.lower()):
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        file_mtime = os.path.getmtime(file_path)
                        results.append({
                            "name": file,
                            "path": file_path,
                            "size": f"{file_size} B",
                            "modified": int(file_mtime)
                        })
                    processed_files += 1
                    self.progress.emit(int(processed_files / total_files * 100))

        self.result.emit(results)

    def stop(self):
        self._is_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PC 파일 검색 프로그램")
        self.setGeometry(100, 100, 800, 600)

        # 중앙 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 레이아웃 설정
        main_layout = QVBoxLayout()

        # 검색조건 입력 타이틀 추가
        search_title = QLabel("검색조건입력", self)
        font = QFont()
        font.setPointSize(14)
        font.setBold(True)
        search_title.setFont(font)
        search_title.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(search_title)

        # 검색 경로 입력 필드
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit(self)
        self.path_input.setPlaceholderText("검색 경로를 입력하거나 탐색 버튼을 클릭하세요.")
        self.path_input.setFixedHeight(30)  # 폰트 크기 조정에 맞춰 입력 필드 높이 설정
        browse_button = QPushButton("탐색")
        browse_button.clicked.connect(self.browse_folder)
        browse_button.setStyleSheet("background-color: #d3e4f6; color: black;")
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(browse_button)
        main_layout.addLayout(path_layout)

        # 검색어 입력 필드
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("검색할 파일명을 입력하세요 (예: *A.*).")
        self.search_input.setFixedHeight(30)  # 폰트 크기 조정에 맞춰 입력 필드 높이 설정
        self.search_input.returnPressed.connect(self.start_search)
        main_layout.addWidget(self.search_input)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        self.search_button = QPushButton("검색")
        self.search_button.clicked.connect(self.start_search)
        self.search_button.setStyleSheet("background-color: #ecd3f6; color: black;")
        self.stop_button = QPushButton("중지")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_search)
        self.stop_button.setStyleSheet("background-color: #e4f6d3; color: black;")
        self.reset_button = QPushButton("초기화")
        self.reset_button.clicked.connect(self.reset_fields)
        self.reset_button.setStyleSheet("background-color: #d3f6d3; color: black;")
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.reset_button)
        main_layout.addLayout(button_layout)

        # 검색결과현황 타이틀 추가
        result_title = QLabel("검색결과현황", self)
        result_font = QFont()
        result_font.setPointSize(14)
        result_font.setBold(True)
        result_title.setFont(result_font)
        result_title.setAlignment(Qt.AlignCenter)
        main_layout.addSpacing(10)  # 버튼 섹션과 검색 리스트 섹션 사이에 간격 추가
        main_layout.addWidget(result_title)

        # 프로그레스 바
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # 검색 결과 테이블
        self.result_table = QTableWidget(self)
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["파일명", "경로", "크기", "수정 날짜"])
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setStyleSheet("background-color: #D3D3D3; color: black;")
        self.result_table.cellClicked.connect(self.preview_item)  # 클릭 시 미리보기 기능
        self.result_table.cellDoubleClicked.connect(self.open_item)  # 더블 클릭 시 파일 실행
        main_layout.addWidget(self.result_table)

        # 검색 중 메시지 레이블
        self.searching_label = QLabel("검색 중입니다. 잠시만 기다려 주세요.", self)
        self.searching_label.setAlignment(Qt.AlignCenter)
        self.searching_label.setStyleSheet("font-size: 12pt; color: gray;")
        self.searching_label.setVisible(False)
        main_layout.addWidget(self.searching_label)

        # 파일 미리보기 영역
        self.preview_area = QTextEdit(self)
        self.preview_area.setReadOnly(True)
        self.preview_area.setFixedHeight(150)  # 미리보기 영역 높이 설정
        main_layout.addWidget(self.preview_area)

        # 인용구 박스
        quote_label = QLabel("made by 나종춘(2024)", self)
        quote_label.setStyleSheet("""
            font-size: 9pt;
            font-style: nomal;
        """)
        quote_label.setAlignment(Qt.AlignRight)  # 텍스트를 오른쪽 정렬
        quote_label.setFixedSize(800, 20)
        main_layout.addWidget(quote_label)

        # 메인 레이아웃 설정
        central_widget.setLayout(main_layout)

        # 검색 스레드 초기화
        self.search_thread = None

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", "")
        if folder:
            self.path_input.setText(folder)

    def start_search(self):
        path = self.path_input.text()
        search_term = self.search_input.text()

        if not search_term:
            return  # 검색어가 없으면 실행하지 않음

        # 와일드카드가 없는 경우 자동으로 * 추가
        if not any(char in search_term for char in ["*", "?"]):
            search_term = f"*{search_term}*"

        if not path:
            if platform.system() == "Windows":
                drives = [f"{drive}:\\" for drive in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{drive}:\\")]
                directories = drives
            else:
                directories = ["/"]
        else:
            directories = [path]

        self.result_table.setRowCount(0)
        self.searching_label.setVisible(True)  # 검색 중 메시지 표시
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.search_thread = FileSearchThread(directories, search_term)
        self.search_thread.progress.connect(self.update_progress)
        self.search_thread.result.connect(self.show_results)
        self.search_thread.start()

        self.search_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def show_results(self, results):
        self.searching_label.setVisible(False)  # 검색 완료 시 메시지 숨기기
        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)
        self.stop_button.setEnabled(False)

        for i, result in enumerate(results):
            self.result_table.insertRow(i)
            self.result_table.setItem(i, 0, QTableWidgetItem(result["name"]))
            self.result_table.setItem(i, 1, QTableWidgetItem(result["path"]))
            self.result_table.setItem(i, 2, QTableWidgetItem(result["size"]))
            self.result_table.setItem(i, 3, QTableWidgetItem(QDateTime.fromSecsSinceEpoch(result["modified"]).toString("yyyy-MM-dd HH:mm:ss")))

        QMessageBox.information(self, "검색 완료", f"{len(results)}개의 파일을 검색했습니다.")

    def stop_search(self):
        if self.search_thread and self.search_thread.isRunning():
            self.search_thread.stop()
            self.search_thread.wait()

        self.searching_label.setVisible(False)
        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def reset_fields(self):
        self.path_input.clear()
        self.search_input.clear()
        self.result_table.setRowCount(0)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.preview_area.clear()  # 미리보기 영역 초기화

    def preview_item(self, row, column):
        file_path = self.result_table.item(row, 1).text()
        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension in ['.txt', '.log']:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read(500)
                self.preview_area.setPlainText(content)

        elif file_extension in ['.jpg', '.jpeg', '.gif', '.png', '.bmp']:
            image = QImage(file_path)
            pixmap = QPixmap.fromImage(image)
            self.preview_area.setPlainText("")
            self.preview_area.insertHtml(f'<img src="{file_path}" width="300" height="150">')

        else:
            self.preview_area.setPlainText("미리보기를 지원하지 않는 파일 유형입니다.")

    def open_item(self, row, column):
        file_path = self.result_table.item(row, 1).text()

        if column == 0:  # 파일명 클릭 시 파일 실행
            try:
                os.startfile(file_path)  # Windows에서 파일 실행
            except AttributeError:
                subprocess.call(["open", file_path])  # macOS에서 파일 실행
        elif column == 1:  # 경로 클릭 시 폴더 열기
            folder_path = os.path.dirname(file_path)
            try:
                os.startfile(folder_path)  # Windows에서 폴더 열기
            except AttributeError:
                subprocess.call(["open", folder_path])  # macOS에서 폴더 열기

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
