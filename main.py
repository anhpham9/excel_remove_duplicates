import sys
import os
import re
import unicodedata
from urllib.parse import urlparse

import pandas as pd

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QTextEdit,
    QComboBox,
    QMessageBox,
)

from PyQt5.QtGui import QColor, QTextCharFormat, QTextCursor, QIcon
from PyQt5.QtCore import Qt


class ExcelCleaner(QWidget):
    def __init__(self):
        super().__init__()

        self.input_file = ""

        self.setWindowTitle("xử lý trùng lặp excel")
        self.setGeometry(300, 200, 700, 500)

        # icon app
        self.setWindowIcon(QIcon("icon.ico"))

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # chọn file
        file_layout = QHBoxLayout()

        self.file_label = QLabel("Chưa chọn file")
        self.file_btn = QPushButton("Chọn file Excel")
        self.file_btn.clicked.connect(self.select_file)

        file_layout.addWidget(self.file_btn)
        file_layout.addWidget(self.file_label)

        layout.addLayout(file_layout)

        # tên cột
        layout.addWidget(QLabel("Tên cột cần xử lý"))

        self.column_input = QLineEdit()
        self.column_input.setPlaceholderText("Ví dụ: 会社名")
        layout.addWidget(self.column_input)

        # loại xử lý
        layout.addWidget(QLabel("Loại xử lý"))

        self.process_type = QComboBox()
        self.process_type.addItems([
            "会社名",
            "URL",
        ])

        layout.addWidget(self.process_type)

        # output
        layout.addWidget(QLabel("Tên file output"))

        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("Để trống sẽ tự động tạo")
        layout.addWidget(self.output_input)

        # nút chạy
        self.run_btn = QPushButton("Bắt đầu xử lý")
        self.run_btn.clicked.connect(self.process_excel)

        layout.addWidget(self.run_btn)

        # log
        layout.addWidget(QLabel("Log"))

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        layout.addWidget(self.log_box)

        self.setLayout(layout)

    def log(self, text, color="black"):
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))

        cursor = self.log_box.textCursor()
        cursor.movePosition(QTextCursor.End)

        cursor.insertText(text + "\n", fmt)

        self.log_box.setTextCursor(cursor)
        self.log_box.ensureCursorVisible()

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel",
            "",
            "Excel Files (*.xlsx *.xls)"
        )

        if file_name:
            self.input_file = file_name
            self.file_label.setText(os.path.basename(file_name))
            self.log(f"Đã chọn file: {file_name}", "green")

    def normalize_company(self, text):
        if pd.isna(text):
            return ""

        text = str(text)

        # full width -> half width
        text = unicodedata.normalize("NFKC", text)

        # bỏ sau dấu /
        text = text.split("/")[0]

        text = text.strip().lower()

        return text

    def normalize_url(self, url):
        if pd.isna(url):
            return ""

        url = str(url).strip().lower()

        parsed = urlparse(url)

        domain = parsed.netloc

        domain = domain.replace("www.", "")

        return domain

    def process_excel(self):
        try:
            if not self.input_file:
                self.log("Chưa chọn file Excel", "red")
                return

            column_name = self.column_input.text().strip()

            if not column_name:
                self.log("Chưa nhập tên cột", "red")
                return

            process_type = self.process_type.currentText()

            self.log("Đang đọc file Excel...", "black")

            df = pd.read_excel(self.input_file)

            if column_name not in df.columns:
                self.log(f"Không tìm thấy cột: {column_name}", "red")
                return

            self.log(f"Tổng số dòng ban đầu: {len(df)}", "black")

            # xử lý normalize
            if process_type == "会社名":
                df["_temp_key"] = df[column_name].apply(
                    self.normalize_company
                )

            elif process_type == "URL":
                df["_temp_key"] = df[column_name].apply(
                    self.normalize_url
                )

            before = len(df)

            # xóa duplicate
            df = df.drop_duplicates(
                subset=["_temp_key"],
                keep="first"
            )

            after = len(df)

            removed = before - after

            # xóa cột temp
            df = df.drop(columns=["_temp_key"])

            # output name
            output_name = self.output_input.text().strip()

            if not output_name:
                base = os.path.splitext(
                    os.path.basename(self.input_file)
                )[0]

                output_name = base + "_done.xlsx"

            if not output_name.endswith(".xlsx"):
                output_name += ".xlsx"

            output_path = os.path.join(
                os.path.dirname(self.input_file),
                output_name
            )

            df.to_excel(output_path, index=False)

            self.log(
                f"Đã xóa {removed} dòng trùng lặp",
                "orange"
            )

            self.log(
                f"Xuất file thành công: {output_path}",
                "green"
            )

            QMessageBox.information(
                self,
                "Thành công",
                f"Đã xuất file:\n{output_path}"
            )

        except Exception as e:
            self.log(f"Lỗi: {str(e)}", "red")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = ExcelCleaner()

    window.show()

    sys.exit(app.exec_())