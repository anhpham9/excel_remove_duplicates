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

from PyQt5.QtGui import QColor, QTextCharFormat, QTextCursor, QIcon, QFont
from PyQt5.QtCore import Qt


class ExcelCleaner(QWidget):
    def __init__(self):
        super().__init__()

        self.input_file = ""

        self.setWindowTitle("Excel重複処理ツール")
        self.setGeometry(300, 200, 700, 500)

        # アイコン設定
        self.setWindowIcon(QIcon("assets/icon.ico"))

        # スタイルシート
        self.setStyleSheet("""
            QWidget {
                font-size: 14px;
                font-family: "Segoe UI", "Meiryo", sans-serif;
                background-color: #f5f6fa;
                color: #222;
            }

            QLabel {
                font-size: 14px;
                font-weight: 600;
            }

            QLineEdit, QComboBox {
                background: white;
                border: 1px solid #cfd3dc;
                border-radius: 8px;
                padding: 8px;
                font-size: 14px;
            }

            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #4a90e2;
            }

            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #357ABD;
            }

            QPushButton:pressed {
                background-color: #2c5f94;
            }

            QTextEdit {
                background: white;
                border: 1px solid #cfd3dc;
                border-radius: 10px;
                padding: 8px;
                font-size: 13px;
            }
        """)

        self.init_ui()

        # ボタンの高さを統一
        self.run_btn.setMinimumHeight(45)
        self.file_btn.setMinimumHeight(40)


    def init_ui(self):
        layout = QVBoxLayout()

        # ファイルを選択
        file_layout = QHBoxLayout()

        self.file_label = QLabel("ファイルが選択されていません")
        self.file_btn = QPushButton("Excelファイルを選択")
        self.file_btn.clicked.connect(self.select_file)

        file_layout.addWidget(self.file_btn)
        file_layout.addWidget(self.file_label)

        layout.addLayout(file_layout)

        # カラム名
        layout.addWidget(QLabel("処理するカラム名"))

        self.column_input = QLineEdit()
        self.column_input.setPlaceholderText("例: 会社名")
        layout.addWidget(self.column_input)

        # 処理タイプ
        layout.addWidget(QLabel("処理タイプ"))

        self.process_type = QComboBox()
        self.process_type.addItems([
            "会社名",
            "URL",
        ])

        layout.addWidget(self.process_type)

        # 出力
        layout.addWidget(QLabel("出力ファイル名"))

        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("空欄の場合は自動生成されます")
        layout.addWidget(self.output_input)

        # 実行ボタン
        self.run_btn = QPushButton("処理を開始")
        self.run_btn.clicked.connect(self.process_excel)

        layout.addWidget(self.run_btn)

        # ログ
        layout.addWidget(QLabel("ログ"))

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
            "Excelファイルを選択",
            "",
            "Excel Files (*.xlsx *.xls)"
        )

        if file_name:
            self.input_file = file_name
            self.file_label.setText(os.path.basename(file_name))
            self.log(f"ファイルを選択しました: {file_name}", "green")

    def normalize_company(self, text):
        if pd.isna(text):
            return ""

        text = str(text)

        # 全角 -> 半角
        text = unicodedata.normalize("NFKC", text)

        # /以降を削除
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
                self.log("Excelファイルが選択されていません", "red")
                return

            column_name = self.column_input.text().strip()

            if not column_name:
                self.log("カラム名が入力されていません", "red")
                return

            process_type = self.process_type.currentText()

            self.log("Excelファイルを読み込んでいます...", "black")

            df = pd.read_excel(self.input_file)

            if column_name not in df.columns:
                self.log(f"カラムが見つかりません: {column_name}", "red")
                return

            self.log(f"初期行数: {len(df)}", "black")

            # 正規化処理
            if process_type == "会社名":
                df["_temp_key"] = df[column_name].apply(
                    self.normalize_company
                )

            elif process_type == "URL":
                df["_temp_key"] = df[column_name].apply(
                    self.normalize_url
                )

            before = len(df)

            # 重複行を削除
            df = df.drop_duplicates(
                subset=["_temp_key"],
                keep="first"
            )

            after = len(df)

            removed = before - after

            # 一時カラムを削除
            df = df.drop(columns=["_temp_key"])

            # 出力ファイル名
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
                f"重複行を {removed} 件削除しました",
                "orange"
            )

            self.log(
                f"ファイルの出力に成功しました: {output_path}",
                "green"
            )

            QMessageBox.information(
                self,
                "成功",
                f"ファイルを出力しました:\n{output_path}"
            )

        except Exception as e:
            self.log(f"エラー: {str(e)}", "red")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("assets/icon.ico"))


    app.setFont(QFont("Meiryo", 14))

    window = ExcelCleaner()

    window.setWindowIcon(QIcon("assets/icon.ico"))

    window.show()

    sys.exit(app.exec_())