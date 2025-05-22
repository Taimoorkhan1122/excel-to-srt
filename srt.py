import sys
import os
import pandas as pd
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


class ConverterThread(QThread):
    update_status = pyqtSignal(str)
    show_message = pyqtSignal(str, str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            self.update_status.emit("Converting...")
            csv_file = self.file_path

            if self.file_path.lower().endswith(('.xls', '.xlsx')):
                df = pd.read_excel(self.file_path)
                csv_file = self.file_path.rsplit('.', 1)[0] + '.csv'
                df.to_csv(csv_file, index=False, encoding='utf-8')

            self.convert_csv_to_srt(csv_file)
            self.update_status.emit("Conversion completed")
            self.show_message.emit("Success", "File converted successfully!")

        except Exception as e:
            self.update_status.emit("Error occurred")
            self.show_message.emit("Error", str(e))

    def parse_time(self, time_str):
        if pd.isna(time_str) or time_str == '-:-:-:--':
            return None
        time_str = str(time_str).strip()
        parts = time_str.split(' ')
        if len(parts) != 2:
            raise ValueError(f"Invalid time format: {time_str}")
        time_part, ms_part = parts
        hours, minutes, seconds = map(int, time_part.split(':'))
        try:
            ms = int(ms_part)
        except ValueError:
            ms = 0
        return hours, minutes, seconds, ms

    def parse_duration(self, duration_str):
        if pd.isna(duration_str):
            return timedelta(0)
        time_components = self.parse_time(duration_str)
        if time_components is None:
            return timedelta(0)
        hours, minutes, seconds, ms = time_components
        return timedelta(hours=hours, minutes=minutes, seconds=seconds, milliseconds=ms)

    def time_to_srt_format(self, time_components):
        if time_components is None:
            return "00:00:00,000"
        hours, minutes, seconds, ms = time_components
        return f"{hours:02d}:{minutes:02d}:{seconds:02d},{ms:03d}"

    def add_time_and_duration(self, start_time_str, duration_str):
        if start_time_str == '-:-:-:--' or pd.isna(start_time_str):
            return "00:00:00,000"
        time_components = self.parse_time(start_time_str)
        if time_components is None:
            return "00:00:00,000"
        hours, minutes, seconds, ms = time_components
        start_time = datetime(1, 1, 1, hour=hours, minute=minutes, second=seconds, microsecond=ms*1000)
        duration = self.parse_duration(duration_str)
        end_time = start_time + duration
        return f"{end_time.hour:02d}:{end_time.minute:02d}:{end_time.second:02d},{end_time.microsecond//1000:03d}"

    def convert_csv_to_srt(self, csv_file):
        df = pd.read_csv(csv_file)
        required_columns = ['\u30da\u30fc\u30b8\u30ca\u30f3\u30d0\u30fc', '\u9001\u51fa\u30bf\u30a4\u30df\u30f3\u30b0', 'Duration', '\u5b57\u5e55\u6587\u5b57\u5217']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")

        srt_lines = []
        for i, row in df.iterrows():
            srt_lines.append(str(i + 1))
            start_time_str = row['\u9001\u51fa\u30bf\u30a4\u30df\u30f3\u30b0']
            duration_str = row['Duration']
            start_time = self.time_to_srt_format(self.parse_time(start_time_str))
            end_time = self.add_time_and_duration(start_time_str, duration_str)
            srt_lines.append(f"{start_time} --> {end_time}")
            subtitle_text = str(row['\u5b57\u5e55\u6587\u5b57\u5217']).strip() if pd.notna(row['\u5b57\u5e55\u6587\u5b57\u5217']) else ''
            srt_lines.append(subtitle_text)
            srt_lines.append('')

        srt_file = csv_file.rsplit('.', 1)[0] + '.srt'
        with open(srt_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(srt_lines))


class SimpleConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Simple Excel to SRT Converter")
        self.setGeometry(100, 100, 500, 200)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Select Excel/CSV File:"))

        file_layout = QHBoxLayout()
        self.file_path_input = QLineEdit()
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_input)
        file_layout.addWidget(browse_button)

        layout.addLayout(file_layout)

        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        button_layout = QHBoxLayout()
        convert_button = QPushButton("Convert")
        convert_button.clicked.connect(self.start_conversion)
        exit_button = QPushButton("Exit")
        exit_button.clicked.connect(self.close)
        button_layout.addWidget(convert_button)
        button_layout.addWidget(exit_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel/CSV Files (*.xls *.xlsx *.csv)")
        if path:
            self.file_path_input.setText(path)

    def start_conversion(self):
        file_path = self.file_path_input.text().strip()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.critical(self, "Error", "Please select a valid file.")
            return

        self.status_label.setText("Starting conversion...")
        self.thread = ConverterThread(file_path)
        self.thread.update_status.connect(self.status_label.setText)
        self.thread.show_message.connect(lambda title, msg: QMessageBox.information(self, title, msg))
        self.thread.start()


def main():
    app = QApplication(sys.argv)
    window = SimpleConverterApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()