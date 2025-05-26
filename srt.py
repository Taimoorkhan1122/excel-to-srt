import sys
import os
import pandas as pd
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

def ensure_results_dir():
    results_dir = os.path.join(os.path.dirname(__file__), 'results')
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)
    return results_dir

class ConverterThread(QThread):
    update_status = pyqtSignal(str)
    show_message = pyqtSignal(str, str)

    def __init__(self, file_path, output_dir=None, output_filename=None):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir
        self.output_filename = output_filename

    def run(self):
        try:
            self.update_status.emit("Converting...")
            results_dir = self.output_dir or ensure_results_dir()

            if self.file_path.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(self.file_path)
            else:
                df = pd.read_csv(self.file_path)

            self.convert_df_to_srt(df, results_dir, self.output_filename)
            self.update_status.emit("Conversion completed")
            self.show_message.emit("Success", f"File converted successfully! Check the '{results_dir}' folder.")

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

    def convert_df_to_srt(self, df, output_dir=None, output_filename=None):
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

        output_dir = output_dir or ensure_results_dir()
        if output_filename:
            if not output_filename.lower().endswith('.srt'):
                output_filename += '.srt'
            srt_file = os.path.join(output_dir, output_filename)
        else:
            srt_file = os.path.join(output_dir, 'output.srt')
        with open(srt_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(srt_lines))


class SimpleConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Simple Excel to SRT Converter")
        self.setGeometry(100, 100, 500, 200)
        self.output_dir = None

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Select Excel/CSV File:"))

        file_layout = QHBoxLayout()
        self.file_path_input = QLineEdit()
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_input)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # Output filename input
        filename_layout = QHBoxLayout()
        filename_label = QLabel("Output Filename:")
        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("example.srt (optional)")
        filename_layout.addWidget(filename_label)
        filename_layout.addWidget(self.filename_input)
        layout.addLayout(filename_layout)

        # Add Save to... button
        save_layout = QHBoxLayout()
        self.save_to_label = QLabel("Output: results/")
        save_to_button = QPushButton("Save to...")
        save_to_button.clicked.connect(self.select_output_dir)
        save_layout.addWidget(self.save_to_label)
        save_layout.addWidget(save_to_button)
        layout.addLayout(save_layout)

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

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Folder", "")
        if dir_path:
            self.output_dir = dir_path
            self.save_to_label.setText(f"Output: {dir_path}")
        else:
            self.output_dir = None
            self.save_to_label.setText("Output: results/")

    def start_conversion(self):
        file_path = self.file_path_input.text().strip()
        output_filename = self.filename_input.text().strip()
        if not file_path or not os.path.exists(file_path):
            QMessageBox.critical(self, "Error", "Please select a valid file.")
            return

        self.status_label.setText("Starting conversion...")
        self.thread = ConverterThread(file_path, self.output_dir, output_filename)
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