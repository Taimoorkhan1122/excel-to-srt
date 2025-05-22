# Excel/CSV to SRT Converter

A simple GUI application that converts Excel/CSV files containing subtitle data into SRT format.

## Features

- Supports both Excel (.xls, .xlsx) and CSV files
- Simple graphical user interface
- Handles Japanese subtitle data
- Automatic time and duration calculations

## Requirements

- Python 3.6 or higher
- PyQt5
- Required Python packages (listed in requirements.txt)

## Installation

1. Clone or download this repository
2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the script:
```bash
python srt.py
```

2. In the application window:
   - Click "Browse" to select your Excel/CSV file
   - Click "Convert" to start the conversion
   - The converted SRT file will be saved in the same directory as the input file

## Input File Format

The Excel/CSV file must contain the following columns:
- ページナンバー (Page Number)
- 送出タイミング (Timing)
- Duration
- 字幕文字列 (Subtitle Text)

## Error Handling

The application will display error messages if:
- Required columns are missing
- Input file format is invalid
- Time format is incorrect

## License

[MIT License](LICENSE)