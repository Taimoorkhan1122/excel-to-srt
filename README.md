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

## Creating an Executable

You can create a standalone executable for this application using [PyInstaller](https://pyinstaller.org/):

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Run PyInstaller to generate the executable. You have two options:

   - **Single file (`--onefile`):**
     ```bash
     pyinstaller --onefile --windowed srt.py
     ```
     - The `--onefile` flag creates a single executable file.
     - The `--windowed` flag prevents a terminal window from opening (for GUI apps).
     - **Pros:** Easy to distribute, just one file.
     - **Cons:** Startup/loading time can be slower, as everything is unpacked to a temporary folder at runtime.

   - **Directory (`--onedir`, default):**
     ```bash
     pyinstaller --onedir --windowed srt.py
     ```
     - The `--onedir` flag (default) creates a folder containing the executable and all dependencies.
     - **Pros:** Faster startup/loading time.
     - **Cons:** Distribution requires sending the whole folder, not just a single file.

3. The executable or folder will be located in the `dist` directory.

4. Distribute the executable file (`dist/srt.exe` on Windows, `dist/srt` on macOS/Linux) or the entire folder (for `--onedir`) to users. They do not need Python installed to run it.

## License

[MIT License](LICENSE)