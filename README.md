# QuantaKey Project README

## Overview
QuantaKey is a Python application designed to automate various desktop tasks using image recognition, window management, and keyboard interactions. It leverages libraries such as pyautogui, pytesseract, and pywinauto to interact with any desktop and application environment.

## Features
- **Automated GUI interactions**: Simulate mouse interactions using only keystrokes

- **Window management**: Identify and manipulate application windows, elements, and more using pywinauto and direct Windows API calls.

- **OCR capabilities**: Extract text from the screen using pytesseract as fallback when desired search not found via pywinauto


## Installation

### Prerequisites
- Python 3.x
- PyQt5
- pytesseract
- pyautogui
- pywinauto
- OpenCV

### Setup
1. Clone the repository:
   ```
   git clone https://github.com/your-repository/QuantaKey.git
   ```
2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```

## Usage
To use the application, follow these steps:
1. Activation: Use the keyboard shortcut Ctrl+Shift+Space to activate
2. Navigation: Once activated, navigate through the search suggestions using the Tab key or the arrow keys.
3. Tray Menu: The tray menu provides options to restart or exit the application.
To run the application, execute the command below in your terminal:
   ```
   python main.py
   ```

## Compilation
To compile the project into a standalone executable, install and use PyInstaller with the following commands:
   ```
   pip install pyinstaller
   ```
   ```
   pyinstaller --onefile --windowed --exclude-module PySide6 --name QuantaKey --icon tray.ico --add-data "tray.ico;." main.py tray.py
   ```

## License
QuantaKey is distributed under a custom license. For more details, see the [LICENSE](./LICENSE) file.

