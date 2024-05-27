import sys
import pyautogui
import pygetwindow as gw
from pywinauto import Desktop, Application
from PIL import Image, ImageDraw, ImageGrab
import pytesseract
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QLineEdit, QLabel, QInputDialog
from PyQt5.QtGui import QPalette, QColor, QFont, QGuiApplication
from PyQt5.QtCore import Qt, QCoreApplication
import keyboard
import time
from difflib import SequenceMatcher, get_close_matches
import cv2
import numpy as np
import threading
import sys
from ctypes import windll
import pythoncom
from winrt import _winrt
import os
from tray import start_tray
import subprocess
_winrt.uninit_apartment()
pythoncom.CoInitialize()
windll.ole32.CoInitialize(None)


highlighted_elements = {}
active_window_title = None
highlighted_elements = {}
ocr_elements = {}
active_window_title = None

def highlight_elements():
    
    global highlighted_elements, active_window_title
    highlighted_elements.clear()

    print(f"Active window: {active_window_title}")
    try:
        app = Application(backend="uia").connect(title=active_window_title, timeout=0.5)
        window = app.window(title=active_window_title)
        windows_to_highlight = [window]
    except pywinauto.timings.TimeoutError:
        print(f"Timeout occurred while connecting to window '{active_window_title}'. Skipping highlighting elements for the active window.")
        windows_to_highlight = []

    # Always highlight taskbar elements
    desktop = Desktop(backend="uia")
    taskbar = desktop.window(class_name="Shell_TrayWnd")
    windows_to_highlight.append(taskbar)

    for window in windows_to_highlight:
        try:
            for element in window.descendants():
                if element.is_visible() and element.is_enabled():
                    element_text = element.window_text().strip().lower()
                    try:
                        element_text.encode('charmap')
                    except UnicodeEncodeError:
                        print(f"Skipping element with unencodable character: {element_text}")
                        continue
                    
                    # Split the element_text and take the first part
                    element_name = element_text.split("-")[0].strip()
                    
                    if element_name not in highlighted_elements:
                        rect = element.rectangle()
                        highlighted_elements[element_name] = rect
                        print(f"Element '{element_name}' added to highlighted elements at {rect}.")
        except Exception as e:
            print(f"Error updating highlighted elements for window '{window.window_text()}': {e}")
def highlight_ocr():
    global ocr_elements, active_window_title
    ocr_elements.clear()  # Clear previous OCR elements before updating

    print(f"Performing OCR on active window: {active_window_title}")
    try:
        screenshot = ImageGrab.grab()
        grayscale_image = screenshot.convert("L")
        np_image = np.array(grayscale_image)
        enhanced_image = cv2.filter2D(np_image, -1, kernel=np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]]))
        ocr_text = pytesseract.image_to_data(enhanced_image, output_type=pytesseract.Output.DICT)
        
        for i in range(len(ocr_text["text"])):
            text = ocr_text["text"][i].strip().lower()
            if text:
                left = ocr_text["left"][i]
                top = ocr_text["top"][i]
                width = ocr_text["width"][i]
                height = ocr_text["height"][i]
                center_x = left + (width // 2)
                center_y = top + (height // 2)
                ocr_elements[text] = (center_x, center_y)
                print(f"OCR element '{text}' added at ({center_x}, {center_y}).")
    except Exception as e:
        print(f"Error performing OCR on active window: {e}")


def click_element(element_name):
    global highlighted_elements
    element_name = element_name.lower()
    
    # Find the closest matching element based on similarity
    closest_match = None
    highest_similarity = 0
    for highlighted_element in highlighted_elements:
        similarity = SequenceMatcher(None, element_name, highlighted_element).ratio()
        if similarity > highest_similarity:
            highest_similarity = similarity
            closest_match = highlighted_element
    
    if closest_match:
        if highest_similarity < 0.7:  # Adjust the threshold as needed
            suggestions = get_close_matches(element_name, highlighted_elements.keys(), n=3, cutoff=0.6)
            if suggestions:
                print(f"Did you mean any of these elements: {', '.join(suggestions)}?")
            else:
                print(f"No close matches found for '{element_name}'. Please check the element name.")
            return False
        
        element = highlighted_elements[closest_match]
        left, top, right, bottom = element.left, element.top, element.right, element.bottom
        center_x = left + (right - left) // 2
        center_y = top + (bottom - top) // 2
        try:
            original_x, original_y = pyautogui.position()  # Store the original cursor position
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click(button='left', _pause=False)
            pyautogui.moveTo(original_x, original_y)  # Move the cursor back to its original position
            print(f"Clicked on element '{closest_match}' at ({center_x}, {center_y}).")
            return True
        except Exception as e:
            print(f"Error clicking on element '{closest_match}' at ({center_x}, {center_y}): {e}")
            return False
    else:
        print(f"No matching element found for '{element_name}'.")
        return False


def click_element_backup(element_name):
    global ocr_elements
    element_name = element_name.lower()
    
    # Find the closest matching element based on similarity
    closest_match = None
    highest_similarity = 0
    for ocr_element in ocr_elements:
        similarity = SequenceMatcher(None, element_name, ocr_element).ratio()
        if similarity > highest_similarity:
            highest_similarity = similarity
            closest_match = ocr_element
    
    if closest_match:
        if highest_similarity < 0.7:  # Adjust the threshold as needed
            suggestions = get_close_matches(element_name, ocr_elements.keys(), n=3, cutoff=0.6)
            if suggestions:
                print(f"Did you mean any of these elements: {', '.join(suggestions)}?")
            else:
                print(f"No close matches found for '{element_name}' using OCR.")
            return False
        
        center_x, center_y = ocr_elements[closest_match]
        try:
            original_x, original_y = pyautogui.position()
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click(button='left', _pause=False)
            pyautogui.moveTo(original_x, original_y)
            print(f"Clicked on element '{closest_match}' at ({center_x}, {center_y}) using OCR.")
            return True
        except Exception as e:
            print(f"Error clicking on element '{closest_match}' at ({center_x}, {center_y}): {e}")
            return False
    else:
        print(f"No matching element found for '{element_name}' using OCR.")
        return False

def focus_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if window.isMinimized:
            window.restore()
        window.activate()
        print(f"Window '{window_title}' is now active.")
    except IndexError:
        print(f"No window found with title '{window_title}'")

class SearchBar(QDialog):
    def __init__(self, parent, search_window_title):
        super(SearchBar, self).__init__(parent, Qt.FramelessWindowHint)
        self.search_window_title = search_window_title
        screen_width = QApplication.desktop().screenGeometry().width()
        screen_height = QApplication.desktop().screenGeometry().height()
        width = int(screen_width * 0.12)
        height = 120
        x = (screen_width - width) // 2
        y = int(screen_height * 0.60)
        self.setGeometry(x, y, width, height)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.init_ui()
        self.element_list = []

    def init_ui(self):
        layout = QVBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setFont(QFont("Google Sans", 15))
        self.search_edit.setStyleSheet("""
            QLineEdit {
                border: 2px solid #4B4D50;
                border-radius: 22px;
                padding: 12px;
                background-color: #202124;
                color: #E8EAED;
            }
        """)
        self.search_edit.textChanged.connect(self.update_suggestions)
        layout.addWidget(self.search_edit)
        
        self.suggestion_label = QLabel()
        self.suggestion_label.setStyleSheet("""
            QLabel {
                color: #888888;
                font-size: 15px;
                padding-left: 12px;
            }
        """)
        layout.addWidget(self.suggestion_label)
        
        self.setLayout(layout)
        self.search_edit.returnPressed.connect(self.apply)

    def update_suggestions(self):
        search_input = self.search_edit.text().strip().lower()
        if search_input:
            # Search highlighted_elements first
            suggestions = get_close_matches(search_input, highlighted_elements.keys(), n=3, cutoff=0.35)
            if not suggestions:
                # If no suggestions found in highlighted_elements, search ocr_elements
                suggestions = get_close_matches(search_input, ocr_elements.keys(), n=3, cutoff=0.4)
            
            if suggestions:
                self.suggestion_label.setText(suggestions[0])
            else:
                self.suggestion_label.clear()
        else:
            self.suggestion_label.clear()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Tab:
            suggestion = self.suggestion_label.text()
            if suggestion:
                self.search_edit.setText(suggestion)
                self.search_edit.setFocus()
        else:
            super().keyPressEvent(event)

    def apply(self):
        search_input = self.search_edit.text().strip().lower()
        self.accept()  # Close the dialog
        return search_input

import keyboard

def wait_for_hotkey(hotkey):
    while True:
        if keyboard.is_pressed(hotkey):
            break
        time.sleep(0.1)

import sys
import pyautogui
import pygetwindow as gw
from pywinauto import Desktop, Application
from PIL import Image, ImageDraw, ImageGrab
import pytesseract
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QLineEdit, QLabel
from PyQt5.QtGui import QPalette, QColor, QFont
from PyQt5.QtCore import Qt
import keyboard
import time
from difflib import SequenceMatcher, get_close_matches
from fuzzywuzzy import fuzz
import cv2
import numpy as np
import pywinauto

active_window_title = None
highlighted_elements = {}
ocr_elements = {}

def highlight_elements():
    global highlighted_elements, active_window_title
    highlighted_elements.clear()
    print(f"Active window: {active_window_title}")
    try:
        app = Application(backend="uia").connect(title=active_window_title, timeout=0.5)
        window = app.window(title=active_window_title)
        windows_to_highlight = [window]
    except pywinauto.timings.TimeoutError:
        print(f"Timeout occurred while connecting to window '{active_window_title}'. Skipping highlighting elements for the active window.")
        windows_to_highlight = []
    desktop = Desktop(backend="uia")
    taskbar = desktop.window(class_name="Shell_TrayWnd")
    windows_to_highlight.append(taskbar)
    for window in windows_to_highlight:
        try:
            for element in window.descendants():
                if element.is_visible() and element.is_enabled():
                    element_text = element.window_text().strip().lower()
                    try:
                        element_text.encode('charmap')
                    except UnicodeEncodeError:
                        print(f"Skipping element with unencodable character: {element_text}")
                        continue
                    element_name = element_text.split("-")[0].strip()
                    
                    if element_name not in highlighted_elements:
                        rect = element.rectangle()
                        highlighted_elements[element_name] = rect
                        print(f"Element '{element_name}' added to highlighted elements at {rect}.")
        except Exception as e:
            print(f"Error updating highlighted elements for window '{window.window_text()}': {e}")
            
def highlight_ocr():
    global ocr_elements, active_window_title
    ocr_elements.clear()
    print(f"Performing OCR on active window: {active_window_title}")
    try:
        screenshot = ImageGrab.grab()
        grayscale_image = screenshot.convert("L")
        np_image = np.array(grayscale_image)
        enhanced_image = cv2.filter2D(np_image, -1, kernel=np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]]))
        ocr_text = pytesseract.image_to_data(enhanced_image, output_type=pytesseract.Output.DICT)
        
        for i in range(len(ocr_text["text"])):
            text = ocr_text["text"][i].strip().lower()
            if text:
                left = ocr_text["left"][i]
                top = ocr_text["top"][i]
                width = ocr_text["width"][i]
                height = ocr_text["height"][i]
                center_x = left + (width // 2)
                center_y = top + (height // 2)
                ocr_elements[text] = (center_x, center_y)
                print(f"OCR element '{text}' added at ({center_x}, {center_y}).")
    except Exception as e:
        print(f"Error performing OCR on active window: {e}")

def click_element(element_name):
    global highlighted_elements
    element_name = element_name.lower()
    closest_match = None
    highest_similarity = 0
    for highlighted_element in highlighted_elements:
        similarity = SequenceMatcher(None, element_name, highlighted_element).ratio()
        if similarity > highest_similarity:
            highest_similarity = similarity
            closest_match = highlighted_element
    if closest_match:
        if highest_similarity < 0.7:
            suggestions = get_close_matches(element_name, highlighted_elements.keys(), n=3, cutoff=0.6)
            if suggestions:
                print(f"Did you mean any of these elements: {', '.join(suggestions)}?")
            else:
                print(f"No close matches found for '{element_name}'. Please check the element name.")
            return False
        element = highlighted_elements[closest_match]
        left, top, right, bottom = element.left, element.top, element.right, element.bottom
        center_x = left + (right - left) // 2
        center_y = top + (bottom - top) // 2
        try:
            original_x, original_y = pyautogui.position()
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click(button='left', _pause=False)
            pyautogui.moveTo(original_x, original_y)
            print(f"Clicked on element '{closest_match}' at ({center_x}, {center_y}).")
            return True
        except Exception as e:
            print(f"Error clicking on element '{closest_match}' at ({center_x}, {center_y}): {e}")
            return False
    else:
        print(f"No matching element found for '{element_name}'.")
        return False

def click_element_backup(element_name):
    global ocr_elements
    element_name = element_name.lower()
    closest_match = None
    highest_similarity = 0
    for ocr_element in ocr_elements:
        similarity = SequenceMatcher(None, element_name, ocr_element).ratio()
        if similarity > highest_similarity:
            highest_similarity = similarity
            closest_match = ocr_element
    if closest_match:
        if highest_similarity < 0.7:
            suggestions = get_close_matches(element_name, ocr_elements.keys(), n=3, cutoff=0.6)
            if suggestions:
                print(f"Did you mean any of these elements: {', '.join(suggestions)}?")
            else:
                print(f"No close matches found for '{element_name}' using OCR.")
            return False
        center_x, center_y = ocr_elements[closest_match]
        try:
            original_x, original_y = pyautogui.position()
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click(button='left', _pause=False)
            pyautogui.moveTo(original_x, original_y)
            print(f"Clicked on element '{closest_match}' at ({center_x}, {center_y}) using OCR.")
            return True
        except Exception as e:
            print(f"Error clicking on element '{closest_match}' at ({center_x}, {center_y}): {e}")
            return False
    else:
        print(f"No matching element found for '{element_name}' using OCR.")
        return False

def focus_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if window.isMinimized:
            window.restore()
        window.activate()
        print(f"Window '{window_title}' is now active.")
    except IndexError:
        print(f"No window found with title '{window_title}'")

class SearchBar(QDialog):
    def __init__(self, parent, search_window_title):
        super(SearchBar, self).__init__(parent, Qt.FramelessWindowHint)
        self.search_window_title = search_window_title
        screen_width = QApplication.desktop().screenGeometry().width()
        screen_height = QApplication.desktop().screenGeometry().height()
        width = int(screen_width * 0.12)
        height = 100
        x = (screen_width - width) // 2
        y = int(screen_height * 0.60)
        self.setGeometry(x, y, width, height)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.init_ui()
        self.element_list = []

    def init_ui(self):
        layout = QVBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setFont(QFont("Google Sans", 15))
        self.search_edit.setStyleSheet("""
            QLineEdit {
                border: 2px solid #4B4D50;
                border-radius: 22px;
                padding: 12px;
                background-color: #202124;
                color: #E8EAED;
                min-width: 200px;
                max-width: 200px;
            }
        """)
        self.search_edit.textChanged.connect(self.update_suggestions)
        layout.addWidget(self.search_edit, alignment=Qt.AlignCenter)
        self.suggestion_labels = []
        for i in range(3):
            suggestion_label = QLabel()
            suggestion_label.setStyleSheet("""
                QLabel {
                    color: #888888;
                    font-size: 18px;
                    padding-left: 12px;
                    min-width: 200px;
                    max-width: 200px;
                }
                QLabel[selected=true] {
                    color: #FFFFFF;
                    background-color: #4B4D50;
                }
            """)
            suggestion_label.setProperty("selected", False)
            self.suggestion_labels.append(suggestion_label)
            layout.addWidget(suggestion_label, alignment=Qt.AlignCenter)
        self.selected_suggestion = 0
        self.setLayout(layout)
        self.search_edit.returnPressed.connect(self.apply)

    def update_suggestions(self):
        search_input = self.search_edit.text().strip().lower()
        print(f"Search input: {search_input}")
        if search_input:
            highlighted_suggestions = []
            for element_name in highlighted_elements.keys():
                if search_input in element_name:
                    highlighted_suggestions.append(element_name)
            ocr_suggestions = []
            for element_name in ocr_elements.keys():
                if search_input in element_name:
                    ocr_suggestions.append(element_name)
            suggestions = highlighted_suggestions + ocr_suggestions
            if len(suggestions) < 3:
                remaining_slots = 3 - len(suggestions)
                close_match_suggestions = get_close_matches(search_input, highlighted_elements.keys(), n=remaining_slots, cutoff=0.45)
                close_match_suggestions.extend(get_close_matches(search_input, ocr_elements.keys(), n=remaining_slots, cutoff=0.45))
                suggestions.extend(close_match_suggestions)
            suggestions = list(dict.fromkeys(suggestions))
            suggestions = suggestions[:3]
            print(f"Final suggestions: {suggestions}")
            for i in range(3):
                if i < len(suggestions):
                    self.suggestion_labels[i].setText(suggestions[i])
                    print(f"Setting suggestion label {i} text: {suggestions[i]}")
                    self.suggestion_labels[i].show()
                    self.suggestion_labels[i].setProperty("selected", i == self.selected_suggestion)
                    self.suggestion_labels[i].style().unpolish(self.suggestion_labels[i])
                    self.suggestion_labels[i].style().polish(self.suggestion_labels[i])
                else:
                    self.suggestion_labels[i].clear()
                    self.suggestion_labels[i].hide()
        else:
            for label in self.suggestion_labels:
                label.clear()
                label.hide()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Tab:
            self.selected_suggestion = (self.selected_suggestion + 1) % 3
            self.update_suggestion_selection()
        elif event.key() == Qt.Key_Up:
            self.selected_suggestion = (self.selected_suggestion - 1) % 3
            self.update_suggestion_selection()
        elif event.key() == Qt.Key_Down:
            self.selected_suggestion = (self.selected_suggestion + 1) % 3
            self.update_suggestion_selection()
        else:
            super().keyPressEvent(event)

    def update_suggestion_selection(self):
        for i in range(3):
            self.suggestion_labels[i].setProperty("selected", i == self.selected_suggestion)
            self.suggestion_labels[i].style().unpolish(self.suggestion_labels[i])
            self.suggestion_labels[i].style().polish(self.suggestion_labels[i])

    def apply(self):
        search_input = self.search_edit.text().strip().lower()
        suggestion = self.suggestion_labels[self.selected_suggestion].text().strip().lower()
        if suggestion:
            search_input = suggestion
        self.accept()
        return search_input

def wait_for_hotkey(hotkey):
    while True:
        if keyboard.is_pressed(hotkey):
            break
        time.sleep(0.1)
tray_started = False
def main():
    global tray_started
    if not tray_started:
        start_tray()
    tray_started = True
    global app
    global active_window_title

    QGuiApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app = QApplication(sys.argv)
    search_bar = SearchBar(None, active_window_title)
    search_bar.show()
    search_bar.accept()

    while True:
        wait_for_hotkey('ctrl+shift+space')
        ocr_elements.clear()
        highlighted_elements.clear() 
        if search_bar.isVisible():
            search_bar.reject()
        active_window = gw.getActiveWindow()
        if active_window:
            active_window_title = active_window.title
        else:
            active_window_title = "Default Window Title"
        search_bar.search_edit.clear()
        search_bar.show()
        search_bar.activateWindow()
        search_bar.raise_()
        highlight_thread = threading.Thread(target=highlight_elements)
        ocr_thread = threading.Thread(target=highlight_ocr)
        highlight_thread.start()
        ocr_thread.start()
        if search_bar.exec_() == QDialog.Accepted:
            search_input = search_bar.apply()
            if search_input:
                if search_input in highlighted_elements:
                    if not click_element(search_input):
                        print(f"Failed to click on '{search_input}' using highlighted elements.")
                elif search_input in ocr_elements:
                    if not click_element_backup(search_input):
                        print(f"Failed to click on '{search_input}' using OCR.")
                else:
                    print(f"Element '{search_input}' not found in either highlighted elements or OCR.")
        highlight_thread.join()
        ocr_thread.join()
        search_bar.hide()

if __name__ == "__main__":
    main()
pythoncom.CoUninitialize()