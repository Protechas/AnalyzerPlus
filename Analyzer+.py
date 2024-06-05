import sys
import os
import json
import re
import sqlite3
import logging
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QFrame, QMainWindow, QToolBar, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit, QPushButton, QComboBox, QTextBrowser, QMessageBox, QStatusBar, QInputDialog, QFileDialog, QSlider, QLabel, QListWidget, QSpacerItem, QSizePolicy, QSplitter)
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon, QDesktopServices

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class ResizableToolBar(QToolBar):
    def __init__(self, parent=None):
        super(ResizableToolBar, self).__init__(parent)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not self.actions():
            return
        available_width = self.width() - 20
        button_count = sum(1 for action in self.actions() if action.isVisible())
        if button_count == 0:
            return
        button_width = max(50, available_width // button_count)
        for action in self.actions():
            widget = action.defaultWidget()
            if widget:
                widget.setMinimumWidth(button_width)
                widget.setMaximumWidth(button_width)

def initialize_db(db_path='data.db'):
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS prequal (
                id INTEGER PRIMARY KEY,
                folder_path TEXT,
                data TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS blacklist (
                id INTEGER PRIMARY KEY,
                dtcCode TEXT,
                genericSystemName TEXT,
                dtcDescription TEXT,
                dtcSys TEXT,
                carMake TEXT,
                comments TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS goldlist (
                id INTEGER PRIMARY KEY,
                dtcCode TEXT,
                genericSystemName TEXT,
                dtcDescription TEXT,
                dtcSys TEXT,
                carMake TEXT,
                comments TEXT
            );
        ''')
        conn.commit()
    except sqlite3.Error as e:
        logging.error(f"Failed to initialize database tables: {e}")
    finally:
        conn.close()

def update_configuration(config_type, folder_path, data, db_path='data.db'):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    data_json = json.dumps(data)
    cursor.execute(f'''
        INSERT INTO {config_type} (folder_path, data)
        VALUES (?, ?)
    ''', (folder_path, data_json))
    conn.commit()
    conn.close()

def load_configuration(config_type, db_path='data.db'):
    conn = sqlite3.connect(db_path)
    result = []
    try:
        cursor = conn.cursor()
        query = f'SELECT data FROM {config_type}'
        logging.debug(f"Executing query: {query}")
        cursor.execute(query)
        data = cursor.fetchall()
        if data:
            logging.debug(f"Data retrieved from {config_type}: {data[:3]}...")
        for item in data:
            try:
                entries = json.loads(item[0])
                for entry in entries:
                    entry['Make'] = str(entry['Make']) if pd.notna(entry['Make']) else "Unknown"
                result.extend(entries)
            except json.JSONDecodeError as je:
                logging.error(f"JSON decoding error for item {item[0]}: {je}")
    except sqlite3.Error as e:
        logging.error(f"SQLite error encountered when loading configuration for {config_type}: {e}")
    finally:
        conn.close()
        logging.debug(f"Connection to database '{db_path}' closed.")
        return result

def load_excel_data_to_db(excel_path, table_name, db_path='data.db', sheet_name=1, parent=None):
    error_messages = []
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        logging.debug(f"Original data loaded: {df.head()}")

        df = df.rename(columns={
            'Generic System Name': 'genericSystemName',
            'dtcCode': 'dtcCode', 
            'dtcDescription': 'dtcDescription', 
            'dtcSys': 'dtcSys', 
            'CarMake': 'carMake', 
            'Comments': 'comments'
        })

        expected_columns = ['genericSystemName', 'dtcCode', 'dtcDescription', 'dtcSys', 'carMake', 'comments']
        df = df[expected_columns]

        df = df.where(pd.notnull(df), None)

        int_columns = ['dtcCode']
        df[int_columns] = df[int_columns].astype(str)

        conn = sqlite3.connect(db_path)
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        return "Data loaded successfully"
    except Exception as e:
        error_message = f"Failed to load data from {excel_path} into {table_name}: {str(e)}"
        logging.error(error_message)
        error_messages.append(error_message)

    if error_messages:
        QMessageBox.critical(parent, "Errors Encountered", "\n".join(error_messages))
        return "\n".join(error_messages)
    return "Data loaded successfully"

class App(QMainWindow):
    def __init__(self, db_path='data.db'):
        super().__init__()
        self.db_path = db_path
        self.settings_file = 'settings.json'
        self.current_theme = self.load_settings().get('theme', 'Dark')
        initialize_db(self.db_path)
        self.data = {'blacklist': [], 'goldlist': [], 'prequal': []}
        self.setup_ui()
        self.load_configurations()
        self.check_data_loaded()
        self.apply_selected_theme()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r') as file:
                return json.load(file)
        return {}

    def save_settings(self, settings):
        with open(self.settings_file, 'w') as file:
            json.dump(settings, file)

    def check_data_loaded(self):
        if not self.data['prequal']:
            self.make_dropdown.setDisabled(True)
            self.model_dropdown.setDisabled(True)
            self.year_dropdown.setDisabled(True)
            QMessageBox.information(self, "Load Data", "Please load data through the Admin console to enable functionality.")
        else:
            self.make_dropdown.setDisabled(False)
            self.model_dropdown.setDisabled(False)
            self.year_dropdown.setDisabled(False)

    def setup_ui(self):
        self.setWindowTitle("Analyzer+")
        self.setGeometry(100, 100, 900, 600)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)

        self.toolbar = ResizableToolBar(self)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.add_toolbar_button("Admin", self.open_admin, "admin_button")
        self.add_toolbar_button("Export", self.export_data, "export_button")
        self.add_pin_on_top_button()

        self.theme_dropdown = QComboBox()
        themes = ["Dark", "Light", "Red", "Blue", "Green", "Yellow", "Pink", "Purple", "Teal", "Cyan", "Orange"]
        self.theme_dropdown.addItems(themes)
        self.theme_dropdown.currentIndexChanged.connect(self.apply_selected_theme)
        self.toolbar.addWidget(self.theme_dropdown)

        # Create a horizontal layout for the dropdowns
        dropdown_layout = QHBoxLayout()

        self.make_dropdown = QComboBox()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        self.make_dropdown.currentIndexChanged.connect(self.update_model_dropdown)
        self.make_dropdown.currentIndexChanged.connect(self.clear_search_bar)
        self.make_dropdown.focusInEvent = self.on_make_dropdown_focus  # Connect focus event
        dropdown_layout.addWidget(self.make_dropdown)

        self.model_dropdown = QComboBox()
        self.model_dropdown.addItem("Select Model")
        self.model_dropdown.currentIndexChanged.connect(self.handle_model_change)
        dropdown_layout.addWidget(self.model_dropdown)

        self.year_dropdown = QComboBox()
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.currentIndexChanged.connect(self.perform_search)
        dropdown_layout.addWidget(self.year_dropdown)

        main_layout.addLayout(dropdown_layout)  # Add the dropdown layout to the top of the main layout

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Enter DTC code or description")
        self.search_bar.returnPressed.connect(self.perform_search)
        main_layout.addWidget(self.search_bar)

        self.filter_dropdown = QComboBox()
        self.filter_dropdown.addItems(["Select List", "All", "Blacklist", "Goldlist", "Gold and Black", "Prequals"])
        self.filter_dropdown.currentIndexChanged.connect(self.perform_search)
        main_layout.addWidget(self.filter_dropdown)

        self.suggestions_list = QListWidget(self)
        self.suggestions_list.setMaximumHeight(100)
        self.suggestions_list.hide()
        main_layout.addWidget(self.suggestions_list)

        # Add a horizontal line separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)

        self.splitter = QSplitter(Qt.Horizontal)

        self.left_panel = QTextBrowser()
        self.splitter.addWidget(self.left_panel)

        self.right_panel = QTextBrowser()
        self.splitter.addWidget(self.right_panel)

        main_layout.addWidget(self.splitter)

        # Add transparency bar
        transparency_layout = QHBoxLayout()
        self.opacity_label = QLabel("Transparency")
        self.opacity_slider = QSlider(Qt.Horizontal)
        self.opacity_slider.setMinimum(20)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.valueChanged.connect(self.change_opacity)
        transparency_layout.addWidget(self.opacity_label)
        transparency_layout.addWidget(self.opacity_slider)
        main_layout.addLayout(transparency_layout)

        self.apply_selected_theme()
        self.populate_dropdowns()

    def on_make_dropdown_focus(self, event):
        # Do nothing if "All" is selected
        if self.make_dropdown.currentText() == "All":
            return
        # Set list dropdown to default "Select List" when focus is on make dropdown
        self.filter_dropdown.setCurrentText("Select List")
        QComboBox.focusInEvent(self.make_dropdown, event)

    def populate_search_make_dropdown(self):
        self.make_dropdown.clear()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        makes = sorted(set(item['Make'].strip() for item in self.data['prequal'] if item['Make'].strip() and item['Make'].strip().lower() != 'unknown'))
        self.make_dropdown.addItems(makes)

    def reset_filter_dropdown(self):
        self.filter_dropdown.setCurrentIndex(0)

    def make_dropdown_focus_in(self, event):
        QComboBox.focusInEvent(self.make_dropdown, event)

    def add_toolbar_button(self, text, slot, object_name):
        button = QPushButton(text)
        button.setObjectName(object_name)
        button.clicked.connect(slot)
        self.toolbar.addWidget(button)

    def add_pin_on_top_button(self):
        button = QPushButton("Pin on Top")
        button.setObjectName("pin_on_top_button")
        button.setCheckable(True)
        button.toggled.connect(self.toggle_always_on_top)
        button.setStyleSheet("""
            QPushButton {
                border: 2px solid #FFFF00;
                padding: 4px;
            }
            QPushButton:checked {
                border-color: #00FF00;
                border-style: solid;
                border-width: 2px;
            }
        """)
        self.toolbar.addWidget(button)

    def toggle_always_on_top(self, checked):
        flags = self.windowFlags()
        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        self.show()

    def update_suggestions(self, text):
        if text:
            suggestions = self.get_suggestions(text)
            self.suggestions_list.clear()
            if suggestions:
                for suggestion in suggestions:
                    self.suggestions_list.addItem(suggestion)
                self.suggestions_list.show()
                logging.debug("Suggestions list updated and shown.")
            else:
                self.suggestions_list.hide()
                logging.debug("No suggestions found, list hidden.")
        else:
            self.suggestions_list.hide()
            logging.debug("Input cleared, list hidden.")

    def get_suggestions(self, text):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        query = """
        SELECT dtcCode || ' - ' || dtcDescription AS suggestion
        FROM (
            SELECT dtcCode, dtcDescription FROM blacklist
            UNION
            SELECT dtcCode, dtcDescription FROM goldlist
        )
        WHERE dtcCode LIKE ? OR dtcDescription LIKE ?
        """
        cursor.execute(query, ('%' + text + '%', '%' + text + '%'))
        suggestions = [row[0] for row in cursor.fetchall()]
        conn.close()
        logging.debug(f"Suggestions fetched: {suggestions}")
        return suggestions

    def on_suggestion_clicked(self):
        item = self.suggestions_list.currentItem()
        if item:
            self.search_bar.setText(item.text())
            self.suggestions_list.hide()

    def auto_perform_search(self, index):
        if index != 0:
            self.perform_search()

    def handle_model_change(self, index):
        if index != 0:
            self.update_year_dropdown()

    def toggle_theme(self):
        if self.dark_theme_enabled:
            self.apply_light_theme()
        else:
            self.apply_dark_theme()
        self.dark_theme_enabled = not self.dark_theme_enabled
    
    def open_link(self, url):
        url_string = url.toString()
        if url_string.startswith('click:'):
            clean_url = url_string[6:]
            actual_url = QUrl(clean_url)
            logging.debug(f"Clean URL: {actual_url.toString()}")
        else:
            actual_url = url

        if not QDesktopServices.openUrl(actual_url):
            logging.error(f"Failed to open URL: {actual_url.toString()}")
        else:
            logging.info(f"URL opened successfully: {actual_url.toString()}")

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r') as file:
                return json.load(file)
        return {"theme": "Dark"}  # Default to Dark theme if no settings file exists

    def apply_selected_theme(self):
        theme = self.theme_dropdown.currentText()
        self.current_theme = theme
        self.save_settings({'theme': self.current_theme})

        if theme == "Red":
            self.apply_color_theme("#DC143C", "#A52A2A")
        elif theme == "Blue":
            self.apply_color_theme("#4169E1", "#1E3A5F")
        elif theme == "Dark":
            self.apply_dark_theme()
        elif theme == "Light":
            self.apply_color_theme("#f0f0f0", "#d0d0d0", text_color="black")
        elif theme == "Green":
            self.apply_color_theme("#006400", "#003a00")
        elif theme == "Yellow":
            self.apply_color_theme("#ffd700", "#b29500", text_color="black")
        elif theme == "Pink":
            self.apply_color_theme("#ff69b4", "#b2477d", text_color="black")
        elif theme == "Purple":
            self.apply_color_theme("#800080", "#4b004b")
        elif theme == "Teal":
            self.apply_color_theme("#008080", "#004b4b")
        elif theme == "Cyan":
            self.apply_color_theme("#00ffff", "#00b2b2", text_color="black")
        elif theme == "Orange":
            self.apply_color_theme("#ff8c00", "#ff4500", text_color="black")

    def apply_color_theme(self, color_start, color_end, text_color="#ddd"):
        style = f"""
        QWidget {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
        }}
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {{
            font: bold 11px;
            color: {text_color};
        }}
        QPushButton {{
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
        }}
        QPushButton:hover {{
            background-color: #555;
        }}
        QPushButton:pressed {{
            background-color: #666;
            border-style: inset;
        }}
        QPushButton#search_button {{
            color: white;
        }}
        QPushButton#admin_button, QPushButton#export_button, QPushButton#pin_on_top_button {{
            color: white;
        }}
        QTextBrowser {{
            background-color: #323232;
            border: 1px solid #444;
            color: white;
        }}
        QComboBox {{
            border: 2px solid #555;
            border-radius: 10px;
            padding: 1px 18px 1px 3px;
            min-width: 6em;
        }}
        QComboBox:hover {{
            border: 2px solid #aaa;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 15px;
            border-left-width: 1px;
            border-left-color: darkgray;
            border-left-style: solid;
            border-top-right-radius: 3px;
            border-bottom-right-radius: 3px;
        }}
        QComboBox QAbstractItemView {{
            border: 2px solid darkgray;
            selection-background-color: lightgray;
        }}
        QLineEdit {{
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
        }}
        QLineEdit:hover {{
            border: 2px solid #aaa;
        }}
        QListWidget {{
            border: 1px solid #444;
            border-radius: 5px;
            background-color: #323232;
            color: {text_color};
        }}
        QListWidget::item {{
            border-bottom: 1px solid #555;
        }}
        QListWidget::item:selected {{
            background-color: #555;
        }}
        QStatusBar {{
            border-top: 1px solid #444;
            background: #404040;
        }}
        QToolBar {{
            background-color: #404040;
            border-bottom: 1px solid #444;
        }}
        QSlider::groove:horizontal {{
            border: 1px solid #bbb;
            background: white;
            height: 5px;
            border-radius: 4px;
        }}
        QSlider::sub-page:horizontal {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #66e, stop:1 #bbf);
            background: qlineargradient(x1: 0, y1: 0.2, x2: 1, y2: 1, stop: 0 #5DCCFF, stop: 1 #5DCCFF);
            border: 1px solid #777;
            height: 10px;
            border-radius: 4px;
        }}
        QSlider::add-page:horizontal {{
            background: #fff;
            border: 1px solid #777;
            height: 10px;
            border-radius: 4px;
        }}
        QSlider::handle:horizontal {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #eee, stop:1 #ccc);
            border: 1px solid #777;
            width: 18px;
            margin-top: -4px;
            margin-bottom: -4px;
            border-radius: 4px;
        }}
        QSlider::handle:horizontal:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #fff, stop:1 #ddd);
            border: 1px solid #444;
            border-radius: 4px;
        }}
        QSlider::sub-page:horizontal:disabled {{
            background: #bbb;
            border-color: #999;
        }}
        QSlider::add-page:horizontal:disabled {{
            background: #eee;
            border-color: #999;
        }}
        QSlider::handle:horizontal:disabled {{
            background: #eee;
            border: 1px solid #aaa;
            border-radius: 4px;
        }}
        """
        self.apply_stylesheet(style)

    def apply_dark_theme(self):
        style = """
        QWidget {
            background-color: #2b2b2b;
            color: #ddd;
        }
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {
            font: bold 11px;
            color: #ddd;
        }
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        QPushButton#search_button {
            color: white;
        }
        QPushButton#admin_button, QPushButton#export_button, QPushButton#pin_on_top_button {
            color: white;
        }
        QTextBrowser {
            background-color: #323232;
            border: 1px solid #444;
            color: white;
        }
        QComboBox {
            border: 2px solid #555;
            border-radius: 10px;
            padding: 1px 18px 1px 3px;
            min-width: 6em;
        }
        QComboBox:hover {
            border: 2px solid #aaa;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 15px;
            border-left-width: 1px;
            border-left-color: darkgray;
            border-left-style: solid;
            border-top-right-radius: 3px;
            border-bottom-right-radius: 3px;
        }
        QComboBox QAbstractItemView {
            border: 2px solid darkgray;
            selection-background-color: lightgray;
        }
        QLineEdit {
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
        }
        QLineEdit:hover {
            border: 2px solid #aaa;
        }
        QTextBrowser {
            background-color: #323232;
            border: 1px solid #444;
        }
        QListWidget {
            border: 1px solid #444;
            border-radius: 5px;
            background-color: #323232;
            color: #ddd;
        }
        QListWidget::item {
            border-bottom: 1px solid #555;
        }
        QListWidget::item:selected {
            background-color: #555;
        }
        QStatusBar {
            border-top: 1px solid #444;
            background: #404040;
        }
        QToolBar {
            background-color: #404040;
            border-bottom: 1px solid #444;
        }
        QSlider::groove:horizontal {
            border: 1px solid #bbb;
            background: white;
            height: 5px;
            border-radius: 4px;
        }
        QSlider::sub-page:horizontal {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #66e, stop:1 #bbf);
            background: qlineargradient(x1: 0, y1: 0.2, x2: 1, y2: 1, stop: 0 #5DCCFF, stop: 1 #5DCCFF);
            border: 1px solid #777;
            height: 10px;
            border-radius: 4px;
        }
        QSlider::add-page:horizontal {
            background: #fff;
            border: 1px solid #777;
            height: 10px;
            border-radius: 4px;
        }
        QSlider::handle:horizontal {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #eee, stop:1 #ccc);
            border: 1px solid #777;
            width: 18px;
            margin-top: -4px;
            margin-bottom: -4px;
            border-radius: 4px;
        }
        QSlider::handle:horizontal:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #fff, stop:1 #ddd);
            border: 1px solid #444;
            border-radius: 4px;
        }
        QSlider::sub-page:horizontal:disabled {
            background: #bbb;
            border-color: #999;
        }
        QSlider::add-page:horizontal:disabled {
            background: #eee;
            border-color: #999;
        }
        QSlider::handle:horizontal:disabled {
            background: #eee;
            border: 1px solid #aaa;
            border-radius: 4px;
        }
        """
        self.apply_stylesheet(style)

    def apply_stylesheet(self, style):
        self.central_widget.setStyleSheet(style)
        self.toolbar.setStyleSheet(style)
        self.status_bar.setStyleSheet(style)

    def change_opacity(self, value):
        self.setWindowOpacity(value / 100.0)

    def clear_data(self, config_type=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            if config_type:
                cursor.execute(f"DELETE FROM {config_type}")
                logging.info(f"Data cleared from {config_type}")
            else:
                cursor.execute("DROP TABLE IF EXISTS blacklist")
                cursor.execute("DROP TABLE IF EXISTS goldlist")
                cursor.execute("DROP TABLE IF EXISTS prequal")
                initialize_db(self.db_path)
                logging.info("Database reset complete.")
            conn.commit()
        except sqlite3.Error as e:
            logging.error(f"Failed to clear data: {e}")
            QMessageBox.critical(self, "Error", "Failed to clear database.")
        finally:
            conn.close()
        self.load_configurations()

    def export_data(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;Text Files (*.txt);;CSV Files (*.csv);;JSON Files (*.json)", options=options)
        if fileName:
            if fileName.endswith('.csv'):
                self.export_to_csv(fileName)
            elif fileName.endswith('.json'):
                self.export_to_json(fileName)

    def export_to_csv(self, path):
        conn = sqlite3.connect(self.db_path)
        for table in ['blacklist', 'goldlist', 'prequal']:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            df.to_csv(f"{path}_{table}.csv", index=False)
        conn.close()
        QMessageBox.information(self, "Export Successful", "Data exported successfully to CSV.")

    def export_to_json(self, path):
        conn = sqlite3.connect(self.db_path)
        for table in ['blacklist', 'goldlist', 'prequal']:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            df.to_json(f"{path}_{table}.json", orient='records', lines=True)
        conn.close()
        QMessageBox.information(self, "Export Successful", "Data exported successfully to JSON.")

    def update_model_dropdown(self):
        selected_make = self.make_dropdown.currentText().strip()
        if selected_make == "All":
            self.left_panel.clear()  # Clear the prequal display box
        elif selected_make != "Select Make":
            models = [item['Model'] for item in self.data['prequal'] if item['Make'] == selected_make]
            unique_models = sorted(set(str(model) for model in models))
            self.model_dropdown.clear()
            self.model_dropdown.addItem("Select Model")
            self.model_dropdown.addItems(unique_models)

    def handle_model_change(self, index):
        selected_model = self.model_dropdown.currentText().strip()
        if selected_model != "Select Model":
            self.update_year_dropdown()

    def open_admin(self):
        password, ok = QInputDialog.getText(self, 'Admin Login', 'Enter password:', QLineEdit.Password)
        if ok and password == "ADAS":
            choice, ok = QInputDialog.getItem(self, "Admin Actions", "Select action:", ["Update Paths", "Clear Data"], 0, False)
            if ok and choice == "Clear Data":
                config_type, ok = QInputDialog.getItem(self, "Clear Data", "Select configuration to clear or 'All' to reset database:", ["blacklist", "goldlist", "prequal", "All"], 0, False)
                if ok:
                    if config_type == "All":
                        self.clear_data()
                    else:
                        self.clear_data(config_type)
            elif ok:
                self.manage_paths()
        else:
            QMessageBox.warning(self, "Error", "Incorrect password")

    def update_year_dropdown(self):
        selected_make = self.make_dropdown.currentText().strip()
        selected_model = self.model_dropdown.currentText().strip()
        if selected_make != "Select Make" and selected_model != "Select Model":
            years = [str(item['Year']) for item in self.data['prequal'] if item['Make'] == selected_make and item['Model'] == selected_model]
            unique_years = sorted(set(years))
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")
            self.year_dropdown.addItems(unique_years)
        else:
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")

    def get_valid_excel_files(self, folder_path):
        file_pattern = re.compile(r'(.+)\.xlsx$', re.IGNORECASE)
        skip_pattern = re.compile(r'.*X\.X\.xlsx$', re.IGNORECASE)
        valid_files = {}
        for file_name in os.listdir(folder_path):
            if file_pattern.match(file_name) and not skip_pattern.match(file_name):
                full_path = os.path.join(folder_path, file_name)
                valid_files[file_name] = full_path
        if not valid_files:
            logging.error("No valid Excel files found.")
        return valid_files

    def manage_paths(self):
        config_type, ok = QInputDialog.getItem(self, "Select Config Type", "Choose the configuration to update:", ["blacklist", "goldlist", "prequal"], 0, False)
        if ok:
            folder_path = QFileDialog.getExistingDirectory(self, "Select Directory")
            if folder_path:
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    QMessageBox.warning(self, "Load Error", "No valid Excel files found in the directory.")
                    return

                data_loaded = False
                for filename, filepath in files.items():
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, sheet_name=1)
                            if result != "Failed to load data":
                                data_loaded = True
                        else:
                            df = pd.read_excel(filepath)
                            if df.empty:
                                logging.warning(f"{filename} is empty.")
                                continue
                            data = df.to_dict(orient='records')
                            update_configuration(config_type, folder_path, data, self.db_path)
                            data_loaded = True
                    except Exception as e:
                        logging.error(f"Error loading {filename}: {str(e)}")
                        QMessageBox.critical(self, "Load Error", f"Failed to load {filename}: {str(e)}")

                if data_loaded:
                    self.load_configurations()
                    QMessageBox.information(self, "Update Successful", "Data loaded successfully.")
                    self.status_bar.showMessage(f"Data updated from: {folder_path}")
                else:
                    QMessageBox.warning(self, "Load Error", "No data was loaded from the files. Check the logs for more details.")

    def load_configurations(self):
        logging.debug("Loading configurations...")
        for config_type in ['blacklist', 'goldlist', 'prequal']:
            data = load_configuration(config_type, self.db_path)
            self.data[config_type] = data if data else []
            logging.debug(f"Loaded {len(data)} items for {config_type}")
        if 'prequal' in self.data:
            self.populate_dropdowns()

    def populate_dropdowns(self):
        logging.debug("Populating dropdowns...")
        self.make_dropdown.clear()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")

        makes = sorted(set(item['Make'].strip() for item in self.data['prequal'] if item['Make'].strip() and item['Make'].strip().lower() != 'unknown'))

        if '' in makes:
            makes.remove('')

        self.make_dropdown.addItems(makes)
        logging.debug(f"Makes populated: {makes}")

    def perform_search(self):
        dtc_code = self.search_bar.text().strip().upper()
        selected_filter = self.filter_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()

        # Display nothing if the default option in the list dropdown is selected
        if selected_filter == "Select List":
            return

        # Handle case where "All" is selected in list dropdown but no make is selected
        if selected_filter == "All" and (selected_make == "Select Make" or selected_make == "All"):
            QMessageBox.warning(self, "Selection Error", "Please select a vehicle for prequal information")
            self.filter_dropdown.setCurrentText("Select List")
            return

        # Handle prequals selection
        if selected_filter == "Prequals":
            if selected_make == "Select Make" or selected_make == "All":
                QMessageBox.warning(self, "Selection Error", "Please select a vehicle for prequal information")
                self.filter_dropdown.setCurrentText("Select List")
                return
            results = []
            self.splitter.widget(0).show()
            self.splitter.widget(1).hide()  # Hide the right panel
            self.splitter.widget(0).setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            if selected_make != "Select Make" and selected_model != "Select Model" and selected_year != "Select Year":
                results = [item for item in self.data['prequal']
                        if item['Make'] == selected_make and item['Model'] == selected_model and str(item['Year']) == selected_year]
            elif selected_make == "All":
                results = self.data['prequal']
            self.display_results(results, context='prequal')

        # Handle Gold and Black, Blacklist, Goldlist selections
        elif selected_filter in ["Gold and Black", "Blacklist", "Goldlist"]:
            if not dtc_code and selected_make == "All":
                self.right_panel.setPlainText("Please enter a DTC code or description to search.")
                return
            self.splitter.widget(0).hide()  # Hide the left panel
            self.splitter.widget(1).show()
            self.splitter.widget(1).setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.splitter.widget(1).resize(self.splitter.size())
            self.search_dtc_codes(dtc_code, selected_filter, selected_make)

        # Handle All selection
        elif selected_filter == "All":
            # Display prequal results in the left panel
            self.splitter.widget(0).show()
            self.splitter.widget(1).show()
            self.splitter.widget(0).setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.splitter.widget(1).setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            prequal_results = []
            if selected_make != "Select Make" and selected_model != "Select Model" and selected_year != "Select Year":
                prequal_results = [item for item in self.data['prequal']
                                if item['Make'] == selected_make and item['Model'] == selected_model and str(item['Year']) == selected_year]
            elif selected_make == "All":
                prequal_results = self.data['prequal']
            self.display_results(prequal_results, context='prequal')

            # Display gold and black list results in the right panel
            if dtc_code:
                self.search_dtc_codes(dtc_code, "Gold and Black", selected_make)
            else:
                self.display_gold_and_black(selected_make)

        else:
            self.right_panel.setPlainText("Please select a valid option.")

    def search_dtc_codes(self, dtc_code, filter_type, selected_make):
        conn = sqlite3.connect(self.db_path)
        query = ""

        if filter_type == "All" or filter_type == "Gold and Black":
            if selected_make == "All":
                query = f"""
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                WHERE dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%'
                UNION ALL
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
                WHERE dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%'
                """
            else:
                query = f"""
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%') AND carMake = '{selected_make}'
                UNION ALL
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
                WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%') AND carMake = '{selected_make}'
                """
        elif filter_type == "Blacklist":
            query = f"""
            SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
            WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%')
            """
            if selected_make != "Select Make" and selected_make != "All":
                query += f" AND carMake = '{selected_make}'"
        elif filter_type == "Goldlist":
            query = f"""
            SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
            WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%')
            """
            if selected_make != "Select Make" and selected_make != "All":
                query += f" AND carMake = '{selected_make}'"

        query += ";"  # Ensuring the query ends with a semicolon.

        try:
            df = pd.read_sql_query(query, conn)
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.right_panel.setPlainText("An error occurred while fetching the data.")
            return

        conn.close()

        if not df.empty:
            self.right_panel.setHtml(df.to_html(index=False, escape=False))
        else:
            self.right_panel.setPlainText("No DTC code results found.")

    def display_gold_and_black(self, selected_make):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if selected_make == "All":
            query = """
            SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
            UNION ALL
            SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
            """
        else:
            query = f"""
            SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist WHERE carMake = '{selected_make}'
            UNION ALL
            SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist WHERE carMake = '{selected_make}'
            """
        
        try:
            df = pd.read_sql_query(query, conn)
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.right_panel.setPlainText("An error occurred while fetching the data.")
            return

        conn.close()

        if not df.empty:
            self.right_panel.setHtml(df.to_html(index=False, escape=False))
        else:
            self.right_panel.setPlainText("No DTC code results found.")

    def display_results(self, results, context='prequal'):
        display_text = ""
        for result in results:
            calibration_type = result.get('Calibration Type', 'N/A')
            if isinstance(calibration_type, float) and pd.isna(calibration_type):
                continue
            elif isinstance(calibration_type, str) and calibration_type.lower() == 'nan':
                continue

            link = result.get('Service Information Hyperlink', '#')
            if isinstance(link, float) and pd.isna(link):
                link = '#'
            elif not link.startswith(('http://', 'https://')):
                link = 'http://' + link

            system_acronym = result.get('Protech Generic System Name.1', 'N/A')
            parts_code = result.get('Parts Code Table Value', 'N/A')
            calibration_prerequisites = result.get('Calibration Pre-Requisites', 'N/A')

            if context == 'dtc':
                display_text += f"""
                <b>Source:</b> {result['Source']}<br>
                <b>DTC Code:</b> {result['dtcCode']}<br>
                <b>Generic System Name:</b> {result['genericSystemName']}<br>
                <b>Description:</b> {result['dtcDescription']}<br>
                <b>DTC System:</b> {result['dtcSys']}<br>
                <b>Car Make:</b> {result['carMake']}<br>
                <b>Comments:</b> {result['comments']}<br>
                <b>Service Information:</b> <a href='{link}'>Click Here</a><br><br>
                """
            else:
                display_text += f"""
                <b>Make:</b> {result['Make']}<br>
                <b>Model:</b> {result['Model']}<br>
                <b>Year:</b> {result['Year']}<br>
                <b>System Acronym:</b> {system_acronym}<br>
                <b>Parts Code Table Value:</b> {parts_code}<br>
                <b>Calibration Type:</b> {calibration_type}<br>
                <b>Service Information:</b> <a href='{link}'>Click Here</a><br><br>
                <b>Pre-Quals:</b> {calibration_prerequisites}<br><br>
                """
            self.left_panel.setHtml(display_text)
            self.left_panel.setOpenExternalLinks(True)

    def clear_search_bar(self):
        self.search_bar.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())
