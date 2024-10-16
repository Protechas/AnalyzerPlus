import csv
import shutil
import sys
import os
import json
import re
import sqlite3
import logging
import threading
import pandas as pd
import pytz
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QDialog, QFrame, QMainWindow, QProgressBar, QToolBar, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit, QPushButton, QComboBox, QTextBrowser, QMessageBox, QStatusBar, QInputDialog, QFileDialog, QSlider, QLabel, QListWidget, QSpacerItem, QSizePolicy, QSplitter)
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon, QDesktopServices

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class PopOutWindow(QDialog):
    def __init__(self, title, content, stylesheet, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 600, 400)
        
        layout = QVBoxLayout()
        
        self.text_browser = QTextBrowser()
        self.text_browser.setHtml(content)
        layout.addWidget(self.text_browser)
        
        # Add transparency slider
        transparency_layout = QHBoxLayout()
        self.opacity_label = QLabel("Transparency")
        self.opacity_slider = QSlider(Qt.Horizontal)
        self.opacity_slider.setMinimum(20)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.valueChanged.connect(self.change_opacity)
        transparency_layout.addWidget(self.opacity_label)
        transparency_layout.addWidget(self.opacity_slider)
        layout.addLayout(transparency_layout)

        self.setLayout(layout)
        self.setStyleSheet(stylesheet)  # Apply the stylesheet

    def change_opacity(self, value):
        self.setWindowOpacity(value / 100.0)

class PinManagementDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Users and PINs")
        self.setGeometry(100, 100, 400, 300)
        self.db_path = parent.db_path
        self.setup_ui()
        self.load_users_and_pins()

    def setup_ui(self):
        layout = QVBoxLayout()

        self.user_list = QListWidget()
        layout.addWidget(self.user_list)

        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add")
        self.add_button.clicked.connect(self.add_user)
        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_user)
        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_user)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def load_users_and_pins(self):
        self.user_list.clear()
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT pin, name FROM leader_log')
        users = cursor.fetchall()
        for pin, name in users:
            self.user_list.addItem(f"{pin} - {name}")

    def add_user(self):
        pin, ok = QInputDialog.getText(self, 'Add User', 'Enter new PIN:')
        if ok and pin:
            name, ok = QInputDialog.getText(self, 'Add User', 'Enter user name:')
            if ok and name:
                self.save_user(pin, name)

    def edit_user(self):
        current_item = self.user_list.currentItem()
        if current_item:
            pin, name = current_item.text().split(" - ")
            new_pin, ok = QInputDialog.getText(self, 'Edit User', 'Edit PIN:', QLineEdit.Normal, pin)
            if ok and new_pin:
                new_name, ok = QInputDialog.getText(self, 'Edit User', 'Edit user name:', QLineEdit.Normal, name)
                if ok and new_name:
                    self.save_user(new_pin, new_name, old_pin=pin)

    def delete_user(self):
        current_item = self.user_list.currentItem()
        if current_item:
            pin, name = current_item.text().split(" - ")
            conn = self.parent().get_db_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM leader_log WHERE pin = ?', (pin,))
            conn.commit()
            self.load_users_and_pins()
            self.parent().log_action(self.parent().current_user, f"User removed: pin={pin}, name={name}")

    def save_user(self, pin, name, old_pin=None):
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        if old_pin:
            cursor.execute('UPDATE leader_log SET pin = ?, name = ? WHERE pin = ?', (pin, name, old_pin))
        else:
            cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', (pin, name))
        conn.commit()
        self.load_users_and_pins()
        self.parent().log_action(self.parent().current_user, f"User added or updated: pin={pin}, name={name}")

class AdminOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Admin Options")
        self.setGeometry(100, 100, 300, 150)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        self.update_users_button = QPushButton("Update Users")
        self.update_users_button.clicked.connect(self.update_users)
        layout.addWidget(self.update_users_button)

        self.export_data_button = QPushButton("Export Data")
        self.export_data_button.clicked.connect(self.export_data)
        layout.addWidget(self.export_data_button)

        self.setLayout(layout)

    def update_users(self):
        self.accept()
        self.parent().open_pin_management()

    def export_data(self):
        self.accept()
        self.parent().export_data()

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
            CREATE TABLE IF NOT EXISTS carsys (
                id INTEGER PRIMARY KEY,
                genericSystemName TEXT,
                dtcSys TEXT,
                carMake TEXT,
                comments TEXT
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
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS mag_glass (
                id INTEGER PRIMARY KEY,
                genericSystemName TEXT,
                adasModuleName TEXT,
                carMake TEXT,
                manufacturer TEXT,
                autelOrBosch TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS leader_log (
                id INTEGER PRIMARY KEY,
                pin TEXT,
                name TEXT
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_actions (
                id INTEGER PRIMARY KEY,
                user TEXT,
                action TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS paths (
                config_type TEXT PRIMARY KEY,
                folder_path TEXT
            );
        ''')

        # Insert the "Set Up" user if it doesn't exist
        cursor.execute('SELECT * FROM leader_log WHERE name = "Set Up"')
        if not cursor.fetchone():
            cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', ('0000', 'Set Up'))

        # Create indexes on frequently queried columns
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_leader_log_pin ON leader_log (pin)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_leader_log_name ON leader_log (name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_blacklist_dtcCode ON blacklist (dtcCode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_blacklist_carMake ON blacklist (carMake)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_goldlist_dtcCode ON goldlist (dtcCode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_goldlist_carMake ON goldlist (carMake)')

        conn.commit()
    except sqlite3.Error as e:
        logging.error(f"Failed to initialize database tables: {e}")
    finally:
        conn.close()

def get_db_connection(self):
    return sqlite3.connect(self.db_path)

def save_path_to_db(config_type, folder_path, db_path='data.db'):
    conn = sqlite3.connect(db_path)
    try:
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO paths (config_type, folder_path)
            VALUES (?, ?)
        ''', (config_type, folder_path))
        conn.commit()

        # Normalize paths to avoid issues with backslashes
        backup_path = os.path.normpath(os.path.join(folder_path, f"{config_type}_backup"))
        folder_path = os.path.normpath(folder_path)

        if os.path.exists(backup_path):
            logging.warning(f"Backup path already exists: {backup_path}. Removing existing backup.")
            shutil.rmtree(backup_path)

        logging.info(f"Copying from {folder_path} to {backup_path}.")
        try:
            shutil.copytree(folder_path, backup_path, symlinks=False, ignore_dangling_symlinks=True)
        except shutil.Error as e:
            logging.error(f"Error during copy operation: {e}")
        except OSError as e:
            logging.error(f"OS error during copy operation: {e}")
    except sqlite3.Error as e:
        logging.error(f"Failed to save path to database: {e}")
    finally:
        conn.close()

def handle_error(func, path, exc_info):
    logging.error(f"Error copying {path}: {exc_info}")

def load_path_from_db(config_type, db_path='data.db'):
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT folder_path FROM paths WHERE config_type = ?
        ''', (config_type,))
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error as e:
        logging.error(f"Failed to load path from database: {e}")
        return None
    finally:
        conn.close()

def update_configuration(config_type, folder_path, data, db_path='data.db'):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    if config_type == 'mag_glass':
        # Assuming columns A, B, C, D, and E correspond to specific fields
        cursor.executemany('''
            INSERT INTO mag_glass (genericSystemName, adasModuleName, carMake, manufacturer, autelOrBosch)
            VALUES (?, ?, ?, ?, ?)
        ''', [(item['genericSystemName'], item['adasModuleName'], item['carMake'], item['manufacturer'], item['autelOrBosch']) for item in data])

    else:
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
            logging.debug(f"Data retrieved from {config_type}: {data[:3]}...")  # Show first 3 entries for brevity
        for item in data:
            try:
                entries = json.loads(item[0])
                for entry in entries:
                    entry['Make'] = str(entry['Make']).strip() if pd.notna(entry['Make']) else "Unknown"
                    entry['Model'] = str(entry['Model']).strip() if pd.notna(entry['Model']) else "Unknown"
                result.extend(entries)
            except json.JSONDecodeError as je:
                logging.error(f"JSON decoding error for item {item[0]}: {je}")
    except sqlite3.Error as e:
        logging.error(f"SQLite error encountered when loading configuration for {config_type}: {e}")
    finally:
        conn.close()
        return result

def load_excel_data_to_db(excel_path, table_name, db_path='data.db', sheet_index=0, parent=None):
    error_messages = []
    try:
        # Load specific sheet from the Excel file
        df = pd.read_excel(excel_path, sheet_name=sheet_index)
        logging.debug(f"Data loaded from sheet index '{sheet_index}': {df.head()}")

        # Adjust column names and expected columns based on the table name
        if table_name == 'mag_glass':
            df = df.rename(columns={
                'Generic System Name': 'genericSystemName',
                'ADAS Module Name': 'adasModuleName',
                'CarMake': 'carMake',
                'Manufacturer': 'manufacturer',
                'AUTEL or BOSCH': 'autelOrBosch'
            })
            expected_columns = ['genericSystemName', 'adasModuleName', 'carMake', 'manufacturer', 'autelOrBosch']
            df = df[expected_columns]
        else:
            df = df.rename(columns={
                'Generic System Name': 'genericSystemName',
                'DTC Code': 'dtcCode',
                'DTC Description': 'dtcDescription',
                'DTC Sys': 'dtcSys',
                'CarMake': 'carMake',  # Adjusted from 'CarMake' to match Excel
                'Comments': 'comments'
            })
            expected_columns = ['genericSystemName', 'dtcCode', 'dtcDescription', 'dtcSys', 'carMake', 'comments']
            df = df[expected_columns]

        # Clean up the DataFrame: replace NaNs with None and drop rows where all elements are NaN
        df = df.where(pd.notnull(df), None)
        df.dropna(how='all', inplace=True)

        # Convert necessary columns to strings (applicable for non-mag_glass tables)
        if table_name != 'mag_glass':
            int_columns = ['dtcCode']
            df[int_columns] = df[int_columns].astype(str)

        # Save the DataFrame to the specified table in the SQLite database
        conn = sqlite3.connect(db_path)
        df.to_sql(table_name, conn, if_exists='replace', index=False)

        # Verify data saved to the database
        df_saved = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        logging.debug(f"Data saved to {table_name}: {df_saved}")

        # Save the DataFrame to the specified table in the SQLite database
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

def load_carsys_data_to_db(excel_path, table_name='carsys', db_path='data.db', parent=None):
    error_messages = []
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Drop the table if it already exists
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
        conn.commit()

        # Load the first sheet from the Excel file (assuming column headers are in the first row)
        df = pd.read_excel(excel_path, sheet_name=0, usecols="A:D", header=0)
        logging.debug(f"Data loaded from sheet index '0': {df.head()}")

        # Clean column names by stripping any extra whitespace
        df.columns = df.columns.str.strip()

        # Expected columns
        expected_columns = ['Generic System Name', 'DTCsys', 'CarMake', 'Comments']

        # Validate that the necessary columns exist in the DataFrame
        if not all(col in df.columns for col in expected_columns):
            missing_cols = [col for col in expected_columns if col not in df.columns]
            raise ValueError(f"Missing expected columns in the Excel file: {', '.join(missing_cols)}")

        # Rename the columns to match the database schema
        df = df.rename(columns={
            'Generic System Name': 'genericSystemName',
            'DTCsys': 'dtcSys',
            'CarMake': 'carMake',
            'Comments': 'comments'
        })

        # Select only the necessary columns (A, B, C, D)
        df = df[['genericSystemName', 'dtcSys', 'carMake', 'comments']]

        # Convert necessary columns to strings (or other expected types)
        df['genericSystemName'] = df['genericSystemName'].astype(str)
        df['dtcSys'] = df['dtcSys'].astype(str)
        df['carMake'] = df['carMake'].astype(str)
        df['comments'] = df['comments'].astype(str)

        # Clean up the DataFrame: replace NaNs with None and drop rows where all elements are NaN
        df = df.where(pd.notnull(df), None)
        df.dropna(how='all', inplace=True)

        # Save the DataFrame to the specified table in the SQLite database
        df.to_sql(table_name, conn, if_exists='replace', index=False)

        # Verify data saved to the database
        df_saved = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        logging.debug(f"Data saved to {table_name}: {df_saved.head()}")
        conn.close()
        return "Data loaded successfully"

    except Exception as e:
        error_message = f"Failed to load data from {excel_path} into {table_name}: {str(e)}"
        logging.error(error_message)
        error_messages.append(error_message)

    if error_messages:
        QMessageBox.critical(parent, "Errors Encountered", "\n".join(error_messages))
        return "\n.join(error_messages)"

    return "Data loaded successfully"

def load_mag_glass_data_from_5th_sheet(excel_path, table_name='mag_glass', db_path='data.db'):
    error_messages = []
    try:
        # Load the 5th sheet of the Excel file (index starts from 0, so the 5th sheet is index 4)
        df = pd.read_excel(excel_path, sheet_name=4)
        logging.debug(f"Data loaded from the 5th sheet: {df.head()}")

        # Clean up the DataFrame: replace NaNs with None and drop rows where all elements are NaN
        df = df.where(pd.notnull(df), None)
        df.dropna(how='all', inplace=True)

        # Save the DataFrame to the specified table in the SQLite database
        conn = sqlite3.connect(db_path)
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        return "Data loaded successfully"
    except Exception as e:
        error_message = f"Failed to load data from {excel_path} into {table_name}: {str(e)}"
        logging.error(error_message)
        error_messages.append(error_message)

    if error_messages:
        QMessageBox.critical(None, "Errors Encountered", "\n".join(error_messages))
        return "\n".join(error_messages)
    return "Data loaded successfully"

class App(QMainWindow):
    def __init__(self, db_path='data.db'):
        super().__init__()
        self.db_path = db_path
        self.settings_file = 'settings.json'
        self.current_theme = 'Dark'
        self.current_user = None
        self.connection_pool = None
        self.conn = sqlite3.connect(self.db_path)
        initialize_db(self.db_path)
        self.current_theme = self.get_last_logged_theme()  # Load the last logged theme
        self.data = {'blacklist': [], 'goldlist': [], 'prequal': [], 'mag_glass': []}
        self.make_map = {}  # Initialize make_map
        self.model_map = {}  # Initialize model_map
        self.setup_ui()
        self.prompt_user_pin()  # Prompt for PIN during initialization
        self.load_configurations()
        self.check_data_loaded()
        self.apply_saved_theme()  # Apply the last saved theme

    def get_last_logged_theme(self):
        conn = self.get_db_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT action FROM user_actions WHERE action LIKE 'Selected theme:%' ORDER BY timestamp DESC LIMIT 1")
            result = cursor.fetchone()
            if result:
                last_theme_action = result[0]
                return last_theme_action.split(":")[1].strip()
        except sqlite3.OperationalError as e:
            logging.error(f"Database error: {e}")
        return 'Dark'  # Default theme if none is found or if there's an error

    def log_action(self, user, action):
        conn = self.get_db_connection()
        try:
            cursor = conn.cursor()
            cst = pytz.timezone('America/Chicago')
            now = datetime.now(cst)
            timestamp = now.strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute('INSERT INTO user_actions (user, action, timestamp) VALUES (?, ?, ?)', (user, action, timestamp))
            conn.commit()
        except sqlite3.Error as e:
            logging.error(f"Failed to log action: {e}")
        finally:
            conn.close()

    def prompt_user_pin(self):
        while True:
            pin, ok = QInputDialog.getText(self, 'User Login', 'Enter your PIN:', QLineEdit.Password)
            if ok:
                conn = self.get_db_connection()  # Ensure the connection is opened for each attempt
                cursor = conn.cursor()
                cursor.execute('SELECT name FROM leader_log WHERE pin = ?', (pin,))
                result = cursor.fetchone()
                if result:
                    self.current_user = result[0]
                    self.log_action(self.current_user, "Logged in")
                    conn.close()
                    break
                elif pin == '9716':  # Check for the standard PIN
                    self.current_user = "Set Up"
                    self.log_action(self.current_user, "Logged in with standard PIN")
                    conn.close()
                    break
                else:
                    QMessageBox.warning(self, "Access Denied", "Incorrect User PIN.")
                    self.log_action("Unknown", "Failed User PIN attempt")
                    conn.close()
            else:
                QMessageBox.warning(self, "Access Denied", "Login Cancelled.")
                sys.exit()

    def check_data_loaded(self):
        if not self.data['prequal']:
            self.make_dropdown.setDisabled(True)
            self.model_dropdown.setDisabled(True)
            self.year_dropdown.setDisabled(True)
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

        self.add_toolbar_button("Manage Lists", self.open_admin, "admin_button")
        self.add_toolbar_button("Refresh Lists", self.refresh_lists, "refresh_button")
        self.add_pin_on_top_button()
        self.add_clear_filters_button()  # Add the "Clear Filters" button here

        # Create a vertical layout for the Leader button and theme dropdown
        button_theme_layout = QVBoxLayout()

        # Add the Leader button
        self.leader_button = QPushButton("ADMIN")
        self.leader_button.setObjectName("leader_button") 
        self.leader_button.setFixedHeight(30)  # Fixed height of 30
        self.leader_button.clicked.connect(self.leader_button_clicked)  # Connect to a function
        button_theme_layout.addWidget(self.leader_button)

        # Create theme dropdown
        self.theme_dropdown = QComboBox()
        themes = ["Dark", "Light", "Red", "Blue", "Green", "Yellow", "Pink", "Purple", "Teal", "Cyan", "Orange"]
        self.theme_dropdown.addItems(themes)
        self.theme_dropdown.currentIndexChanged.connect(self.apply_selected_theme)
        button_theme_layout.addWidget(self.theme_dropdown)

        # Add the vertical layout to the toolbar
        button_theme_container = QWidget()
        button_theme_container.setLayout(button_theme_layout)
        button_theme_container.setStyleSheet("background: #404040;")
        self.toolbar.addWidget(button_theme_container)

        # Add a QProgressBar to the status bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)  # Initially hidden
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #555;
                border-radius: 5px;
                text-align: center;
                background: #333;
            }
            QProgressBar::chunk {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #4f70e0,
                    stop: 1 #4f70e0);
                border-radius: 5px;
            }
        """)
        self.progress_bar.setTextVisible(False)  # Hide the text on the progress bar for a cleaner look
        self.status_bar.addPermanentWidget(self.progress_bar)

        # Create a horizontal layout for the dropdowns
        dropdown_layout = QHBoxLayout()

        self.make_dropdown = QComboBox()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        self.make_dropdown.currentIndexChanged.connect(self.update_model_dropdown)
        self.make_dropdown.focusInEvent = self.on_make_dropdown_focus  # Connect focus event
        dropdown_layout.addWidget(self.make_dropdown)

        self.model_dropdown = QComboBox()
        self.model_dropdown.addItem("Select Model")
        self.model_dropdown.currentIndexChanged.connect(self.handle_model_change)
        dropdown_layout.addWidget(self.model_dropdown)

        self.year_dropdown = QComboBox()
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.currentIndexChanged.connect(self.on_year_selected)
        dropdown_layout.addWidget(self.year_dropdown)

        main_layout.addLayout(dropdown_layout)  # Add the dropdown layout to the top of the main layout

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Enter DTC code or description")
        self.search_bar.returnPressed.connect(self.perform_search)
        self.year_dropdown.currentIndexChanged.connect(self.perform_search)
        main_layout.addWidget(self.search_bar)

        self.filter_dropdown = QComboBox()
        self.filter_dropdown.addItems(["Select List", "Blacklist", "Goldlist", "Prequals", "Mag Glass", "Gold and Black", "Black/Gold/Mag", "CarSys", "All"])
        self.filter_dropdown.currentIndexChanged.connect(self.filter_changed)  # Connect this to a new method
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

        # Create container widgets for panels
        self.left_panel_container = QWidget()
        self.left_panel_layout = QVBoxLayout(self.left_panel_container)
        
        # Add label and pop out button for left panel
        left_panel_header_layout = QHBoxLayout()
        self.left_panel_label = QLabel("Prequals")
        left_panel_header_layout.addWidget(self.left_panel_label)
        self.left_panel_pop_out_button = QPushButton("PO")
        self.left_panel_pop_out_button.setFixedSize(40, 30)
        self.left_panel_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Prequals", self.left_panel.toHtml()))
        left_panel_header_layout.addWidget(self.left_panel_pop_out_button)
        self.left_panel_layout.addLayout(left_panel_header_layout)
        
        self.left_panel = QTextBrowser()
        self.left_panel_layout.addWidget(self.left_panel)
        self.splitter.addWidget(self.left_panel_container)

        self.right_panel_container = QWidget()
        self.right_panel_layout = QVBoxLayout(self.right_panel_container)
        
        # Add label and pop out button for right panel
        right_panel_header_layout = QHBoxLayout()
        self.right_panel_label = QLabel("Gold and Black")
        right_panel_header_layout.addWidget(self.right_panel_label)
        self.right_panel_pop_out_button = QPushButton("PO")
        self.right_panel_pop_out_button.setFixedSize(40, 30)
        self.right_panel_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Gold and Black", self.right_panel.toHtml()))
        right_panel_header_layout.addWidget(self.right_panel_pop_out_button)
        self.right_panel_layout.addLayout(right_panel_header_layout)
        
        self.right_panel = QTextBrowser()
        self.right_panel_layout.addWidget(self.right_panel)
        self.splitter.addWidget(self.right_panel_container)

        self.mag_glass_container = QWidget()
        self.mag_glass_layout = QVBoxLayout(self.mag_glass_container)
        
        # Add label and pop out button for mag glass panel
        mag_glass_header_layout = QHBoxLayout()
        self.mag_glass_label = QLabel("Mag Glass")
        mag_glass_header_layout.addWidget(self.mag_glass_label)
        self.mag_glass_pop_out_button = QPushButton("PO")
        self.mag_glass_pop_out_button.setFixedSize(40, 30)
        self.mag_glass_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Mag Glass", self.mag_glass_panel.toHtml()))
        mag_glass_header_layout.addWidget(self.mag_glass_pop_out_button)
        self.mag_glass_layout.addLayout(mag_glass_header_layout)
        
        self.mag_glass_panel = QTextBrowser()
        self.mag_glass_layout.addWidget(self.mag_glass_panel)
        self.splitter.addWidget(self.mag_glass_container)
        self.mag_glass_container.setVisible(False)

        self.carsys_container = QWidget()
        self.carsys_layout = QVBoxLayout(self.carsys_container)

        # Add label and pop out button for CarSys panel
        carsys_header_layout = QHBoxLayout()
        self.carsys_label = QLabel("CarSys")
        carsys_header_layout.addWidget(self.carsys_label)
        self.carsys_pop_out_button = QPushButton("PO")
        self.carsys_pop_out_button.setFixedSize(40, 30)
        self.carsys_pop_out_button.clicked.connect(lambda: self.pop_out_panel("CarSys", self.carsys_panel.toHtml()))
        carsys_header_layout.addWidget(self.carsys_pop_out_button)
        self.carsys_layout.addLayout(carsys_header_layout)

        self.carsys_panel = QTextBrowser()
        self.carsys_layout.addWidget(self.carsys_panel)
        self.splitter.addWidget(self.carsys_container)
        self.carsys_container.setVisible(False)

        main_layout.addWidget(self.splitter)

        self.add_hide_show_buttons(main_layout)  # Add hide/show buttons

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

        self.apply_saved_theme()  # Apply the last saved theme
        self.populate_dropdowns()
        self.apply_dropdown_styles()  # Apply dropdown styles initially

    def pop_out_panel(self, title, content):
        pop_out_window = PopOutWindow(title, content, self.central_widget.styleSheet(), self)
        pop_out_window.show()

    def get_db_connection(self):
        return sqlite3.connect(self.db_path)

    def on_year_selected(self):
        selected_filter = self.filter_dropdown.currentText()
        # Check if a specific list that requires updating on year change is selected
        if selected_filter in ["Prequals", "Gold and Black", "Blacklist", "Goldlist", "Mag Glass", "All"]:
            self.perform_search()

    def filter_changed(self):
        selected_filter = self.filter_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()

        if selected_make == "All":
            self.model_dropdown.setCurrentIndex(0)
            self.year_dropdown.setCurrentIndex(0)
            selected_year = "Select Year"

        if selected_filter in ["All", "Prequals"] and selected_year == "Select Year":
            self.filter_dropdown.setCurrentIndex(0)  # Reset the filter dropdown to the default value
            self.left_panel.clear()  # Optionally clear the left panel
            return  # Early return to stop further processing
        
        if selected_filter == "CarSys":
            self.carsys_container.setVisible(True)
            self.left_panel_container.setVisible(False)
            self.right_panel_container.setVisible(False)
            self.display_carsys_data(selected_make)
        elif selected_filter == "Mag Glass":
            self.mag_glass_container.setVisible(True)
            self.left_panel_container.setVisible(False)
            self.right_panel_container.setVisible(False)
            self.display_mag_glass(selected_make)  # Pass the selected make to the display_mag_glass method
        elif selected_filter in ["All", "Black/Gold/Mag"]:
            self.mag_glass_container.setVisible(True)
            self.left_panel_container.setVisible(selected_filter == "All" and selected_year != "Select Year")
            self.right_panel_container.setVisible(True)
            self.display_mag_glass(selected_make)  # Pass the selected make to the display_mag_glass method
        else:
            self.mag_glass_container.setVisible(False)
            self.carsys_container.setVisible(False)
        self.perform_search()

    def add_hide_show_buttons(self, layout):
        button_layout = QHBoxLayout()
        self.left_hide_show_button = QPushButton("Hide Prequals")
        self.left_hide_show_button.setFixedSize(120, 30)
        self.left_hide_show_button.clicked.connect(self.toggle_left_panel)
        button_layout.addWidget(self.left_hide_show_button)

        self.right_hide_show_button = QPushButton("Hide Gold and Black")
        self.right_hide_show_button.setFixedSize(160, 30)
        self.right_hide_show_button.clicked.connect(self.toggle_right_panel)
        button_layout.addWidget(self.right_hide_show_button)

        self.mag_glass_hide_show_button = QPushButton("Hide Mag Glass")
        self.mag_glass_hide_show_button.setFixedSize(120, 30)
        self.mag_glass_hide_show_button.clicked.connect(self.toggle_mag_glass_panel)
        button_layout.addWidget(self.mag_glass_hide_show_button)

        self.carsys_hide_show_button = QPushButton("Hide CarSys")
        self.carsys_hide_show_button.setFixedSize(120, 30)
        self.carsys_hide_show_button.clicked.connect(self.toggle_carsys_panel)
        button_layout.addWidget(self.carsys_hide_show_button)

        layout.addLayout(button_layout)

    def toggle_left_panel(self):
        if self.left_panel_container.isVisible():
            self.left_panel_container.hide()
            self.left_hide_show_button.setText("Show Prequals")
        else:
            self.left_panel_container.show()
            self.left_hide_show_button.setText("Hide Prequals")

    def toggle_right_panel(self):
        if self.right_panel_container.isVisible():
            self.right_panel_container.hide()
            self.right_hide_show_button.setText("Show Gold and Black")
        else:
            self.right_panel_container.show()
            self.right_hide_show_button.setText("Hide Gold and Black")

    def toggle_mag_glass_panel(self):
        if self.mag_glass_container.isVisible():
            self.mag_glass_container.hide()
            self.mag_glass_hide_show_button.setText("Hide Mag Glass")
        else:
            self.mag_glass_container.show()
            self.mag_glass_hide_show_button.setText("Show Mag Glass")

    def toggle_carsys_panel(self):
        if self.carsys_container.isVisible():
            self.carsys_container.hide()
            self.carsys_hide_show_button.setText("Hide CarSys")
        else:
            self.carsys_container.show()
            self.carsys_hide_show_button.setText("Show CarSys")

    def set_refresh_button_style(self):
        self.refresh_button.setStyleSheet("""
            QPushButton#refresh_button {
                color: white;
                background-color: #333; /* Background color same as other buttons */
                border: 2px solid #555;
                border-radius: 10px;
                padding: 5px;
                height: 28px;
                margin: 2px;
            }
            QPushButton#refresh_button:hover {
                background-color: #555;
            }
            QPushButton#refresh_button:pressed {
                background-color: #666;
                border-style: inset;
            }
        """)

    def leader_button_clicked(self):
        admin_pin, ok = QInputDialog.getText(self, 'Admin Login', 'Enter Admin PIN:', QLineEdit.Password)
        if ok and admin_pin == '9716':  # Admin PIN check
            self.log_action(self.current_user, "Accessed Admin options with correct Admin PIN")
            self.open_admin_options()
        else:
            QMessageBox.warning(self, "Access Denied", "Incorrect Admin PIN.")
            self.log_action(self.current_user, "Failed Admin PIN attempt")

    def add_refresh_button(self):
        self.refresh_button = QPushButton("Refresh Lists")
        self.refresh_button.setObjectName("refresh_button")
        self.refresh_button.clicked.connect(self.refresh_lists)
        self.toolbar.addWidget(self.refresh_button)
        self.set_refresh_button_style()

    def apply_stylesheet(self, style):
        self.central_widget.setStyleSheet(style)
        self.toolbar.setStyleSheet(style)
        self.status_bar.setStyleSheet(style)
        self.set_refresh_button_style()  # Ensure the Refresh Lists button retains its style

    def open_admin_options(self):
        admin_options_dialog = AdminOptionsDialog(self)
        admin_options_dialog.exec_()

    def open_pin_management(self):
        pin_dialog = PinManagementDialog(self)
        pin_dialog.exec_()
        self.load_users_and_pins()  # Reload the users and pins after the dialog is closed

    def load_users_and_pins(self):
        self.valid_pins = {}
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT pin, name FROM leader_log')
        users = cursor.fetchall()
        for pin, name in users:
            self.valid_pins[pin] = name

    def log_leader_pin(self, pin):
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO leader_log (pin, name)
            VALUES (?, ?)
        ''', (pin, self.valid_pins[pin]))
        conn.commit()

    def enable_ui(self):
        self.central_widget.setEnabled(True)
        self.toolbar.setEnabled(True)
        self.status_bar.setEnabled(True)

    def disable_ui(self):
        self.central_widget.setEnabled(False)
        self.toolbar.setEnabled(False)
        self.status_bar.setEnabled(False)

    def authenticate_user(self):
        while not self.is_authenticated:
            pin, ok = QInputDialog.getText(self, 'Login', 'Enter your PIN:', QLineEdit.Password)
            if ok and pin:
                if pin == self.admin_pin:
                    self.is_authenticated = True
                    QMessageBox.information(self, "Access Granted", "Welcome, Admin!")
                    self.enable_ui()
                elif pin in self.valid_pins:
                    self.is_authenticated = True
                    user_name = self.valid_pins[pin]
                    QMessageBox.information(self, "Access Granted", f"Welcome, {user_name}!")
                    self.enable_ui()
                else:
                    QMessageBox.warning(self, "Access Denied", "Invalid PIN. Please try again.")

    def clear_filters(self):
        self.log_action(self.current_user, "Clicked Clear Filters button")
        self.make_dropdown.setCurrentIndex(0)
        self.model_dropdown.setCurrentIndex(0)
        self.year_dropdown.setCurrentIndex(0)
        self.filter_dropdown.setCurrentIndex(0)
        self.search_bar.clear()
        self.left_panel.clear()
        self.right_panel.clear()

    def add_clear_filters_button(self):
        button = QPushButton("Clear Filters")
        button.setObjectName("clear_filters_button")
        button.clicked.connect(self.clear_filters)
        button.setMinimumWidth(50)
        button.setMaximumWidth(50)
        self.toolbar.addWidget(button)

    def update_black_and_gold_display(self):
        selected_make = self.make_dropdown.currentText().strip()
        if selected_make and selected_make != "Select Make":
            self.display_gold_and_black(selected_make)
        else:
            self.right_panel.clear()

    def on_make_dropdown_focus(self, event):
        # Do nothing if "All" is selected
        if self.make_dropdown.currentText() == "All":
            return
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
        conn = self.get_db_connection()
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
        selected_model = self.model_dropdown.currentText().strip()
        logging.debug(f"Model selected: {selected_model}")

        if selected_model != "Select Model":
            self.update_year_dropdown()
        else:
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")

        self.perform_search()

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

    def apply_selected_theme(self):
        theme = self.theme_dropdown.currentText()
        self.current_theme = theme
        self.save_settings({'theme': self.current_theme})

        self.log_action(self.current_user, f"Selected theme: {theme}")

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
        elif theme == "RGB":
            self.apply_color_theme("#ff0000", "#0000ff", "#00ff00")  # Applying an RGB theme with red, green, and blue

        self.apply_button_styles()  # Apply the button styles after setting the theme
        self.apply_dropdown_styles()  # Apply dropdown styles

    def apply_dropdown_styles(self):
        theme = self.theme_dropdown.currentText()
        if theme == "Red":
            color_start = "#DC143C"
            color_end = "#A52A2A"
            text_color = "#fff"
        elif theme == "Blue":
            color_start = "#4169E1"
            color_end = "#1E3A5F"
            text_color = "#fff"
        elif theme == "Dark":
            color_start = "#2b2b2b"
            color_end = "#2b2b2b"
            text_color = "#ddd"
        elif theme == "Light":
            color_start = "#f0f0f0"
            color_end = "#d0d0d0"
            text_color = "black"
        elif theme == "Green":
            color_start = "#006400"
            color_end = "#003a00"
            text_color = "#fff"
        elif theme == "Yellow":
            color_start = "#ffd700"
            color_end = "#b29500"
            text_color = "black"
        elif theme == "Pink":
            color_start = "#ff69b4"
            color_end = "#b2477d"
            text_color = "black"
        elif theme == "Purple":
            color_start = "#800080"
            color_end = "#4b004b"
            text_color = "#fff"
        elif theme == "Teal":
            color_start = "#008080"
            color_end = "#004b4b"
            text_color = "#fff"
        elif theme == "Cyan":
            color_start = "#00ffff"
            color_end = "#00b2b2"
            text_color = "black"
        elif theme == "Orange":
            color_start = "#ff8c00"
            color_end = "#ff4500"
            text_color = "black"
        else:
            color_start = "#2b2b2b"
            color_end = "#2b2b2b"
            text_color = "#ddd"

        dropdown_style = f"""
        QComboBox {{
            border: 2px solid #555;
            border-radius: 10px;
            padding: 1px 18px 1px 3px;
            min-width: 6em;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
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
        """

        self.theme_dropdown.setStyleSheet(dropdown_style)

    def update_button_styles(self):
        theme = self.theme_dropdown.currentText()
        buttons = [self.findChild(QPushButton, "refresh_button"),
                self.findChild(QPushButton, "clear_filters_button"),
                self.leader_button]

        for button in buttons:
            if theme in ["Light", "Yellow", "Pink", "Cyan", "Orange"]:
                button.setStyleSheet("color: white;")
            else:
                button.setStyleSheet("")

        # Apply the common style for the Leader and Export buttons
        common_style = """
            QPushButton {
                background-color: #333;
                border: 2px solid #555;
                border-radius: 10px;
                padding: 5px;
                height: 28px;
                margin: 2px;
                color: white;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QPushButton:pressed {
                background-color: #666;
                border-style: inset;
            }
        """
        self.leader_button.setStyleSheet(common_style)
        self.findChild(QPushButton, "refresh_button").setStyleSheet(common_style)

    def update_theme_dropdown_style(self):
        theme = self.theme_dropdown.currentText()
        if theme == "Red":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #DC143C, stop:1 #A52A2A); color: white;")
        elif theme == "Blue":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #4169E1, stop:1 #1E3A5F); color: white;")
        elif theme == "Dark":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2b2b2b, stop:1 #2b2b2b); color: #ddd;")
        elif theme == "Light":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f0f0f0, stop:1 #d0d0d0); color: black;")
        elif theme == "Green":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #006400, stop:1 #003a00); color: white;")
        elif theme == "Yellow":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ffd700, stop:1 #b29500); color: black;")
        elif theme == "Pink":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ff69b4, stop:1 #b2477d); color: black;")
        elif theme == "Purple":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #800080, stop:1 #4b004b); color: white;")
        elif theme == "Teal":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #008080, stop:1 #004b4b); color: white;")
        elif theme == "Cyan":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #00ffff, stop:1 #00b2b2); color: black;")
        elif theme == "Orange":
            self.theme_dropdown.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ff8c00, stop:1 #ff4500); color: black;")

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
         QPushButton#refresh_button {{
            color: white;
        }}
        QPushButton#admin_button, QPushButton#export_button, QPushButton#pin_on_top_button, QPushButton#leader_button {{
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
        QPushButton#refresh_button {
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

    def apply_button_styles(self):
        button_style = """
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 30px;
            margin: 2px;
            color: white;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        """
        self.findChild(QPushButton, "clear_filters_button").setStyleSheet(button_style)
        self.findChild(QPushButton, "leader_button").setStyleSheet(button_style)

    def apply_stylesheet(self, style):
        self.central_widget.setStyleSheet(style)
        self.toolbar.setStyleSheet(style)
        self.status_bar.setStyleSheet(style)
        self.apply_button_styles()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r') as file:
                return json.load(file)
        return {'theme': 'Dark'}  # Default to Dark theme if no settings file exists

    def save_settings(self, settings):
        with open(self.settings_file, 'w') as file:
            json.dump(settings, file)

    def apply_saved_theme(self):
        saved_settings = self.load_settings()
        theme = saved_settings.get('theme', 'Dark')
        index = self.theme_dropdown.findText(theme)
        if index != -1:
            self.theme_dropdown.setCurrentIndex(index)
        self.apply_selected_theme()  # Apply the selected theme

    def closeEvent(self, event):
        self.log_action(self.current_user, f"Closed the program with theme: {self.current_theme}")
        if self.connection_pool:
            self.connection_pool.close()
        event.accept()

    def change_opacity(self, value):
        self.setWindowOpacity(value / 100.0)

    def clear_data(self, config_type=None):
        conn = self.get_db_connection()
        cursor = conn.cursor()
        try:
            if config_type:
                cursor.execute(f"DELETE FROM {config_type}")
                logging.info(f"Data cleared from {config_type}")
            else:
                cursor.execute("DROP TABLE IF EXISTS blacklist")
                cursor.execute("DROP TABLE IF EXISTS goldlist")
                cursor.execute("DROP TABLE IF EXISTS prequal")
                cursor.execute("DROP TABLE IF EXISTS mag_glass")
                initialize_db(self.db_path)
                logging.info("Database reset complete.")
            conn.commit()
        except sqlite3.Error as e:
            logging.error(f"Failed to clear data: {e}")
            QMessageBox.critical(self, "Error", "Failed to clear database.")
        self.load_configurations()

    def export_data(self):
        self.log_action(self.current_user, "Clicked Export button")
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", "", "All Files (*);;Text Files (*.txt);;CSV Files (*.csv);;JSON Files (*.json)", options=options)
        if fileName:
            if fileName.endswith('.csv'):
                self.export_to_csv(fileName)
            elif fileName.endswith('.json'):
                self.export_to_json(fileName)

            # Copy the folders with the new name
            conn = self.get_db_connection()  # Ensure the connection is open
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM paths")
                paths = cursor.fetchall()
                for config_type, folder_path in paths:
                    if os.path.exists(folder_path):
                        new_folder_name = os.path.join(os.path.dirname(fileName), f"{os.path.basename(fileName)}_{config_type}_folder")
                        if os.path.exists(new_folder_name):
                            shutil.rmtree(new_folder_name)
                        shutil.copytree(folder_path, new_folder_name)
            except sqlite3.Error as e:
                logging.error(f"Failed to retrieve paths from database: {e}")
            finally:
                conn.close()

    def export_to_csv(self, path):
        conn = self.get_db_connection()
        try:
            user_actions_df = pd.read_sql_query("SELECT * FROM user_actions", conn)
            user_actions_df.to_csv(f"{path}_user_actions.csv", index=False)
            
            QMessageBox.information(self, "Export Successful", "Data exported successfully to CSV.")
        except Exception as e:
            logging.error(f"Failed to export data to CSV: {e}")
            QMessageBox.critical(self, "Export Error", "Failed to export data to CSV.")
        finally:
            conn.close()

    def export_to_json(self, path):
        conn = self.get_db_connection()
        try:
            user_actions_df = pd.read_sql_query("SELECT * FROM user_actions", conn)
            user_actions_df.to_json(f"{path}_user_actions.json", orient='records', lines=True)
            
            QMessageBox.information(self, "Export Successful", "Data exported successfully to JSON.")
        except Exception as e:
            logging.error(f"Failed to export data to JSON: {e}")
            QMessageBox.critical(self, "Export Error", "Failed to export data to JSON.")
        finally:
            conn.close()

    def update_model_dropdown(self):
        selected_make = self.make_dropdown.currentText().strip()
        selected_filter = self.filter_dropdown.currentText()

        if selected_make == "All" and selected_filter in ["All", "Prequals"]:
            self.left_panel.clear()  # Optionally clear the prequals display if needed
            self.model_dropdown.setCurrentIndex(0)
            self.year_dropdown.setCurrentIndex(0)
            logging.debug("Make is 'All' and Filter is 'All' or 'Prequals'. Exiting early.")
            return  # Early return to stop further processing

        if selected_make != "Select Make":
            # Filter prequal data to only include entries with valid prequalification information
            valid_prequal_data = [item for item in self.data['prequal'] if self.has_valid_prequal(item)]

            # Ensure models are treated as strings
            models = [str(item['Model']) for item in valid_prequal_data if item['Make'] == selected_make]
            unique_models = sorted(set(models))

            self.model_dropdown.clear()
            self.model_dropdown.addItem("Select Model")
            self.model_dropdown.addItems(unique_models)
        else:
            self.model_dropdown.clear()
            self.model_dropdown.addItem("Select Model")
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")

    def handle_model_change(self, index):
        selected_model = self.model_dropdown.currentText().strip()
        if selected_model != "Select Model":
            self.update_year_dropdown()
        else:
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")
            self.perform_search()

    def open_admin(self):
        self.log_action(self.current_user, "Clicked Manage Lists button")
        password, ok = QInputDialog.getText(self, 'Login', 'Enter password:', QLineEdit.Password)
        if ok and password == "ADAS":
            choice, ok = QInputDialog.getItem(self, "Admin Actions", "Select action:", ["Update Paths", "Clear Data"], 0, False)
            if ok and choice == "Clear Data":
                # Added 'mag_glass' to the list of options
                config_type, ok = QInputDialog.getItem(self, "Clear Data", "Select configuration to clear or 'All' to reset database:", ["blacklist", "goldlist", "prequal", "mag_glass", "All", "CarSys"], 0, False)
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
            # Filter prequal data to only include entries with valid prequalification information
            valid_prequal_data = [item for item in self.data['prequal'] if self.has_valid_prequal(item)]

            # Ensure the year values are whole numbers (integers) and remove any non-integer values
            years = []
            for item in valid_prequal_data:
                try:
                    year = int(float(item['Year']))  # Convert the year to an integer
                    years.append(year)
                except (ValueError, TypeError):
                    continue

            # Remove duplicates and sort the years
            unique_years = sorted(set(years))

            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")
            self.year_dropdown.addItems([str(year) for year in unique_years])
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

    def manage_paths_thread(self):
        threading.Thread(target=self.manage_paths).start()

    def manage_paths(self):
        config_type, ok = QInputDialog.getItem(self, "Select Config Type", "Choose the configuration to update:", ["blacklist", "goldlist", "prequal", "mag_glass", "CarSys"], 0, False)
        if ok:
            folder_path = QFileDialog.getExistingDirectory(self, "Select Directory")
            if folder_path:
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    QMessageBox.warning(self, "Load Error", "No valid Excel files found in the directory.")
                    return

                data_loaded = False
                self.progress_bar.setVisible(True)
                self.progress_bar.setMaximum(len(files))
                self.progress_bar.setValue(0)

                for i, (filename, filepath) in enumerate(files.items()):
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=1)
                            if result != "Failed to load data":
                                data_loaded = True
                        elif config_type == 'mag_glass':
                            result = load_mag_glass_data_from_5th_sheet(filepath, config_type)
                            if result != "Failed to load data":
                                data_loaded = True
                        elif config_type == 'CarSys':
                            result = load_carsys_data_to_db(filepath, table_name=config_type, db_path=self.db_path)
                            if result != "Data loaded successfully":
                                data_loaded = False  # Handle errors more specifically here
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

                    self.progress_bar.setValue(i + 1)  # Update progress bar

                self.progress_bar.setVisible(False)

                if data_loaded:
                    self.load_configurations()
                    save_path_to_db(config_type, folder_path, self.db_path)  # Save the selected path
                    self.populate_dropdowns()  # Repopulate dropdowns
                    self.check_data_loaded()  # Ensure the dropdowns are enabled if data is loaded
                    QMessageBox.information(self, "Update Successful", "Data loaded successfully.")
                    self.status_bar.showMessage(f"Data updated from: {folder_path}")

    def refresh_lists_thread(self):
        thread = threading.Thread(target=self.refresh_lists)
        thread.start()

    def refresh_lists(self):
        self.log_action(self.current_user, "Clicked Refresh Lists button")
        
        # Include 'mag_glass' and 'CarSys' in the list
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'CarSys']:
            folder_path = load_path_from_db(config_type, self.db_path)
            
            if folder_path:
                # First, clear existing data for the config_type
                self.clear_data(config_type)
                logging.info(f"Cleared existing data for {config_type}")

                # Then get the valid files from the folder path
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    QMessageBox.warning(self, "Load Error", f"No valid Excel files found in the directory for {config_type}.")
                    continue

                data_loaded = False
                self.progress_bar.setVisible(True)
                self.progress_bar.setMaximum(len(files))
                self.progress_bar.setValue(0)

                # Load data from each file
                for i, (filename, filepath) in enumerate(files.items()):
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=1)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'mag_glass':
                            result = load_mag_glass_data_from_5th_sheet(filepath, config_type, db_path=self.db_path)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'CarSys':
                            result = load_carsys_data_to_db(filepath, table_name=config_type, db_path=self.db_path)
                            if result == "Data loaded successfully":
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

                    self.progress_bar.setValue(i + 1)  # Update progress bar

                self.progress_bar.setVisible(False)

                if data_loaded:
                    self.load_configurations()
                    self.populate_dropdowns()  # Repopulate dropdowns
                    self.check_data_loaded()  # Ensure the dropdowns are enabled if data is loaded
                    QMessageBox.information(self, "Update Successful", f"Data refreshed successfully from: {folder_path}")
                    self.status_bar.showMessage(f"Data refreshed from: {folder_path}")
            else:
                logging.warning(f"No saved path found for {config_type}")

    def load_configurations(self):
        logging.debug("Loading configurations...")
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass']:
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

        # Filter prequal data to only include entries with valid prequalification information
        valid_prequal_data = [item for item in self.data['prequal'] if self.has_valid_prequal(item)]

        makes = sorted(set(item['Make'].strip() for item in valid_prequal_data if item['Make'].strip() and item['Make'].strip().lower() != 'unknown'))

        if '' in makes:
            makes.remove('')

        self.make_dropdown.addItems(makes)
        logging.debug(f"Makes populated: {makes}")

        # Reset model and year dropdowns
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")
        self.year_dropdown.clear()
        self.year_dropdown.addItem("Select Year")

    def has_valid_prequal(self, item):
        # Add your logic here to determine if the item has valid prequalification information
        return item.get('Calibration Pre-Requisites') not in [None, 'N/A']

    def perform_search_thread(self):
        threading.Thread(target=self.perform_search).start()

    def perform_search(self):
        if not self.current_user:
            return

        dtc_code = self.search_bar.text().strip().upper()
        selected_filter = self.filter_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()

        # Reset model and year if make is set to "All"
        if selected_make == "All":
            self.model_dropdown.setCurrentIndex(0)
            self.year_dropdown.setCurrentIndex(0)
            selected_model = "Select Model"
            selected_year = "Select Year"

        # Clear all panels if no specific make/model/year is selected
        if selected_make == "Select Make" and selected_model == "Select Model" and selected_year == "Select Year":
            self.clear_display_panels()
            return

        # If 'All' or 'Prequals' is selected in the filter dropdown without a year selected, show warning and clear the left panel
        if selected_filter in ["All", "Prequals"] and selected_year == "Select Year":
            self.left_panel.clear()  # Optionally clear the prequals display if needed
            return  # Early return to stop further processing

        # Handle the CarSys search
        if selected_filter == "CarSys":
            self.search_carsys_dtc(dtc_code)
            self.display_carsys_data(selected_make) 
            return  # Stop further processing for CarSys

        # Log the search
        self.log_action(self.current_user, f"Performed search with DTC: {dtc_code}, Filter: {selected_filter}, Make: {selected_make}, Model: {selected_model}, Year: {selected_year}")

        # Handle the "Select List" as a do-nothing case
        if selected_filter == "Select List":
            return
        
        # Handle prequal search based on make, model, and year
        if selected_filter == "Prequals" and selected_make != "Select Make" and selected_model != "Select Model" and selected_year != "Select Year":
            try:
                selected_year = int(selected_year)  # Convert selected year to integer
            except ValueError:
                self.left_panel.clear()  # Clear the display if year is invalid
                return

            # Filter prequal data based on the selected make, model, and year
            filtered_prequals = [
                item for item in self.data['prequal']
                if item['Make'] == selected_make and str(item['Model']) == str(selected_model) and int(float(item['Year'])) == selected_year
            ]

            # Display the filtered prequal data
            if filtered_prequals:
                self.display_results(filtered_prequals, context='prequal')
            else:
                self.left_panel.setPlainText("No prequal data found for the selected criteria.")
            return  # Stop further processing for Prequals

        # Update CarSys panel if the search bar is not empty
        if dtc_code:
            self.search_carsys_dtc(dtc_code)

        # Show or hide panels based on the selected filter and conditions
        self.update_panel_visibility(selected_filter, selected_make)

        # Check conditions and update displays
        self.update_displays_based_on_filter(selected_filter, dtc_code, selected_make, selected_model, selected_year)

    def clear_display_panels(self):
        self.left_panel.clear()
        self.right_panel.clear()
        self.mag_glass_panel.clear()

    def update_panel_visibility(self, selected_filter, selected_make):
        self.splitter.widget(0).setVisible(selected_filter in ["Prequals", "All"] and selected_make != "All")
        self.splitter.widget(1).setVisible(selected_filter in ["Gold and Black", "Blacklist", "Goldlist", "Black/Gold/Mag", "All"])
        self.splitter.widget(2).setVisible(selected_filter in ["Mag Glass", "Black/Gold/Mag", "All"])

    def update_displays_based_on_filter(self, selected_filter, dtc_code, selected_make, selected_model, selected_year):
        if selected_filter == "All":
            # Display all categories including prequals, gold and black lists, and mag glass
            if selected_year != "Select Year":
                self.handle_prequal_search(selected_make, selected_model, selected_year)
            if dtc_code:
                self.search_dtc_codes(dtc_code, "Gold and Black", selected_make)  # This will depend on DTC input
            else:
                self.display_gold_and_black(selected_make)  # Display gold and black lists based on selected make
            self.display_mag_glass(selected_make)  # Display mag glass data

        elif selected_filter == "Prequals":
            self.handle_prequal_search(selected_make, selected_model, selected_year)

        elif selected_filter == "Black/Gold/Mag":
            self.display_gold_and_black(selected_make)  # Display gold and black lists based on selected make
            self.display_mag_glass(selected_make)  # Display mag glass data

        elif selected_filter in ["Gold and Black", "Blacklist", "Goldlist"]:
            if not dtc_code and selected_make == "All":
                self.right_panel.setPlainText("Please enter a DTC code or description to search.")
            else:
                self.search_dtc_codes(dtc_code, selected_filter, selected_make)

        elif selected_filter == "Mag Glass":
            self.display_mag_glass(selected_make)

    def handle_prequal_search(self, selected_make, selected_model, selected_year):
        # Dictionary to hold unique System Acronyms
        unique_results = {}

        # Convert selected_model to a string
        selected_model_str = str(selected_model)

        # Filtering data based on selections
        filtered_results = [item for item in self.data['prequal']
                            if (selected_make == "All" or item['Make'] == selected_make) and
                            (selected_model == "Select Model" or str(item['Model']) == selected_model_str) and
                            (selected_year == "Select Year" or str(item['Year']) == selected_year)]

        # Populating the dictionary with unique entries based on System Acronym
        for item in filtered_results:
            system_acronym = item.get('Protech Generic System Name.1', 'N/A')
            if system_acronym not in unique_results:
                unique_results[system_acronym] = item  # Store the entire item for display

        # Now, pass the unique items to be displayed
        self.display_results(list(unique_results.values()), context='prequal')

    def display_carsys_data(self, selected_make):
        try:
            conn = sqlite3.connect(self.db_path)
            
            # If no make is selected, display all results
            if selected_make == "Select Make" or selected_make == "All":
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                """
                df = pd.read_sql_query(query, conn)
            else:
                # If a make is selected, filter the results by the selected make
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                WHERE carMake = ?
                """
                df = pd.read_sql_query(query, conn, params=(selected_make,))
            
            # Replace NaN values with empty strings
            df.fillna("", inplace=True)
            
            # Display the results in the CarSys panel
            if df.empty:
                self.carsys_panel.setPlainText("No results found.")
            else:
                self.carsys_panel.setHtml(df.to_html(index=False, escape=False))
        
        except Exception as e:
            logging.error(f"Failed to display CarSys data: {e}")
            self.carsys_panel.setPlainText("An error occurred while fetching the CarSys data.")
        
        finally:
            conn.close()

    def display_mag_glass(self, selected_make):
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Prepare the query based on the selected make
            if selected_make == "All":
                query = """
                SELECT "Generic System Name", "ADAS Module Name", "Car Make", "Manufacturer", "AUTEL or BOSCH"
                FROM mag_glass
                """
                df = pd.read_sql_query(query, conn)
            else:
                query = """
                SELECT "Generic System Name", "ADAS Module Name", "Car Make", "Manufacturer", "AUTEL or BOSCH"
                FROM mag_glass
                WHERE "Car Make" = ?
                """
                df = pd.read_sql_query(query, conn, params=(selected_make,))
            
            # Display the data in the Mag Glass panel
            if not df.empty:
                self.mag_glass_panel.setHtml(df.to_html(index=False, escape=False))
            else:
                self.mag_glass_panel.setPlainText(f"No results found for Make: {selected_make}")
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.mag_glass_panel.setPlainText("An error occurred while fetching the data.")
        finally:
            if conn:
                conn.close()

    def search_carsys_dtc(self, dtc_code):
        conn = self.get_db_connection()

        # Modify query to filter by DTCsys and optionally by carMake
        if dtc_code:
            if self.make_dropdown.currentText() not in ["Select Make", "All"]:
                query = f"""
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                WHERE dtcSys LIKE '%{dtc_code}%' AND carMake = ?
                """
                params = (self.make_dropdown.currentText(),)
            else:
                query = f"""
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                WHERE dtcSys LIKE '%{dtc_code}%'
                """
                params = None
        else:
            # If no DTC code is entered, show all results
            query = """
            SELECT genericSystemName, dtcSys, carMake, comments
            FROM carsys
            """
            params = None

        try:
            if params:
                df = pd.read_sql_query(query, conn, params=params)
            else:
                df = pd.read_sql_query(query, conn)
            
            # Replace NaN values with empty strings
            df.fillna("", inplace=True)

            if df.empty:
                self.carsys_panel.setPlainText("No results found for the given criteria.")
            else:
                self.carsys_panel.setHtml(df.to_html(index=False, escape=False))
        
        except Exception as e:
            logging.error(f"Failed to execute CarSys search query: {query}\nError: {e}")
            self.carsys_panel.setPlainText("An error occurred while fetching the CarSys data.")
        
        finally:
            conn.close()

    def search_mag_glass(self, selected_make):
        conn = self.get_db_connection()
        cursor = conn.cursor()

        query = f"""
        SELECT "Generic System Name", "ADAS Module Name", "Car Make", "Manufacturer", "AUTEL or BOSCH"
        FROM mag_glass
        """
        
        if selected_make != "All":
            query += f" WHERE \"Car Make\" = '{selected_make}'"
        
        try:
            df = pd.read_sql_query(query, conn)
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.mag_glass_panel.setPlainText("An error occurred while fetching the data.")
            return

        if not df.empty:
            self.mag_glass_panel.setHtml(df.to_html(index=False, escape=False))
        else:
            self.mag_glass_panel.setPlainText("No Mag Glass results found.")

    def search_dtc_codes(self, dtc_code, filter_type, selected_make):
        conn = self.get_db_connection()
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

        if not df.empty:
            self.right_panel.setHtml(df.to_html(index=False, escape=False))
        else:
            self.right_panel.setPlainText("No DTC code results found.")

    def display_gold_and_black(self, selected_make):
        try:
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
            
            df = pd.read_sql_query(query, conn)
            if not df.empty:
                self.right_panel.setHtml(df.to_html(index=False, escape=False))
            else:
                self.right_panel.setPlainText("No DTC code results found.")
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.right_panel.setPlainText("An error occurred while fetching the data.")
        finally:
            if conn:
                conn.close()

        if not df.empty:
            self.right_panel.setHtml(df.to_html(index=False, escape=False))
        else:
            self.right_panel.setPlainText("No DTC code results found.")

    def display_results(self, results, context='prequal'):
        display_text = ""
        for result in results:
            # Check if any relevant field is "N/A" or None
            if any(
                result.get(key) in [None, 'N/A'] 
                for key in ['Make', 'Model', 'Year', 'Calibration Type', 'Protech Generic System Name.1', 'Parts Code Table Value', 'Calibration Pre-Requisites']
            ):
                continue  # Skip this result

            calibration_type = result.get('Calibration Type', 'N/A')

            # Handle link safely
            link = result.get('Service Information Hyperlink', '#')
            if isinstance(link, float) and pd.isna(link):
                link = '#'
            elif isinstance(link, str) and not link.startswith(('http://', 'https://')):
                link = 'http://' + link

            system_acronym = result.get('Protech Generic System Name.1', 'N/A')
            parts_code = result.get('Parts Code Table Value', 'N/A')
            calibration_prerequisites = result.get('Calibration Pre-Requisites', 'N/A')

            # Ensure the year is displayed as a whole number
            year = result.get('Year', 'N/A')
            if isinstance(year, float):
                year = str(int(year))  # Convert float year to an integer string

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
                <b>Year:</b> {year}<br>  <!-- Display year as a whole number -->
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
