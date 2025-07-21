import random
import sys
import os
import shutil
import sqlite3
import time
import threading
import pytz
import logging
import json
import re
from datetime import datetime
import pandas as pd
from PyQt5.QtCore import Qt, QUrl, QTimer, QPropertyAnimation, QEasingCurve, QRect
from PyQt5.QtGui import QDesktopServices, QIcon, QKeySequence, QFont, QPalette, QColor, QPixmap, QPainter, QLinearGradient
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLineEdit, QComboBox, QTextBrowser, QListWidget, QLabel, 
    QPushButton, QToolBar, QStatusBar, QMessageBox, QDialog, 
    QFileDialog, QSplitter, QFrame, QProgressBar, QInputDialog,
    QShortcut, QSizePolicy, QSlider, QFormLayout, QTabWidget,
    QScrollArea, QGridLayout, QGroupBox, QCheckBox, QRadioButton,
    QButtonGroup, QStackedWidget, QGraphicsDropShadowEffect,
    QGraphicsOpacityEffect, QStyleFactory, QStyle
)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class ModernCard(QWidget):
    """A modern card widget with shadow and rounded corners"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ModernCard")
        self.setStyleSheet("""
            QWidget#ModernCard {
                background-color: #ffffff;
                border-radius: 12px;
                border: 1px solid #e0e0e0;
            }
        """)
        
        # Add shadow effect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)

class ModernButton(QPushButton):
    """A modern button with hover effects and rounded corners"""
    def __init__(self, text="", parent=None, style="primary"):
        super().__init__(text, parent)
        self.style_type = style
        self.setup_style()
        
    def setup_style(self):
        if self.style_type == "primary":
            self.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                        stop:0 #667eea, stop:1 #764ba2);
                    border: none;
                    border-radius: 8px;
                    color: white;
                    font-weight: 600;
                    padding: 12px 24px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                        stop:0 #5a6fd8, stop:1 #6a4190);
                }
                QPushButton:pressed {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                        stop:0 #4f5fc6, stop:1 #5f377e);
                }
            """)
        elif self.style_type == "secondary":
            self.setStyleSheet("""
                QPushButton {
                    background: #f8f9fa;
                    border: 2px solid #e9ecef;
                    border-radius: 8px;
                    color: #495057;
                    font-weight: 600;
                    padding: 10px 20px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background: #e9ecef;
                    border-color: #dee2e6;
                }
                QPushButton:pressed {
                    background: #dee2e6;
                }
            """)
        elif self.style_type == "danger":
            self.setStyleSheet("""
                QPushButton {
                    background: #dc3545;
                    border: none;
                    border-radius: 8px;
                    color: white;
                    font-weight: 600;
                    padding: 10px 20px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background: #c82333;
                }
                QPushButton:pressed {
                    background: #bd2130;
                }
            """)

class ModernComboBox(QComboBox):
    """A modern combobox with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QComboBox {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                padding: 8px 12px;
                background: white;
                font-size: 14px;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #667eea;
            }
            QComboBox:focus {
                border-color: #667eea;
                outline: none;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #6c757d;
                margin-right: 8px;
            }
            QComboBox QAbstractItemView {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                background: white;
                selection-background-color: #667eea;
                selection-color: white;
                /* color: #6c757d; */
            }
        """)

class ModernLineEdit(QLineEdit):
    """A modern line edit with custom styling"""
    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setStyleSheet("""
            QLineEdit {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                padding: 10px 12px;
                background: white;
                font-size: 14px;
                min-height: 20px;
            }
            QLineEdit:hover {
                border-color: #667eea;
            }
            QLineEdit:focus {
                border-color: #667eea;
                outline: none;
            }
        """)

class ModernTextBrowser(QTextBrowser):
    """A modern text browser with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QTextBrowser {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                background: white;
                font-size: 13px;
                line-height: 1.5;
            }
            QTextBrowser:focus {
                border-color: #667eea;
            }
        """)

class ModernProgressBar(QProgressBar):
    """A modern progress bar with gradient styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QProgressBar {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                text-align: center;
                background: #f8f9fa;
                font-weight: 600;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #667eea, stop:1 #764ba2);
                border-radius: 6px;
            }
        """)

class ModernSlider(QSlider):
    """A modern slider with custom styling"""
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setStyleSheet("""
            QSlider::groove:horizontal {
                border: 1px solid #e9ecef;
                height: 8px;
                background: #f8f9fa;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #667eea, stop:1 #764ba2);
                border: 2px solid #667eea;
                width: 18px;
                margin: -5px 0;
                border-radius: 9px;
            }
            QSlider::handle:horizontal:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5a6fd8, stop:1 #6a4190);
            }
        """)

class ModernTabWidget(QTabWidget):
    """A modern tab widget with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background: transparent;
            }
            QTabBar::tab {
                background: #fff;
                color: #20567C;
                border: none;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                padding: 12px 32px;
                margin-right: 8px;
                font-size: 16px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 120px;
                min-height: 36px;
                box-shadow: 0 2px 8px rgba(21,52,74,0.08);
            }
            QTabBar::tab:selected {
                background: #20567C;
                color: #fff;
                font-weight: 700;
                box-shadow: 0 4px 16px rgba(21,52,74,0.12);
            }
            QTabBar::tab:hover:!selected {
                background: #e6eef5;
                color: #20567C;
            }
        """)

class ModernSplitter(QSplitter):
    """A modern splitter with custom styling"""
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setStyleSheet("""
            QSplitter::handle {
                background: #e9ecef;
                border-radius: 2px;
            }
            QSplitter::handle:hover {
                background: #667eea;
            }
        """)

class ModernStatusBar(QStatusBar):
    """A modern status bar with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QStatusBar {
                background: #f8f9fa;
                border-top: 1px solid #e9ecef;
                color: #6c757d;
                font-size: 12px;
            }
        """)

class ModernToolBar(QToolBar):
    """A modern toolbar with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QToolBar {
                background: #ffffff;
                border-bottom: 1px solid #e9ecef;
                spacing: 8px;
                padding: 8px;
            }
            QToolBar QToolButton, QPushButton {
                background: #19507a;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 18px;
                font-size: 15px;
                font-weight: 600;
                margin-right: 8px;
            }
            QToolBar QToolButton:hover, QPushButton:hover {
                background: #1e5d8a;
            }
            QToolBar QToolButton:pressed, QPushButton:pressed {
                background: #102535;
            }
            QLabel {
                color: #fff;
                font-size: 15px;
                font-weight: 500;
            }
            QComboBox {
                background: #fff;
                color: #20567C;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 14px;
                min-width: 110px;
            }
        """)

class ModernMainWindow(QMainWindow):
    """Base class for modern main window styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QMainWindow {
                background: #f8f9fa;
            }
        """)

class ModernDialog(QDialog):
    """Base class for modern dialog styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QDialog {
                background: #fff;
                border-radius: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #20567C;
                font-size: 15px;
            }
            QPushButton {
                background: #19507a;
                color: #fff;
                border-radius: 8px;
                padding: 8px 18px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1e5d8a;
            }
            QPushButton:pressed {
                background: #102535;
            }
            QLineEdit, QComboBox {
                background: #f8fafc;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 14px;
            }
        """)

# Copy all the utility functions and database functions from the original
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

def get_db_connection(db_path='data.db'):
    """Get a database connection"""
    return sqlite3.connect(db_path)

def handle_error(func, path, exc_info):
    """Handle file operation errors"""
    logging.error(f"Error in {func}: {path} - {exc_info}")

def save_path_to_db(config_type, folder_path, db_path='data.db'):
    import os, shutil, sqlite3, logging
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

        # Handle existing backup - remove it if it exists
        if os.path.exists(backup_path):
            logging.info(f"Backup path already exists: {backup_path}. Removing to overwrite.")
            try:
                shutil.rmtree(backup_path)
            except (PermissionError, OSError) as e:
                logging.error(f"Error removing existing backup: {e}. Operations may be incomplete.")

        # Create the new backup
        logging.info(f"Creating backup from {folder_path} to {backup_path}.")
        try:
            shutil.copytree(folder_path, backup_path, symlinks=False)
            logging.info(f"Backup successfully created at {backup_path}")
        except shutil.Error as e:
            logging.warning(f"Some errors during backup: {e}")
        except OSError as e:
            logging.error(f"OS error during copy operation: {e}")
    except sqlite3.Error as e:
        logging.error(f"Failed to save path to database: {e}")
    finally:
        conn.close()

def load_path_from_db(config_type, db_path='data.db'):
    import sqlite3, logging
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
    import sqlite3, json
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    if config_type == 'mag_glass':
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
    import sqlite3, json, logging, pandas as pd
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

def normalize_col(col):
    # Lowercase, remove non-alphanumeric, strip spaces
    return re.sub(r'[^a-z0-9]', '', col.lower())

def load_excel_data_to_db(excel_path, table_name, db_path='data.db', sheet_index=0, parent=None):
    import pandas as pd, sqlite3, logging, json
    from PyQt5.QtWidgets import QMessageBox
    error_messages = []
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_index)
        logging.debug(f"Data loaded from sheet index '{sheet_index}': {df.head()}")
        # Normalize columns
        original_cols = list(df.columns)
        norm_map = {col: normalize_col(col) for col in original_cols}
        df.columns = [normalize_col(col) for col in df.columns]
        # Define expected columns for each table (normalized)
        if table_name == 'mag_glass':
            expected_map = {
                'genericSystemName': 'generic system name',
                'adasModuleName': 'adas module name',
                'carMake': 'car make',
                'manufacturer': 'manufacturer',
                'autelOrBosch': 'autel or bosch'
            }
            expected_norm = {k: normalize_col(v) for k, v in expected_map.items()}
            rename_dict = {v: k for k, v in expected_norm.items()}
            df = df.rename(columns=rename_dict)
            df = df[list(expected_map.keys())]
            df = df.where(pd.notnull(df), None)
            df.dropna(how='all', inplace=True)
            conn = sqlite3.connect(db_path)
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            conn.close()
            return "Data loaded successfully"
        elif table_name in ['blacklist', 'goldlist']:
            expected_map = {
                'genericSystemName': 'generic system name',
                'dtcCode': 'dtc code',
                'dtcDescription': 'dtc description',
                'dtcSys': 'dtc sys',
                'carMake': 'car make',
                'comments': 'comments'
            }
            expected_norm = {k: normalize_col(v) for k, v in expected_map.items()}
            rename_dict = {v: k for k, v in expected_norm.items()}
            df = df.rename(columns=rename_dict)
            df = df[list(expected_map.keys())]
            df = df.where(pd.notnull(df), None)
            df.dropna(how='all', inplace=True)
            df['dtcCode'] = df['dtcCode'].astype(str)
            conn = sqlite3.connect(db_path)
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            conn.close()
            return "Data loaded successfully"
        elif table_name == 'prequal':
            # For prequal, just store as JSON
            df = df.where(pd.notnull(df), None)
            df.dropna(how='all', inplace=True)
            data = df.to_dict(orient='records')
            folder_path = os.path.dirname(excel_path)
            update_configuration('prequal', folder_path, data, db_path)
            return "Data loaded successfully"
        elif table_name == 'carsys':
            expected_map = {
                'genericSystemName': 'generic system name',
                'dtcSys': 'dtcsys',
                'carMake': 'car make',
                'comments': 'comments'
            }
            expected_norm = {k: normalize_col(v) for k, v in expected_map.items()}
            rename_dict = {v: k for k, v in expected_norm.items()}
            df = df.rename(columns=rename_dict)
            df = df[list(expected_map.keys())]
            df = df.where(pd.notnull(df), None)
            df.dropna(how='all', inplace=True)
            conn = sqlite3.connect(db_path)
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            conn.close()
            return "Data loaded successfully"
        elif table_name == 'CarSys':
            # Handle CarSys specifically - use first sheet and specific columns
            df = pd.read_excel(excel_path, sheet_name=0, usecols="A:D", header=0)
            df.columns = df.columns.str.strip()
            expected_columns = ['Generic System Name', 'DTCsys', 'CarMake', 'Comments']
            if not all(col in df.columns for col in expected_columns):
                missing_cols = [col for col in expected_columns if col not in df.columns]
                raise ValueError(f"Missing expected columns in the Excel file: {', '.join(missing_cols)}")
            df = df.rename(columns={
                'Generic System Name': 'genericSystemName',
                'DTCsys': 'dtcSys',
                'CarMake': 'carMake',
                'Comments': 'comments'
            })
            df = df[['genericSystemName', 'dtcSys', 'carMake', 'comments']]
            df['genericSystemName'] = df['genericSystemName'].astype(str)
            df['dtcSys'] = df['dtcSys'].astype(str)
            df['carMake'] = df['carMake'].astype(str)
            df['comments'] = df['comments'].astype(str)
            df = df.where(pd.notnull(df), None)
            df.dropna(how='all', inplace=True)
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute('DROP TABLE IF EXISTS carsys')
            conn.commit()
            df.to_sql('carsys', conn, if_exists='replace', index=False)
            conn.close()
            return "Data loaded successfully"
        else:
            error_message = f"Unknown table name: {table_name}"
            logging.error(error_message)
            error_messages.append(error_message)
    except Exception as e:
        error_message = f"Failed to load data from {excel_path} into {table_name}: {str(e)}"
        logging.error(error_message)
        error_messages.append(error_message)
    if error_messages:
        QMessageBox.critical(parent, "Errors Encountered", "\n".join(error_messages))
        return "\n".join(error_messages)
    return "Data loaded successfully"

def load_carsys_data_to_db(excel_path, table_name='carsys', db_path='data.db', parent=None):
    import pandas as pd, sqlite3, logging
    from PyQt5.QtWidgets import QMessageBox
    error_messages = []
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
        conn.commit()
        df = pd.read_excel(excel_path, sheet_name=0, usecols="A:D", header=0)
        logging.debug(f"Data loaded from sheet index '0': {df.head()}")
        df.columns = df.columns.str.strip()
        expected_columns = ['Generic System Name', 'DTCsys', 'CarMake', 'Comments']
        if not all(col in df.columns for col in expected_columns):
            missing_cols = [col for col in expected_columns if col not in df.columns]
            raise ValueError(f"Missing expected columns in the Excel file: {', '.join(missing_cols)}")
        df = df.rename(columns={
            'Generic System Name': 'genericSystemName',
            'DTCsys': 'dtcSys',
            'CarMake': 'carMake',
            'Comments': 'comments'
        })
        df = df[['genericSystemName', 'dtcSys', 'carMake', 'comments']]
        df['genericSystemName'] = df['genericSystemName'].astype(str)
        df['dtcSys'] = df['dtcSys'].astype(str)
        df['carMake'] = df['carMake'].astype(str)
        df['comments'] = df['comments'].astype(str)
        df = df.where(pd.notnull(df), None)
        df.dropna(how='all', inplace=True)
        df.to_sql(table_name, conn, if_exists='replace', index=False)
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
        return "\n".join(error_messages)
    return "Data loaded successfully"

def load_mag_glass_data(excel_path, table_name='mag_glass', db_path='data.db'):
    import pandas as pd, sqlite3, logging
    from PyQt5.QtWidgets import QMessageBox
    error_messages = []
    try:
        try:
            df = pd.read_excel(excel_path, sheet_name="Mag Glass")
            logging.debug(f"Data loaded from 'Mag Glass' sheet: {df.head()}")
        except Exception as sheet_error:
            error_message = f"Failed to find 'Mag Glass' sheet in {excel_path}: {str(sheet_error)}"
            logging.warning(error_message)
            error_messages.append(error_message)
            return error_message
        df = df.where(pd.notnull(df), None)
        df.dropna(how='all', inplace=True)
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

def load_mag_glass_data_from_5th_sheet(excel_path, table_name='mag_glass', db_path='data.db'):
    return load_mag_glass_data(excel_path, table_name, db_path)

class PopOutWindow(ModernDialog):
    def __init__(self, title, content, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 800, 600)
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        
        layout = QVBoxLayout()
        
        # Create modern card container
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Header with title and close button
        header_layout = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #495057; margin: 10px;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        close_button = ModernButton("✕", style="secondary")
        close_button.setFixedSize(30, 30)
        close_button.clicked.connect(self.close)
        header_layout.addWidget(close_button)
        card_layout.addLayout(header_layout)
        
        # Content area
        self.text_browser = ModernTextBrowser()
        self.text_browser.setHtml(content)
        card_layout.addWidget(self.text_browser)
        
        # Transparency controls
        transparency_layout = QHBoxLayout()
        self.opacity_label = QLabel("Transparency")
        self.opacity_label.setStyleSheet("font-weight: 600; color: #495057;")
        transparency_layout.addWidget(self.opacity_label)
        
        self.opacity_slider = ModernSlider(Qt.Horizontal)
        self.opacity_slider.setMinimum(20)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.valueChanged.connect(self.change_opacity)
        transparency_layout.addWidget(self.opacity_slider)
        
        card_layout.addLayout(transparency_layout)
        layout.addWidget(card)
        self.setLayout(layout)

    def change_opacity(self, value):
        self.setWindowOpacity(value / 100.0)

class UserLoginDialog(ModernDialog):
    def __init__(self, parent=None, theme='Dark'):
        super().__init__(parent)
        self.setWindowTitle("User Login")
        self.setGeometry(100, 100, 400, 300)
        self.setModal(True)
        self.theme = theme
        self.result_code = QDialog.Rejected
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("Analyzer+")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #667eea; margin: 20px;")
        card_layout.addWidget(title_label)
        
        # Subtitle
        subtitle_label = QLabel("Professional Vehicle Analysis System")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("font-size: 14px; color: #6c757d; margin-bottom: 20px;")
        card_layout.addWidget(subtitle_label)
        
        # PIN input
        instruction_label = QLabel("Enter your PIN to continue:")
        instruction_label.setAlignment(Qt.AlignCenter)
        instruction_label.setStyleSheet("font-weight: 600; color: #495057; margin: 10px;")
        card_layout.addWidget(instruction_label)
        
        self.pin_input = ModernLineEdit("Enter 4-digit PIN")
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        self.pin_input.setAlignment(Qt.AlignCenter)
        self.pin_input.returnPressed.connect(self.accept)
        card_layout.addWidget(self.pin_input)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = ModernButton("Cancel", style="secondary")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.login_button = ModernButton("Login", style="primary")
        self.login_button.clicked.connect(self.accept)
        button_layout.addWidget(self.login_button)
        
        card_layout.addLayout(button_layout)
        
        # Additional options
        options_layout = QVBoxLayout()
        
        # Signup option
        signup_layout = QHBoxLayout()
        signup_label = QLabel("New user?")
        signup_label.setStyleSheet("color: #6c757d;")
        signup_layout.addWidget(signup_label)
        
        self.signup_button = ModernButton("Create Account", style="secondary")
        self.signup_button.clicked.connect(self.on_signup_clicked)
        signup_layout.addWidget(self.signup_button)
        options_layout.addLayout(signup_layout)
        
        # Reset PIN option
        reset_layout = QHBoxLayout()
        reset_label = QLabel("Forgot PIN?")
        reset_label.setStyleSheet("color: #6c757d;")
        reset_layout.addWidget(reset_label)
        
        self.reset_button = ModernButton("Reset PIN", style="secondary")
        self.reset_button.clicked.connect(self.on_reset_clicked)
        reset_layout.addWidget(self.reset_button)
        options_layout.addLayout(reset_layout)
        
        card_layout.addLayout(options_layout)
        layout.addWidget(card)
        self.setLayout(layout)
    
    def on_signup_clicked(self):
        self.result_code = 2
        self.done(2)
    
    def on_reset_clicked(self):
        self.result_code = 3
        self.done(3)
    
    def get_pin(self):
        return self.pin_input.text()
    
    def exec_(self):
        self.pin_input.setFocus()
        return super().exec_()

class PinManagementDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("User Management")
        self.setGeometry(100, 100, 600, 500)
        self.setModal(True)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("User Management")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.add_button = ModernButton("Add User", style="primary")
        self.add_button.clicked.connect(self.add_user)
        button_layout.addWidget(self.add_button)
        
        self.edit_button = ModernButton("Edit User", style="secondary")
        self.edit_button.clicked.connect(self.edit_user)
        button_layout.addWidget(self.edit_button)
        
        self.delete_button = ModernButton("Delete User", style="danger")
        self.delete_button.clicked.connect(self.delete_user)
        button_layout.addWidget(self.delete_button)
        
        self.show_all_button = ModernButton("Show All Users", style="secondary")
        self.show_all_button.clicked.connect(self.show_all_users)
        button_layout.addWidget(self.show_all_button)
        
        card_layout.addLayout(button_layout)
        
        # User list
        self.user_list = QListWidget()
        self.user_list.setStyleSheet("""
            QListWidget {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                background: white;
                font-size: 14px;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f8f9fa;
            }
            QListWidget::item:selected {
                background: #667eea;
                color: white;
            }
        """)
        card_layout.addWidget(self.user_list)
        
        layout.addWidget(card)
        self.setLayout(layout)
        
        self.load_users_and_pins()
    
    def load_users_and_pins(self):
        """Load users from database and display in list"""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT pin, name FROM leader_log ORDER BY name')
            users = cursor.fetchall()
            conn.close()
            
            self.user_list.clear()
            for pin, name in users:
                self.user_list.addItem(f"{name} (PIN: {pin})")
        except Exception as e:
            logging.error(f"Error loading users: {e}")
    
    def show_all_users(self):
        """Show all users in a message box"""
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT pin, name FROM leader_log ORDER BY name')
            users = cursor.fetchall()
            conn.close()
            
            if users:
                user_text = "Current Users:\n\n"
                for pin, name in users:
                    user_text += f"• {name} (PIN: {pin})\n"
            else:
                user_text = "No users found."
            
            msg = QMessageBox()
            msg.setWindowTitle("All Users")
            msg.setText(user_text)
            msg.setIcon(QMessageBox.Information)
            msg.exec_()
        except Exception as e:
            logging.error(f"Error showing users: {e}")
    
    def add_user(self):
        """Add a new user"""
        dialog = AddUserDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_users_and_pins()
    
    def edit_user(self):
        """Edit selected user"""
        current_item = self.user_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a user to edit.")
            return
        
        user_text = current_item.text()
        name = user_text.split(" (PIN:")[0]
        
        dialog = EditUserDialog(self, name)
        if dialog.exec_() == QDialog.Accepted:
            self.load_users_and_pins()
    
    def delete_user(self):
        """Delete selected user"""
        current_item = self.user_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a user to delete.")
            return
        
        user_text = current_item.text()
        name = user_text.split(" (PIN:")[0]
        
        reply = QMessageBox.question(self, "Confirm Delete", 
                                   f"Are you sure you want to delete user '{name}'?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute('DELETE FROM leader_log WHERE name = ?', (name,))
                conn.commit()
                conn.close()
                self.load_users_and_pins()
                QMessageBox.information(self, "Success", f"User '{name}' deleted successfully.")
            except Exception as e:
                logging.error(f"Error deleting user: {e}")
                QMessageBox.critical(self, "Error", f"Failed to delete user: {str(e)}")

class AddUserDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New User")
        self.setGeometry(100, 100, 400, 300)
        self.setModal(True)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("Add New User")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Form
        form_layout = QFormLayout()
        
        self.name_input = ModernLineEdit("Enter user name")
        form_layout.addRow("Name:", self.name_input)
        
        self.pin_input = ModernLineEdit("Enter 4-digit PIN")
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        form_layout.addRow("PIN:", self.pin_input)
        
        card_layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = ModernButton("Cancel", style="secondary")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.save_button = ModernButton("Save", style="primary")
        self.save_button.clicked.connect(self.save_user)
        button_layout.addWidget(self.save_button)
        
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        self.setLayout(layout)
    
    def save_user(self):
        name = self.name_input.text().strip()
        pin = self.pin_input.text().strip()
        
        if not name or not pin:
            QMessageBox.warning(self, "Error", "Please fill in all fields.")
            return
        
        if len(pin) != 4 or not pin.isdigit():
            QMessageBox.warning(self, "Error", "PIN must be 4 digits.")
            return
        
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', (pin, name))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", f"User '{name}' added successfully.")
            self.accept()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "This PIN is already in use.")
        except Exception as e:
            logging.error(f"Error adding user: {e}")
            QMessageBox.critical(self, "Error", f"Failed to add user: {str(e)}")

class EditUserDialog(ModernDialog):
    def __init__(self, parent=None, user_name=""):
        super().__init__(parent)
        self.user_name = user_name
        self.setWindowTitle("Edit User")
        self.setGeometry(100, 100, 400, 300)
        self.setModal(True)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel(f"Edit User: {self.user_name}")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Form
        form_layout = QFormLayout()
        
        self.name_input = ModernLineEdit(self.user_name)
        form_layout.addRow("Name:", self.name_input)
        
        # Get current PIN
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT pin FROM leader_log WHERE name = ?', (self.user_name,))
            result = cursor.fetchone()
            conn.close()
            current_pin = result[0] if result else ""
        except:
            current_pin = ""
        
        self.pin_input = ModernLineEdit(current_pin)
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        form_layout.addRow("PIN:", self.pin_input)
        
        card_layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = ModernButton("Cancel", style="secondary")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.save_button = ModernButton("Save", style="primary")
        self.save_button.clicked.connect(self.save_user)
        button_layout.addWidget(self.save_button)
        
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        self.setLayout(layout)
    
    def save_user(self):
        name = self.name_input.text().strip()
        pin = self.pin_input.text().strip()
        
        if not name or not pin:
            QMessageBox.warning(self, "Error", "Please fill in all fields.")
            return
        
        if len(pin) != 4 or not pin.isdigit():
            QMessageBox.warning(self, "Error", "PIN must be 4 digits.")
            return
        
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('UPDATE leader_log SET pin = ?, name = ? WHERE name = ?', 
                          (pin, name, self.user_name))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", f"User updated successfully.")
            self.accept()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "This PIN is already in use by another user.")
        except Exception as e:
            logging.error(f"Error updating user: {e}")
            QMessageBox.critical(self, "Error", f"Failed to update user: {str(e)}")

class AdminOptionsDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Admin Options")
        self.setGeometry(100, 100, 400, 300)
        self.setModal(True)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("Administrative Options")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Buttons
        self.user_management_button = ModernButton("User Management", style="primary")
        self.user_management_button.clicked.connect(self.update_users)
        card_layout.addWidget(self.user_management_button)
        
        self.export_button = ModernButton("Export Data", style="secondary")
        self.export_button.clicked.connect(self.export_data)
        card_layout.addWidget(self.export_button)
        
        layout.addWidget(card)
        self.setLayout(layout)
    
    def update_users(self):
        dialog = PinManagementDialog(self)
        dialog.exec_()
    
    def export_data(self):
        """Export data functionality"""
        # This will be implemented in the main app
        self.parent().export_data()

class SignupDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create Account")
        self.setGeometry(100, 100, 450, 400)
        self.setModal(True)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("Create New Account")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Form
        form_layout = QFormLayout()
        
        self.first_name_input = ModernLineEdit("Enter first name")
        self.first_name_input.textChanged.connect(self.update_username_preview)
        form_layout.addRow("First Name:", self.first_name_input)
        
        self.last_name_input = ModernLineEdit("Enter last name")
        self.last_name_input.textChanged.connect(self.update_username_preview)
        form_layout.addRow("Last Name:", self.last_name_input)
        
        # Username preview
        self.username_preview = QLabel("")
        self.username_preview.setStyleSheet("font-weight: bold; color: #667eea; padding: 5px;")
        form_layout.addRow("Username:", self.username_preview)
        
        self.pin_input = ModernLineEdit("Enter 4-digit PIN")
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        form_layout.addRow("PIN:", self.pin_input)
        
        self.confirm_pin_input = ModernLineEdit("Confirm PIN")
        self.confirm_pin_input.setEchoMode(QLineEdit.Password)
        self.confirm_pin_input.setMaxLength(4)
        form_layout.addRow("Confirm PIN:", self.confirm_pin_input)
        
        card_layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = ModernButton("Cancel", style="secondary")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.create_button = ModernButton("Create Account", style="primary")
        self.create_button.clicked.connect(self.create_user)
        button_layout.addWidget(self.create_button)
        
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        self.setLayout(layout)
    
    def update_username_preview(self):
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        
        if first_name and last_name:
            username = first_name[0] + last_name
            self.username_preview.setText(username)
        else:
            self.username_preview.setText("(Example: JDoe)")
    
    def create_user(self):
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        pin = self.pin_input.text().strip()
        confirm_pin = self.confirm_pin_input.text().strip()
        
        # Validation
        if not first_name or not last_name:
            QMessageBox.warning(self, "Error", "Please enter both first and last name.")
            return
        
        if not pin:
            QMessageBox.warning(self, "Error", "Please enter a PIN.")
            return
        
        if pin != confirm_pin:
            QMessageBox.warning(self, "Error", "PINs do not match.")
            return
        
        if len(pin) != 4 or not pin.isdigit():
            QMessageBox.warning(self, "Error", "PIN must be 4 digits.")
            return
        
        username = first_name[0] + last_name
        
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', (pin, username))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", f"Account created successfully!\nUsername: {username}\nPIN: {pin}")
            self.accept()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Error", "This PIN is already in use. Please choose a different one.")
        except Exception as e:
            logging.error(f"Error creating user: {e}")
            QMessageBox.critical(self, "Error", f"Failed to create account: {str(e)}")

class ResetPinDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reset PIN")
        self.setGeometry(100, 100, 450, 400)
        self.setModal(True)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create modern card
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Title
        title_label = QLabel("Reset Your PIN")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #495057; margin: 15px;")
        card_layout.addWidget(title_label)
        
        # Form
        form_layout = QFormLayout()
        
        self.first_name_input = ModernLineEdit("Enter first name")
        self.first_name_input.textChanged.connect(self.update_username_preview)
        form_layout.addRow("First Name:", self.first_name_input)
        
        self.last_name_input = ModernLineEdit("Enter last name")
        self.last_name_input.textChanged.connect(self.update_username_preview)
        form_layout.addRow("Last Name:", self.last_name_input)
        
        # Username preview
        self.username_preview = QLabel("")
        self.username_preview.setStyleSheet("font-weight: bold; color: #667eea; padding: 5px;")
        form_layout.addRow("Your username:", self.username_preview)
        
        self.pin_input = ModernLineEdit("Enter new 4-digit PIN")
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        form_layout.addRow("New PIN:", self.pin_input)
        
        self.confirm_pin_input = ModernLineEdit("Confirm new PIN")
        self.confirm_pin_input.setEchoMode(QLineEdit.Password)
        self.confirm_pin_input.setMaxLength(4)
        form_layout.addRow("Confirm PIN:", self.confirm_pin_input)
        
        card_layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = ModernButton("Cancel", style="secondary")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.reset_button = ModernButton("Reset PIN", style="primary")
        self.reset_button.clicked.connect(self.reset_pin)
        button_layout.addWidget(self.reset_button)
        
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        self.setLayout(layout)
    
    def update_username_preview(self):
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        
        if first_name and last_name:
            username = first_name[0] + last_name
            self.username_preview.setText(username)
        else:
            self.username_preview.setText("(Example: JDoe)")
    
    def reset_pin(self):
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        pin = self.pin_input.text().strip()
        confirm_pin = self.confirm_pin_input.text().strip()
        
        # Validation
        if not first_name or not last_name:
            QMessageBox.warning(self, "Error", "Please enter both first and last name.")
            return
        
        if not pin:
            QMessageBox.warning(self, "Error", "Please enter a new PIN.")
            return
        
        if pin != confirm_pin:
            QMessageBox.warning(self, "Error", "PINs do not match.")
            return
        
        if len(pin) != 4 or not pin.isdigit():
            QMessageBox.warning(self, "Error", "PIN must be 4 digits.")
            return
        
        expected_username = first_name[0] + last_name
        
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            
            # Check if user exists
            cursor.execute('SELECT pin FROM leader_log WHERE name = ?', (expected_username,))
            existing_user = cursor.fetchone()
            
            if not existing_user:
                QMessageBox.warning(self, "Account Not Found", 
                                   f"No user found with username '{expected_username}'.\n"
                                   "Please check your name spelling or create a new account.")
                conn.close()
                return
            
            # Check if new PIN is already used by another user
            old_pin = existing_user[0]
            if pin != old_pin:
                cursor.execute('SELECT name FROM leader_log WHERE pin = ? AND pin != ?', (pin, old_pin))
                pin_user = cursor.fetchone()
                if pin_user:
                    QMessageBox.warning(self, "Error", "This PIN is already in use by another user.")
                    conn.close()
                    return
            
            # Update PIN
            cursor.execute('UPDATE leader_log SET pin = ? WHERE name = ?', (pin, expected_username))
            conn.commit()
            conn.close()
            
            QMessageBox.information(self, "Success", "PIN reset successfully!")
            self.accept()
        except Exception as e:
            logging.error(f"Failed to reset PIN: {e}")
            QMessageBox.critical(self, "Error", f"Failed to reset PIN: {str(e)}") 

class ManageDataListsDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Data Lists")
        self.setGeometry(100, 100, 700, 500)
        self.setModal(True)
        self.parent = parent
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        card = ModernCard()
        card_layout = QVBoxLayout(card)

        # Title
        title_label = QLabel("Manage Data Lists")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #495057; margin: 10px;")
        card_layout.addWidget(title_label)

        note_label = QLabel("<i>Note: CarSYS and MagGlass automatically use the same path as Goldlist</i>")
        note_label.setAlignment(Qt.AlignCenter)
        note_label.setStyleSheet("font-size: 13px; color: #6c757d; margin-bottom: 10px;")
        card_layout.addWidget(note_label)

        self.path_fields = {}
        self.browse_buttons = {}
        self.clear_buttons = {}
        config_display_names = {
            "blacklist": "Blacklist",
            "goldlist": "Goldlist",
            "prequal": "Prequals (Longsheets)"
        }

        for config_key, display_name in config_display_names.items():
            section_card = ModernCard()
            section_layout = QVBoxLayout(section_card)
            section_label = QLabel(display_name)
            section_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #495057;")
            section_layout.addWidget(section_label)

            button_row = QHBoxLayout()
            browse_btn = ModernButton("Choose Path", style="secondary")
            browse_btn.clicked.connect(lambda checked, k=config_key: self.choose_path(k))
            button_row.addWidget(browse_btn)
            self.browse_buttons[config_key] = browse_btn

            clear_btn = ModernButton("Clear Data", style="danger")
            clear_btn.clicked.connect(lambda checked, k=config_key: self.clear_data(k))
            button_row.addWidget(clear_btn)
            self.clear_buttons[config_key] = clear_btn

            section_layout.addLayout(button_row)

            # Path field
            path_field = ModernLineEdit(f"Select a directory for {display_name}")
            path_field.setReadOnly(True)
            section_layout.addWidget(path_field)
            self.path_fields[config_key] = path_field

            card_layout.addWidget(section_card)

        # Buttons at the bottom
        button_row = QHBoxLayout()
        cancel_btn = ModernButton("Cancel", style="secondary")
        cancel_btn.clicked.connect(self.reject)
        button_row.addWidget(cancel_btn)

        clear_all_btn = ModernButton("Clear All Data", style="danger")
        clear_all_btn.clicked.connect(self.clear_all_data)
        button_row.addWidget(clear_all_btn)

        save_btn = ModernButton("Save & Load Data", style="primary")
        save_btn.clicked.connect(self.save_and_load)
        button_row.addWidget(save_btn)

        card_layout.addLayout(button_row)
        layout.addWidget(card)
        self.setLayout(layout)

        # Pre-populate fields with existing paths
        for config_key in self.path_fields:
            existing_path = load_path_from_db(config_key, self.parent.db_path)
            if existing_path:
                self.path_fields[config_key].setText(existing_path)

    def choose_path(self, config_key):
        folder_path = QFileDialog.getExistingDirectory(self, f"Select directory for {config_key.capitalize()}")
        if folder_path:
            self.path_fields[config_key].setText(folder_path)

    def clear_data(self, config_key):
        reply = QMessageBox.question(self, "Clear Data", f"Are you sure you want to clear all {config_key.capitalize()} data?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            if self.parent:
                self.parent.clear_data(config_key)
            self.path_fields[config_key].clear()
            QMessageBox.information(self, "Data Cleared", f"{config_key.capitalize()} data has been cleared.")

    def clear_all_data(self):
        reply = QMessageBox.question(self, "Clear All Data", "Are you sure you want to clear ALL data? This cannot be undone.", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            if self.parent:
                self.parent.clear_data()
            for field in self.path_fields.values():
                field.clear()
            QMessageBox.information(self, "Data Cleared", "All data has been cleared.")

    def save_and_load(self):
        paths_to_save = {}
        for config_type, path_field in self.path_fields.items():
            folder_path = path_field.text()
            if folder_path:
                paths_to_save[config_type] = folder_path

        if not paths_to_save:
            QMessageBox.warning(self, "No Paths", "No paths have been selected.")
            return

        goldlist_path = paths_to_save.get("goldlist")

        for config_type, folder_path in paths_to_save.items():
            try:
                files = self.parent.get_valid_excel_files(folder_path)
                if not files:
                    QMessageBox.warning(self, "Load Error", f"No valid Excel files found in the {folder_path} directory for {config_type}.")
                    continue

                self.parent.progress_bar.setVisible(True)
                self.parent.progress_bar.setMaximum(len(files))
                self.parent.progress_bar.setValue(0)

                self.parent.clear_data(config_type)
                if config_type == "goldlist":
                    self.parent.clear_data("CarSys")
                    self.parent.clear_data("mag_glass")

                data_loaded = False
                for i, (filename, filepath) in enumerate(files.items()):
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.parent.db_path, sheet_index=1)
                            if result == "Data loaded successfully":
                                data_loaded = True
                                if config_type == 'goldlist':
                                    load_carsys_data_to_db(filepath, table_name='CarSys', db_path=self.parent.db_path)
                                    load_mag_glass_data(filepath, table_name='mag_glass', db_path=self.parent.db_path)
                        elif config_type == 'prequal':
                            df = pd.read_excel(filepath)
                            if df.empty:
                                logging.warning(f"{filename} is empty.")
                                continue
                            data = df.to_dict(orient='records')
                            update_configuration(config_type, folder_path, data, self.parent.db_path)
                            data_loaded = True
                        elif config_type == 'mag_glass':
                            result = load_mag_glass_data(filepath, config_type, db_path=self.parent.db_path)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'CarSys':
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.parent.db_path, sheet_index=0)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        else:
                            df = pd.read_excel(filepath)
                            if df.empty:
                                logging.warning(f"{filename} is empty.")
                                continue
                            data = df.to_dict(orient='records')
                            update_configuration(config_type, folder_path, data, self.parent.db_path)
                            data_loaded = True
                    except Exception as e:
                        logging.error(f"Error loading {filename}: {str(e)}")
                        QMessageBox.critical(self, "Load Error", f"Failed to load {filename}: {str(e)}")
                    self.parent.progress_bar.setValue(i + 1)

                save_path_to_db(config_type, folder_path, self.parent.db_path)
                if config_type == 'goldlist':
                    save_path_to_db('CarSys', folder_path, self.parent.db_path)
                    save_path_to_db('mag_glass', folder_path, self.parent.db_path)

            except Exception as e:
                logging.error(f"Error saving path for {config_type}: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save path for {config_type}: {str(e)}")

        self.parent.progress_bar.setVisible(False)
        self.parent.load_configurations()
        self.parent.populate_dropdowns()
        self.parent.check_data_loaded()
        msg = self.parent.create_styled_messagebox("Success", "Paths saved and data loaded successfully!", QMessageBox.Information)
        msg.exec_()
        self.accept()

def get_theme_palette(theme):
    palettes = {
        "Light": {
            "bg": "#f8f9fa",
            "card": "#ffffff",
            "accent": "#667eea",
            "accent2": "#764ba2",
            "text": "#495057",
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#28a745"
        },
        "Dark": {
            "bg": "#181a20",  # much darker background
            "card": "#23262f",  # dark card color
            "accent": "#4f8cff",
            "accent2": "#23272f",
            "text": "#23262f",  # very light text
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#28a745"
        },
        "Blue": {
            "bg": "#e3f0ff",
            "card": "#f7fbff",
            "accent": "#2196f3",
            "accent2": "#1565c0",
            "text": "#1a237e",
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#28a745"
        },
        "Green": {
            "bg": "#eafaf1",
            "card": "#f6fff9",
            "accent": "#43a047",
            "accent2": "#1b5e20",
            "text": "#1b5e20",
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#388e3c"
        },
        "Purple": {
            "bg": "#f3e8ff",
            "card": "#faf5ff",
            "accent": "#8e24aa",
            "accent2": "#4a148c",
            "text": "#4a148c",
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#28a745"
        },
        "Orange": {
            "bg": "#fff3e0",
            "card": "#fff8e1",
            "accent": "#ff9800",
            "accent2": "#e65100",
            "text": "#e65100",
            "button_fg": "#fff",
            "danger": "#dc3545",
            "success": "#28a745"
        }
    }
    return palettes.get(theme, palettes["Light"])

class ModernAnalyzerApp(ModernMainWindow):
    def __init__(self, db_path='data.db'):
        super().__init__()
        self.db_path = db_path
        self.settings_file = 'settings.json'
        self.current_theme = 'Light'  # Start with light theme for modern look
        self.current_user = None
        self.connection_pool = None
        self.conn = sqlite3.connect(self.db_path)
        self.adas_authenticated = False
        initialize_db(self.db_path)
        self.current_theme = self.get_last_logged_theme()
        self.data = {'blacklist': [], 'goldlist': [], 'prequal': [], 'mag_glass': [], 'carsys': []}
        self.make_map = {}
        self.model_map = {}
        self.setup_ui()
        self.prompt_user_pin()
        self.load_configurations()
        self.check_data_loaded()
        self.apply_saved_theme()

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
        return 'Light'

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

    def create_styled_messagebox(self, title, text, icon_type=QMessageBox.Warning):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(text)
        msg_box.setIcon(icon_type)
        
        style = """
        QMessageBox {
            background-color: #ffffff;
            color: #495057;
        }
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #667eea, stop:1 #764ba2);
            border: none;
            border-radius: 6px;
            color: white;
            font-weight: 600;
            padding: 8px 16px;
            min-width: 80px;
        }
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #5a6fd8, stop:1 #6a4190);
        }
        QLabel {
            color: #495057;
        }
        """
        
        msg_box.setStyleSheet(style)
        return msg_box

    def prompt_user_pin(self):
        attempts = 0
        max_attempts = 3
        while attempts < max_attempts:
            attempts += 1
            login_dialog = UserLoginDialog(self, self.current_theme)
            result = login_dialog.exec_()
            
            if result == QDialog.Accepted:
                pin = login_dialog.get_pin()
                if pin:
                    conn = self.get_db_connection()
                    cursor = conn.cursor()
                    
                    logging.debug(f"Login attempt with PIN: {pin}")
                    
                    debug_cursor = conn.cursor()
                    debug_cursor.execute('SELECT pin, name FROM leader_log')
                    all_users = debug_cursor.fetchall()
                    logging.debug(f"All users in database: {all_users}")
                    
                    if pin == '1234':
                        self.current_user = "Emergency Admin"
                        logging.debug("Logged in with emergency PIN")
                        self.log_action(self.current_user, "Logged in with emergency PIN")
                        conn.close()
                        
                        try:
                            setup_cursor = conn.cursor()
                            setup_cursor.execute('SELECT * FROM leader_log WHERE name = "Set Up"')
                            if not setup_cursor.fetchone():
                                setup_cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', ('0000', 'Set Up'))
                                conn.commit()
                                logging.debug("Created default 'Set Up' user with PIN 0000")
                        except Exception as e:
                            logging.error(f"Error creating default user: {e}")
                        
                        return True
                        
                    cursor.execute('SELECT name FROM leader_log WHERE pin = ?', (pin,))
                    result = cursor.fetchone()
                    if result:
                        self.current_user = result[0]
                        logging.debug(f"Successful login for user: {self.current_user}")
                        self.log_action(self.current_user, "Logged in")
                        conn.close()
                        return True
                    elif pin == '9716':
                        self.current_user = "Set Up"
                        logging.debug("Logged in with standard PIN")
                        self.log_action(self.current_user, "Logged in with standard PIN")
                        conn.close()
                        return True
                    else:
                        debug_msg = "Available users in database:\n"
                        for db_pin, db_name in all_users:
                            debug_msg += f"User: {db_name}, PIN: {db_pin}\n"
                        
                        logging.debug(debug_msg)
                        
                        msg_box = self.create_styled_messagebox(
                            "Access Denied", 
                            f"Incorrect User PIN: {pin}\n\nTry one of these valid PINs:\n" + 
                            "\n".join([f"{u[0]} for {u[1]}" for u in all_users]) +
                            "\n\nOr use PIN 1234 for emergency access."
                        )
                        msg_box.exec_()
                        self.log_action("Unknown", f"Failed User PIN attempt with PIN: {pin}")
                        conn.close()
                else:
                    msg_box = self.create_styled_messagebox("Access Denied", "PIN cannot be empty.")
                    msg_box.exec_()
            elif result == 2:
                signup_dialog = SignupDialog(self)
                if signup_dialog.exec_() == QDialog.Accepted:
                    logging.debug("User was created, continuing with login prompt")
                    attempts -= 1
                    continue
            elif result == 3:
                reset_dialog = ResetPinDialog(self)
                if reset_dialog.exec_() == QDialog.Accepted:
                    logging.debug("PIN was reset, continuing with login prompt")
                    attempts -= 1
                    continue
            else:
                msg_box = QMessageBox()
                msg_box.setWindowTitle("Exit Application?")
                msg_box.setText("Do you want to exit the application?")
                msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                msg_box.setDefaultButton(QMessageBox.No)
                
                if msg_box.exec_() == QMessageBox.Yes:
                    logging.debug("User chose to exit the application")
                    sys.exit()
                else:
                    logging.debug("User chose to return to login")
                    attempts -= 1
                    continue
        
        msg_box = self.create_styled_messagebox(
            "Login Failed", 
            f"Maximum login attempts ({max_attempts}) reached. The application will now exit."
        )
        msg_box.exec_()
        sys.exit()

    def setup_ui(self):
        self.setWindowTitle("Analyzer+ Professional")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1200, 800)
        
        # Set application icon and style
        self.setStyle(QStyleFactory.create('Fusion'))
        
        # Create central widget with modern layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout with modern spacing
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Create modern header
        self.create_header(main_layout)
        
        # Create modern toolbar
        self.create_toolbar()
        
        # Create search and filter section
        self.create_search_section(main_layout)
        
        # Create main content area with multi-toggle tab row and horizontal split display
        self.create_main_content(main_layout)
        
        # Create modern status bar
        self.create_status_bar()
        
        # Apply modern styling
        self.apply_modern_theme()

    def create_header(self, main_layout):
        """Create a professional header with logo, brand name, and user info"""
        header_card = ModernCard()
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(18, 12, 18, 12)
        header_layout.setSpacing(18)

        # Logo and title section
        logo_title_layout = QHBoxLayout()
        logo_title_layout.setSpacing(12)

        # Logo (assume 'protech_logo.png' is in the working directory)
        logo_label = QLabel()
        logo_label.setPixmap(QPixmap('protech_logo.png').scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        logo_label.setFixedSize(48, 48)
        logo_title_layout.addWidget(logo_label)

        # Brand name and subtitle
        brand_layout = QVBoxLayout()
        brand_layout.setSpacing(0)
        brand_label = QLabel("Protech Automotive Solutions")
        brand_label.setObjectName("brand_label")
        brand_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #20567C; font-family: 'Segoe UI', Arial, sans-serif; margin: 0;")
        brand_layout.addWidget(brand_label)
        subtitle_label = QLabel("Advanced Vehicle Diagnostic Analysis System")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_label.setStyleSheet("font-size: 14px; color: #3a4a5a; font-family: 'Segoe UI', Arial, sans-serif; margin: 0;")
        brand_layout.addWidget(subtitle_label)
        logo_title_layout.addLayout(brand_layout)
        header_layout.addLayout(logo_title_layout)

        header_layout.addStretch()

        # User info section
        user_card = ModernCard()
        user_card.setFixedSize(220, 60)
        user_layout = QVBoxLayout(user_card)
        user_layout.setContentsMargins(10, 6, 10, 6)
        self.user_label = QLabel("Welcome")
        self.user_label.setStyleSheet("font-size: 15px; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;")
        user_layout.addWidget(self.user_label)
        self.user_status = QLabel("Ready")
        self.user_status.setStyleSheet("font-size: 12px; color: #28a745; font-family: 'Segoe UI', Arial, sans-serif;")
        user_layout.addWidget(self.user_status)
        header_layout.addWidget(user_card)

        main_layout.addWidget(header_card)

    def create_toolbar(self):
        """Create a professional, branded toolbar"""
        self.toolbar = ModernToolBar()
        self.toolbar.setStyleSheet("""
            QToolBar {
                background: #20567C;
                border-bottom: 2px solid #102535;
                padding: 8px 16px;
                spacing: 10px;
                font-family: 'Segoe UI', Arial, sans-serif;
                box-shadow: 0 2px 8px rgba(21,52,74,0.08);
            }
            QToolBar QToolButton, QPushButton {
                background: #19507a;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 18px;
                font-size: 15px;
                font-weight: 600;
                margin-right: 8px;
            }
            QToolBar QToolButton:hover, QPushButton:hover {
                background: #1e5d8a;
            }
            QToolBar QToolButton:pressed, QPushButton:pressed {
                background: #102535;
            }
            QLabel {
                color: #fff;
                font-size: 15px;
                font-weight: 500;
            }
            QComboBox {
                background: #fff;
                color: #20567C;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 14px;
                min-width: 110px;
            }
        """)
        self.addToolBar(self.toolbar)
        self.add_toolbar_button("Manage Lists", self.open_admin, "admin_button")
        self.add_toolbar_button("Refresh Lists", self.refresh_lists, "refresh_button")
        clear_button = ModernButton("Clear Filters", style="secondary")
        clear_button.clicked.connect(self.clear_filters)
        self.toolbar.addWidget(clear_button)
        self.always_on_top_button = ModernButton("Pin to Top", style="secondary")
        self.always_on_top_button.setCheckable(True)
        self.always_on_top_button.clicked.connect(self.toggle_always_on_top)
        self.toolbar.addWidget(self.always_on_top_button)
        self.toolbar.addSeparator()
        theme_label = QLabel("Theme:")
        self.toolbar.addWidget(theme_label)
        self.theme_dropdown = ModernComboBox()
        themes = ["Light", "Dark", "Blue", "Green", "Purple", "Orange"]
        self.theme_dropdown.addItems(themes)
        self.theme_dropdown.currentIndexChanged.connect(self.apply_selected_theme)
        self.toolbar.addWidget(self.theme_dropdown)

    def create_search_section(self, main_layout):
        """Create a professional, compact search/filter row"""
        search_card = ModernCard()
        search_card.setStyleSheet("""
            QWidget {
                background: #fff;
                border-radius: 14px;
                box-shadow: 0 2px 8px rgba(21,52,74,0.08);
                padding: 12px 18px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
        search_layout = QHBoxLayout(search_card)
        search_layout.setSpacing(16)
        search_layout.setContentsMargins(18, 10, 18, 10)

        # Year
        year_label = QLabel("Year:")
        year_label.setStyleSheet("font-weight: 600; font-size: 14px; color: #20567C; margin-right: 4px;")
        search_layout.addWidget(year_label)
        self.year_dropdown = ModernComboBox()
        self.year_dropdown.setFixedWidth(120)
        self.year_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.currentIndexChanged.connect(self.on_year_selected)
        search_layout.addWidget(self.year_dropdown)

        # Make
        make_label = QLabel("Make:")
        make_label.setStyleSheet("font-weight: 600; font-size: 14px; color: #20567C; margin-left: 8px; margin-right: 4px;")
        search_layout.addWidget(make_label)
        self.make_dropdown = ModernComboBox()
        self.make_dropdown.setFixedWidth(120)
        self.make_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        self.make_dropdown.currentIndexChanged.connect(self.update_model_dropdown)
        search_layout.addWidget(self.make_dropdown)

        # Model
        model_label = QLabel("Model:")
        model_label.setStyleSheet("font-weight: 600; font-size: 14px; color: #20567C; margin-left: 8px; margin-right: 4px;")
        search_layout.addWidget(model_label)
        self.model_dropdown = ModernComboBox()
        self.model_dropdown.setFixedWidth(130)
        self.model_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.model_dropdown.addItem("Select Model")
        self.model_dropdown.currentIndexChanged.connect(self.handle_model_change)
        search_layout.addWidget(self.model_dropdown)

        # Search bar
        search_label = QLabel("Search:")
        search_label.setStyleSheet("font-weight: 600; font-size: 14px; color: #20567C; margin-left: 12px; margin-right: 4px;")
        search_layout.addWidget(search_label)
        self.search_bar = ModernLineEdit("Enter DTC code or description")
        self.search_bar.setFixedWidth(210)
        self.search_bar.setStyleSheet("padding: 6px 12px; font-size: 12px;")
        self.search_bar.returnPressed.connect(self.perform_search)
        self.search_bar.textChanged.connect(self.update_suggestions)
        search_layout.addWidget(self.search_bar)

        # Filter
        filter_label = QLabel("Filter:")
        filter_label.setStyleSheet("font-weight: 600; font-size: 14px; color: #20567C; margin-left: 12px; margin-right: 4px;")
        search_layout.addWidget(filter_label)
        self.filter_dropdown = ModernComboBox()
        self.filter_dropdown.setFixedWidth(130)
        self.filter_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.filter_dropdown.addItems(["Select List", "Blacklist", "Goldlist", "Prequals", "Mag Glass", "CarSys", "Gold and Black", "Black/Gold/Mag", "All"])
        self.filter_dropdown.currentIndexChanged.connect(self.filter_changed)
        search_layout.addWidget(self.filter_dropdown)

        # Suggestions list (hidden by default)
        self.suggestions_list = QListWidget()
        self.suggestions_list.setMaximumHeight(100)
        self.suggestions_list.hide()
        self.suggestions_list.itemClicked.connect(self.on_suggestion_clicked)
        self.suggestions_list.setStyleSheet("""
            QListWidget {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                background: #fff;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #f8f9fa;
            }
            QListWidget::item:hover {
                background: #f8f9fa;
            }
            QListWidget::item:selected {
                background: #20567C;
                color: white;
            }
        """)
        search_layout.addWidget(self.suggestions_list)

        main_layout.addWidget(search_card)

    def create_main_content(self, main_layout):
        """Create main content area with multi-toggle tab row and horizontal split display"""
        content_card = ModernCard()
        content_layout = QVBoxLayout(content_card)

        # Multi-toggle tab row
        self.tab_buttons = {}
        self.tab_panels = {}
        tab_info = [
            ("Prequals", "prequals_panel"),
            ("Blacklist", "blacklist_panel"),
            ("Goldlist", "goldlist_panel"),
            ("Mag Glass", "mag_glass_panel"),
            ("CarSys", "carsys_panel")
        ]
        tab_row = QHBoxLayout()
        for i, (label, attr) in enumerate(tab_info):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setObjectName(f"tab_{attr}")
            btn.setStyleSheet(self.get_tab_button_style(selected=(i==0)))
            btn.setChecked(i==0)  # Only Prequals on by default
            btn.toggled.connect(lambda checked, a=attr, b=btn: self.toggle_tab_panel(a, b, checked))
            tab_row.addWidget(btn)
            self.tab_buttons[attr] = btn
        tab_row.addStretch()
        content_layout.addLayout(tab_row)

        # Panels (horizontal stacking)
        self.prequals_panel = self.create_prequals_panel()
        self.blacklist_panel = self.create_blacklist_panel()
        self.goldlist_panel = self.create_goldlist_panel()
        self.mag_glass_panel = self.create_mag_glass_panel()
        self.carsys_panel = self.create_carsys_panel()
        self.tab_panels = {
            "prequals_panel": self.prequals_panel,
            "blacklist_panel": self.blacklist_panel,
            "goldlist_panel": self.goldlist_panel,
            "mag_glass_panel": self.mag_glass_panel,
            "carsys_panel": self.carsys_panel
        }
        self.panels_layout = QHBoxLayout()  # Ensure horizontal stacking
        for i, (label, attr) in enumerate(tab_info):
            panel = self.tab_panels[attr]
            panel.setVisible(i==0)  # Only Prequals visible by default
            self.panels_layout.addWidget(panel)
        content_layout.addLayout(self.panels_layout)
        main_layout.addWidget(content_card)

    def get_tab_button_style(self, selected=False):
        if selected:
            return (
                "QPushButton { background: #20567C; color: #fff; font-weight: bold; border: none; border-top-left-radius: 14px; border-top-right-radius: 14px; padding: 12px 32px; font-size: 16px; min-width: 120px; min-height: 36px; margin-right: 8px; }"
                "QPushButton:checked { background: #20567C; color: #fff; }"
            )
        else:
            return (
                "QPushButton { background: #fff; color: #20567C; font-weight: bold; border: none; border-top-left-radius: 14px; border-top-right-radius: 14px; padding: 12px 32px; font-size: 16px; min-width: 120px; min-height: 36px; margin-right: 8px; }"
                "QPushButton:checked { background: #20567C; color: #fff; }"
            )

    def toggle_tab_panel(self, attr, btn, checked):
        # Prevent toggling off the last visible tab
        if not checked:
            visible_count = sum(b.isChecked() for b in self.tab_buttons.values())
            if visible_count == 0:
                btn.setChecked(True)
                return
        panel = self.tab_panels[attr]
        panel.setVisible(checked)
        # Update tab style for all buttons
        for a, b in self.tab_buttons.items():
            b.setStyleSheet(self.get_tab_button_style(selected=b.isChecked()))
        # Update panel content when toggled on
        if checked:
            if attr == "prequals_panel":
                self.handle_prequal_search(self.make_dropdown.currentText(), self.model_dropdown.currentText(), self.year_dropdown.currentText())
            elif attr == "blacklist_panel":
                self.display_blacklist(self.make_dropdown.currentText())
            elif attr == "goldlist_panel":
                self.display_goldlist(self.make_dropdown.currentText())
            elif attr == "mag_glass_panel":
                self.display_mag_glass(self.make_dropdown.currentText())
            elif attr == "carsys_panel":
                self.display_carsys_data(self.make_dropdown.currentText())

    def create_prequals_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        header_layout = QHBoxLayout()
        title_label = QLabel("Prequalification Data")
        title_label.setObjectName("panel_title_label")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(lambda: self.pop_out_panel("Prequals", self.left_panel.toHtml()))
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        self.left_panel = ModernTextBrowser()
        self.left_panel.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.left_panel)
        return card

    def create_blacklist_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        header_layout = QHBoxLayout()
        title_label = QLabel("Blacklist Data")
        title_label.setObjectName("panel_title_label")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(lambda: self.pop_out_panel("Blacklist", self.blacklist_panel_widget.toHtml()))
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        self.blacklist_panel_widget = ModernTextBrowser()
        self.blacklist_panel_widget.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.blacklist_panel_widget)
        return card

    def create_goldlist_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        header_layout = QHBoxLayout()
        title_label = QLabel("Goldlist Data")
        title_label.setObjectName("panel_title_label")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(lambda: self.pop_out_panel("Goldlist", self.goldlist_panel_widget.toHtml()))
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        self.goldlist_panel_widget = ModernTextBrowser()
        self.goldlist_panel_widget.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.goldlist_panel_widget)
        return card

    def create_mag_glass_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        header_layout = QHBoxLayout()
        title_label = QLabel("Mag Glass Data")
        title_label.setObjectName("panel_title_label")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(lambda: self.pop_out_panel("Mag Glass", self.mag_glass_panel_widget.toHtml()))
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        self.mag_glass_panel_widget = ModernTextBrowser()
        self.mag_glass_panel_widget.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.mag_glass_panel_widget)
        return card

    def create_carsys_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        header_layout = QHBoxLayout()
        title_label = QLabel("CarSys Data")
        title_label.setObjectName("panel_title_label")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(lambda: self.pop_out_panel("CarSys", self.carsys_panel_widget.toHtml()))
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        self.carsys_panel_widget = ModernTextBrowser()
        self.carsys_panel_widget.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.carsys_panel_widget)
        return card

    def create_status_bar(self):
        """Create modern status bar"""
        self.status_bar = ModernStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Add progress bar
        self.progress_bar = ModernProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
        # Add transparency control
        self.status_bar.addPermanentWidget(QLabel("Transparency:"))
        self.opacity_slider = ModernSlider(Qt.Horizontal)
        self.opacity_slider.setMinimum(20)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.setFixedWidth(100)
        self.opacity_slider.valueChanged.connect(self.change_opacity)
        self.status_bar.addPermanentWidget(self.opacity_slider)

    def apply_modern_theme(self):
        """Apply modern light theme by default"""
        self.setStyleSheet("""
            QMainWindow {
                background: #f8f9fa;
            }
            QLabel {
                color: #495057;
            }
        """)

    def add_toolbar_button(self, text, slot, object_name):
        """Add a button to the toolbar"""
        button = ModernButton(text, style="secondary")
        button.clicked.connect(slot)
        button.setObjectName(object_name)
        self.toolbar.addWidget(button)
        return button

    def pop_out_panel(self, title, content):
        """Create a pop-out window for panel content"""
        popout = PopOutWindow(title, content, self)
        popout.show()

    def get_db_connection(self):
        """Get database connection"""
        return get_db_connection(self.db_path)

    def on_year_selected(self):
        selected_year = self.year_dropdown.currentText()
        selected_filter = self.filter_dropdown.currentText()
        current_make = self.make_dropdown.currentText()
        if selected_year != "Select Year":
            try:
                selected_year_int = int(selected_year)
                valid_prequal_data = []
                for item in self.data['prequal']:
                    try:
                        if (self.has_valid_prequal(item) and 'Year' in item and pd.notna(item['Year'])):
                            item_year = int(float(item['Year']))
                            if item_year == selected_year_int:
                                valid_prequal_data.append(item)
                    except (ValueError, TypeError, KeyError):
                        continue
                makes = []
                for item in valid_prequal_data:
                    try:
                        if 'Make' in item and item['Make'].strip() and item['Make'].strip().lower() != 'unknown':
                            makes.append(item['Make'].strip())
                    except (AttributeError, KeyError):
                        continue
                makes = sorted(set(makes))
                self.make_dropdown.clear()
                self.make_dropdown.addItem("Select Make")
                self.make_dropdown.addItem("All")
                self.make_dropdown.addItems(makes)
                if current_make in makes:
                    index = self.make_dropdown.findText(current_make)
                    if index != -1:
                        self.make_dropdown.setCurrentIndex(index)
                self.model_dropdown.clear()
                self.model_dropdown.addItem("Select Model")
                updated_make = self.make_dropdown.currentText()
                if updated_make not in ["Select Make", "All"]:
                    self.update_model_dropdown()
            except (ValueError, TypeError) as e:
                import logging
                logging.error(f"Error updating makes for year {selected_year}: {e}")
        if selected_filter in ["Prequals", "Gold and Black", "Blacklist", "Goldlist", "Mag Glass", "All"]:
            self.perform_search()

    def update_model_dropdown(self):
        selected_year = self.year_dropdown.currentText().strip()
        selected_make = self.make_dropdown.currentText().strip()
        self.populate_models(selected_year, selected_make)
        import logging
        logging.debug(f"Updated model dropdown for Year: {selected_year}, Make: {selected_make}")

    def handle_model_change(self, index):
        selected_model = self.model_dropdown.currentText().strip()
        import logging
        logging.debug(f"Model selected: {selected_model}")
        self.perform_search()

    def populate_models(self, year_text, make_text):
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")
        if year_text == "Select Year" or make_text == "Select Make" or make_text == "All":
            return
        try:
            year_int = int(year_text)
            matching_models = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Year' in item and 'Make' in item and 'Model' in item and pd.notna(item['Year']) and item['Make'] == make_text):
                        item_year = int(float(item['Year']))
                        if item_year == year_int:
                            matching_models.add(str(item['Model']))
                except (ValueError, TypeError, KeyError):
                    continue
            if matching_models:
                self.model_dropdown.addItems(sorted(matching_models))
                import logging
                logging.info(f"Added {len(matching_models)} models for Year: {year_int}, Make: {make_text}")
            else:
                import logging
                logging.warning(f"No models found for Year: {year_int}, Make: {make_text}")
        except (ValueError, TypeError) as e:
            import logging
            logging.error(f"Error in populate_models: {e}")

    def filter_changed(self):
        """Handle filter dropdown change"""
        selected_filter = self.filter_dropdown.currentText()
        if selected_filter != "Select List":
            self.update_panel_visibility(selected_filter, self.make_dropdown.currentText())
            self.log_action(self.current_user, f"Selected filter: {selected_filter}")

    def clear_filters(self):
        """Clear all filters"""
        self.year_dropdown.setCurrentIndex(0)
        self.make_dropdown.setCurrentIndex(0)
        self.model_dropdown.setCurrentIndex(0)
        self.filter_dropdown.setCurrentIndex(0)
        self.search_bar.clear()
        self.clear_display_panels()
        self.log_action(self.current_user, "Cleared all filters")

    def toggle_always_on_top(self, checked):
        """Toggle always on top behavior"""
        if checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
            self.show()
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
            self.show()

    def update_suggestions(self, text):
        """Update search suggestions"""
        if len(text) >= 2:
            suggestions = self.get_suggestions(text)
            if suggestions:
                self.suggestions_list.clear()
                for suggestion in suggestions[:10]:
                    self.suggestions_list.addItem(suggestion)
                self.suggestions_list.show()
            else:
                self.suggestions_list.hide()
        else:
            self.suggestions_list.hide()

    def get_suggestions(self, text):
        """Get search suggestions"""
        suggestions = []
        text_lower = text.lower()
        
        # Search in all data sources
        for data_type, data_list in self.data.items():
            for item in data_list:
                if isinstance(item, dict):
                    for key, value in item.items():
                        if isinstance(value, str) and text_lower in value.lower():
                            suggestions.append(f"{value} ({data_type})")
                elif isinstance(item, str) and text_lower in item.lower():
                    suggestions.append(f"{item} ({data_type})")
        
        return list(set(suggestions))

    def on_suggestion_clicked(self):
        """Handle suggestion selection"""
        item = self.suggestions_list.currentItem()
        if item:
            suggestion_text = item.text().split(" (")[0]
            self.search_bar.setText(suggestion_text)
            self.suggestions_list.hide()
            self.perform_search()

    def perform_search(self):
        """Perform search based on current filters"""
        dtc_code = self.search_bar.text().strip()
        selected_filter = self.filter_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        
        print(f"[DEBUG] Search criteria: Year={selected_year}, Make={selected_make}, Model={selected_model}, Filter={selected_filter}")
        
        # Print first 3 prequal records loaded
        if hasattr(self, 'data') and 'prequal' in self.data and self.data['prequal']:
            print(f"[DEBUG] First 3 prequal records loaded: {self.data['prequal'][:3]}")
        else:
            print("[DEBUG] No prequal data loaded.")
        
        # Trigger search if DTC code is entered OR if Year/Make/Model are selected
        if dtc_code or (selected_year != "Select Year" and selected_make != "Select Make" and selected_model != "Select Model"):
            self.update_displays_based_on_filter(selected_filter, dtc_code, selected_make, selected_model, selected_year)
            if dtc_code:
                self.log_action(self.current_user, f"Search performed: {dtc_code}")
            else:
                self.log_action(self.current_user, f"Search performed: {selected_year} {selected_make} {selected_model}")

    def clear_display_panels(self):
        """Clear all display panels"""
        self.left_panel.clear()
        self.blacklist_panel_widget.clear()
        self.goldlist_panel_widget.clear()
        self.mag_glass_panel_widget.clear()
        self.carsys_panel_widget.clear()

    def update_panel_visibility(self, selected_filter, selected_make):
        """Update which panels are visible based on filter"""
        # Show/hide tabs based on filter selection
        if selected_filter == "Prequals":
            self.prequals_panel.setVisible(True)
            self.blacklist_panel.setVisible(False)
            self.goldlist_panel.setVisible(False)
            self.mag_glass_panel.setVisible(False)
            self.carsys_panel.setVisible(False)
        elif selected_filter == "Blacklist":
            self.prequals_panel.setVisible(False)
            self.blacklist_panel.setVisible(True)
            self.goldlist_panel.setVisible(False)
            self.mag_glass_panel.setVisible(False)
            self.carsys_panel.setVisible(False)
        elif selected_filter == "Goldlist":
            self.prequals_panel.setVisible(False)
            self.blacklist_panel.setVisible(False)
            self.goldlist_panel.setVisible(True)
            self.mag_glass_panel.setVisible(False)
            self.carsys_panel.setVisible(False)
        elif selected_filter == "Mag Glass":
            self.prequals_panel.setVisible(False)
            self.blacklist_panel.setVisible(False)
            self.goldlist_panel.setVisible(False)
            self.mag_glass_panel.setVisible(True)
            self.carsys_panel.setVisible(False)
        elif selected_filter == "CarSys":
            self.prequals_panel.setVisible(False)
            self.blacklist_panel.setVisible(False)
            self.goldlist_panel.setVisible(False)
            self.mag_glass_panel.setVisible(False)
            self.carsys_panel.setVisible(True)
        elif selected_filter == "Gold and Black":
            # Show blacklist tab by default for combined view
            self.prequals_panel.setVisible(False)
            self.blacklist_panel.setVisible(True)
            self.goldlist_panel.setVisible(True)
            self.mag_glass_panel.setVisible(False)
            self.carsys_panel.setVisible(False)

    def update_displays_based_on_filter(self, selected_filter, dtc_code, selected_make, selected_model, selected_year):
        """Update displays based on selected filter"""
        if selected_filter == "Prequals" or selected_filter == "Select List":
            # If "Select List" is chosen, default to Prequals
            self.handle_prequal_search(selected_make, selected_model, selected_year)
        elif selected_filter == "Mag Glass":
            self.search_mag_glass(selected_make)
        elif selected_filter == "CarSys":
            self.search_carsys_dtc(dtc_code)
        elif selected_filter == "Blacklist":
            self.display_blacklist(selected_make)
        elif selected_filter == "Goldlist":
            self.display_goldlist(selected_make)
        elif selected_filter == "Gold and Black":
            # Display both blacklist and goldlist
            self.display_blacklist(selected_make)
            self.display_goldlist(selected_make)
        elif selected_filter == "All":
            # Show all data
            self.handle_prequal_search(selected_make, selected_model, selected_year)
            self.search_mag_glass(selected_make)
            self.search_carsys_dtc(dtc_code)
            self.display_blacklist(selected_make)
            self.display_goldlist(selected_make)

    def handle_prequal_search(self, selected_make, selected_model, selected_year):
        """Handle prequalification search"""
        # Dictionary to hold unique System Acronyms
        unique_results = {}

        # Convert selected_model to a string
        selected_model_str = str(selected_model)

        # Filtering data based on selections
        filtered_results = [item for item in self.data['prequal']
                            if (selected_make == "All" or item['Make'] == selected_make) and
                            (selected_model == "Select Model" or str(item['Model']) == selected_model_str) and
                            (selected_year == "Select Year" or str(item['Year']) == selected_year)]

        print(f"[DEBUG] Number of prequal results found: {len(filtered_results)}")

        # Populating the dictionary with unique entries based on System Acronym
        for item in filtered_results:
            system_acronym = item.get('Protech Generic System Name.1', 'N/A')  # Match original app column name
            if system_acronym not in unique_results:
                unique_results[system_acronym] = item  # Store the entire item for display

        # Now, pass the unique items to be displayed
        self.display_results(list(unique_results.values()), context='prequal')

    def display_carsys_data(self, selected_make):
        """Display CarSys data"""
        try:
            conn = self.get_db_connection()
            
            # Check if table exists
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='carsys'")
            if not cursor.fetchone():
                self.carsys_panel_widget.setPlainText("No CarSys data found. Please load data first.")
                return
            
            # If no make is selected, display all results
            if selected_make == "Select Make" or selected_make == "All":
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                """
                df = pd.read_sql_query(query, conn)
            else:
                # If a make is selected, filter the results by the selected make (case/whitespace-insensitive)
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                WHERE TRIM(UPPER(carMake)) = TRIM(UPPER(?))
                """
                df = pd.read_sql_query(query, conn, params=(selected_make,))
            
            # Replace NaN values with empty strings
            df.fillna("", inplace=True)
            
            # Display the results in the CarSys panel
            if df.empty:
                self.carsys_panel_widget.setPlainText(f"No CarSys results found for make: {selected_make}")
            else:
                # Rename columns for display to match the original Excel column names
                df = df.rename(columns={
                    'genericSystemName': 'Generic System Name',
                    'dtcSys': 'DTCsys',
                    'carMake': 'CarMake',
                    'comments': 'Comments'
                })
                
                # Add styling to the HTML table
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                styled_html = f"""
                <html>
                <head>
                <style>
                .table {{
                    border-collapse: collapse;
                    width: 100%;
                    font-family: Arial, sans-serif;
                    font-size: 12px;
                }}
                .table th {{
                    background-color: #667eea;
                    color: white;
                    padding: 8px;
                    text-align: left;
                    border: 1px solid #ddd;
                }}
                .table td {{
                    padding: 6px;
                    border: 1px solid #ddd;
                }}
                .table-striped tr:nth-child(even) {{
                    background-color: #f8f9fa;
                }}
                </style>
                </head>
                <body>
                {html_table}
                </body>
                </html>
                """
                self.carsys_panel_widget.setHtml(styled_html)
                self.carsys_panel_widget.setOpenExternalLinks(True)
        
        except Exception as e:
            logging.error(f"Failed to display CarSys data: {e}")
            self.carsys_panel_widget.setPlainText(f"An error occurred while fetching the CarSys data: {str(e)}")
        
        finally:
            conn.close()

    def display_mag_glass(self, selected_make):
        """Display Mag Glass data"""
        try:
            conn = self.get_db_connection()
            
            # Check if table exists
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mag_glass'")
            if not cursor.fetchone():
                self.mag_glass_panel_widget.setPlainText("No Mag Glass data found. Please load data first.")
                return
            
            # Prepare the query based on the selected make - use correct quoted column names
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
                # Add styling to the HTML table
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                styled_html = f"""
                <html>
                <head>
                <style>
                .table {{
                    border-collapse: collapse;
                    width: 100%;
                    font-family: Arial, sans-serif;
                    font-size: 12px;
                }}
                .table th {{
                    background-color: #667eea;
                    color: white;
                    padding: 8px;
                    text-align: left;
                    border: 1px solid #ddd;
                }}
                .table td {{
                    padding: 6px;
                    border: 1px solid #ddd;
                }}
                .table-striped tr:nth-child(even) {{
                    background-color: #f8f9fa;
                }}
                </style>
                </head>
                <body>
                {html_table}
                </body>
                </html>
                """
                self.mag_glass_panel_widget.setHtml(styled_html)
                self.mag_glass_panel_widget.setOpenExternalLinks(True)
            else:
                self.mag_glass_panel_widget.setPlainText(f"No Mag Glass results found for make: {selected_make}")
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.mag_glass_panel_widget.setPlainText(f"An error occurred while fetching the data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def search_carsys_dtc(self, dtc_code):
        """Search CarSys DTC codes"""
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
                self.carsys_panel_widget.setPlainText("No results found for the given criteria.")
            else:
                # Rename columns for display to match the original Excel column names
                df = df.rename(columns={
                    'genericSystemName': 'Generic System Name',
                    'dtcSys': 'DTCsys',
                    'carMake': 'CarMake',
                    'comments': 'Comments'
                })
                
                self.carsys_panel_widget.setHtml(df.to_html(index=False, escape=False))
                self.carsys_panel_widget.setOpenExternalLinks(True)
        
        except Exception as e:
            logging.error(f"Failed to execute CarSys search query: {query}\nError: {e}")
            self.carsys_panel_widget.setPlainText("An error occurred while fetching the CarSys data.")
        
        finally:
            conn.close()

    def search_mag_glass(self, selected_make):
        """Search Mag Glass data"""
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
            self.mag_glass_panel_widget.setPlainText("An error occurred while fetching the data.")
            return

        if not df.empty:
            self.mag_glass_panel_widget.setHtml(df.to_html(index=False, escape=False))
            self.mag_glass_panel_widget.setOpenExternalLinks(True)
        else:
            self.mag_glass_panel_widget.setPlainText("No Mag Glass results found.")

    def search_dtc_codes(self, dtc_code, filter_type, selected_make):
        """Search DTC codes"""
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
            # Determine which panel to use based on filter type
            if filter_type == "Blacklist":
                self.blacklist_panel_widget.setHtml(df.to_html(index=False, escape=False))
                self.blacklist_panel_widget.setOpenExternalLinks(True)
            elif filter_type == "Goldlist":
                self.goldlist_panel_widget.setHtml(df.to_html(index=False, escape=False))
                self.goldlist_panel_widget.setOpenExternalLinks(True)
            else:
                # For "All" or "Gold and Black", show in both panels
                self.blacklist_panel_widget.setHtml(df.to_html(index=False, escape=False))
                self.blacklist_panel_widget.setOpenExternalLinks(True)
                self.goldlist_panel_widget.setHtml(df.to_html(index=False, escape=False))
                self.goldlist_panel_widget.setOpenExternalLinks(True)
        else:
            if filter_type == "Blacklist":
                self.blacklist_panel_widget.setPlainText("No DTC code results found.")
            elif filter_type == "Goldlist":
                self.goldlist_panel_widget.setPlainText("No DTC code results found.")
            else:
                self.blacklist_panel_widget.setPlainText("No DTC code results found.")
                self.goldlist_panel_widget.setPlainText("No DTC code results found.")



    def display_results(self, results, context='prequal'):
        """Display search results in markdown format"""
        if results:
            print(f"[DEBUG] Displaying {len(results)} prequal results.")
            print(f"[DEBUG] Keys of first result: {list(results[0].keys())}")
        else:
            print("[DEBUG] No results to display in prequal box.")
        
        # Convert to markdown format
        markdown_text = ""
        
        for result in results:
            # Check if any relevant field is "N/A" or None (match original app logic)
            if any(
                result.get(key) in [None, 'N/A'] 
                for key in ['Make', 'Model', 'Year', 'Calibration Type', 'Protech Generic System Name.1', 'Parts Code Table Value', 'Calibration Pre-Requisites']
            ):
                continue  # Skip this result
            
            # Get all available fields from the result (match original app)
            year = result.get('Year', 'N/A')
            if isinstance(year, float):
                year = str(int(year))
            
            make = result.get('Make', 'N/A')
            model = result.get('Model', 'N/A')
            system_acronym = result.get('Protech Generic System Name.1', 'N/A')
            parts_code = result.get('Parts Code Table Value', 'N/A')
            calibration_type = result.get('Calibration Type', 'N/A')
            og_calibration_type = result.get('OG Calibration Type', 'N/A')
            calibration_prerequisites = result.get('Calibration Pre-Requisites', 'N/A')
            
            # Handle link safely (match original app)
            service_link = result.get('Service Information Hyperlink', '#')
            if isinstance(service_link, float) and pd.isna(service_link):
                service_link = '#'
            elif isinstance(service_link, str) and not service_link.startswith(('http://', 'https://')):
                service_link = 'http://' + service_link
            
            # Create markdown entry (match original app format with clickable link)
            markdown_text += f"""
**Make:** {make}  
**Model:** {model}  
**Year:** {year}  
**System Acronym:** {system_acronym}  
**Parts Code Table Value:** {parts_code}  
**Calibration Type:** {calibration_type}  
**OG Calibration Type:** {og_calibration_type}  
**Service Information:** <a href='{service_link}'>Click Here</a>  

**Pre-Quals:** {calibration_prerequisites}  

---
"""
        
        if not markdown_text:
            markdown_text = "*No prequalification data found for the selected criteria.*"
        
        # Convert markdown to HTML for display
        html_text = self.convert_markdown_to_html(markdown_text)
        self.left_panel.setHtml(html_text)
        self.left_panel.setOpenExternalLinks(True)

    def convert_markdown_to_html(self, markdown_text):
        """Convert markdown text to HTML for display with smaller font size"""
        html = markdown_text
        
        # Convert markdown headers to HTML
        html = html.replace('## ', '<h2 style="color: #20567C; margin-top: 15px; margin-bottom: 8px; border-bottom: 2px solid #20567C; padding-bottom: 3px; font-size: 12px;">')
        html = html.replace('\n\n', '</h2>\n')
        
        # Convert bold text
        html = html.replace('**', '<strong style="color: #20567C; font-size: 10px;">')
        html = html.replace('**', '</strong>')
        
        # Convert italic text
        html = html.replace('*', '<em style="font-size: 10px;">')
        html = html.replace('*', '</em>')
        
        # Convert horizontal rules
        html = html.replace('---', '<hr style="border: 1px solid #e0e0e0; margin: 12px 0;">')
        
        # Convert line breaks
        html = html.replace('\n', '<br>')
        
        # Add styling wrapper with smaller font size
        html = f"""
        <div style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 9px; line-height: 1.4; color: #333;">
        {html}
        </div>
        """
        
        return html

    def clear_search_bar(self):
        """Clear search bar"""
        self.search_bar.clear()
        self.suggestions_list.hide()

    def change_opacity(self, value):
        """Change window opacity"""
        self.setWindowOpacity(value / 100.0)

    def open_admin(self):
        """Open admin panel"""
        dialog = ManageDataListsDialog(self)
        dialog.exec_()
        
    def manage_paths(self):
        """Create a dialog to manage paths"""
        self.path_dialog = QDialog(self)
        self.path_dialog.setWindowTitle("Manage Lists")
        self.path_dialog.setGeometry(100, 100, 500, 350)
        
        # Create a title label
        title_label = QLabel("Manage Data Lists")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        
        # Add a note about CarSYS and MagGlass using goldlist path
        note_label = QLabel("Note: CarSYS and MagGlass automatically use the same path as Goldlist")
        note_label.setAlignment(Qt.AlignCenter)
        note_label.setStyleSheet("font-style: italic; color: #888;")
        
        # Create a form layout to hold the path selectors
        form_layout = QFormLayout()
        
        self.load_buttons = {}  # Store load buttons for easy access
        self.clear_buttons = {}  # Store clear buttons for easy access
        self.path_labels = {}  # Store labels for easy access
        self.browse_buttons = {}  # Store browse buttons for easy access
        
        # Define new display names - remove CarSys and mag_glass
        config_display_names = {
            "blacklist": "Blacklist",
            "goldlist": "Goldlist",
            "prequal": "Prequals (Longsheets)"
        }
        
        # Create a function to handle load button clicks
        def on_load_button_clicked(config_type):
            folder_path = QFileDialog.getExistingDirectory(self, f"Select {config_display_names[config_type]} Directory")
            if folder_path:
                self.browse_buttons[config_type].setText(folder_path)
        
        # Create a function to handle clear button clicks
        def on_clear_button_clicked(config_type):
            msg_box = self.create_styled_messagebox(
                "Clear Data", 
                f"Are you sure you want to clear all {config_display_names[config_type]} data?",
                QMessageBox.Question
            )
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.No)
            
            if msg_box.exec_() == QMessageBox.Yes:
                self.clear_data(config_type)
                self.browse_buttons[config_type].clear()
                
                # Special case for goldlist to also clear CarSYS and MagGlass
                if config_type == "goldlist":
                    self.clear_data("CarSys")
                    self.clear_data("mag_glass")
                    success_msg = self.create_styled_messagebox(
                        "Data Cleared", 
                        f"{config_display_names[config_type]}, CarSYS, and MagGlass data has been cleared.",
                        QMessageBox.Information
                    )
                else:
                    success_msg = self.create_styled_messagebox(
                        "Data Cleared", 
                        f"{config_display_names[config_type]} data has been cleared.",
                        QMessageBox.Information
                    )
                success_msg.exec_()
        
        for config_name, display_name in config_display_names.items():
            # Create a horizontal layout for the buttons
            buttons_layout = QHBoxLayout()
            
            # Create a styled section label
            section_label = QLabel(display_name)
            section_label.setStyleSheet("font-weight: bold; font-size: 12px; margin-top: 10px;")
            
            load_button = QPushButton("Choose Path")
            load_button.setObjectName("action_button")
            load_button.clicked.connect(lambda checked, config=config_name: on_load_button_clicked(config))
            self.load_buttons[config_name] = load_button
            buttons_layout.addWidget(load_button)
            
            clear_button = QPushButton("Clear Data")
            clear_button.setObjectName("danger_button")
            clear_button.clicked.connect(lambda checked, config=config_name: on_clear_button_clicked(config))
            self.clear_buttons[config_name] = clear_button
            buttons_layout.addWidget(clear_button)
            
            # Create a widget to hold the buttons
            buttons_widget = QWidget()
            buttons_widget.setLayout(buttons_layout)
            
            form_layout.addRow(section_label, buttons_widget)
            self.path_labels[config_name] = QLabel(f"Path:")
            self.browse_buttons[config_name] = QLineEdit()
            self.browse_buttons[config_name].setReadOnly(True)
            self.browse_buttons[config_name].setPlaceholderText(f"Select a directory for {display_name}")
            form_layout.addRow(self.path_labels[config_name], self.browse_buttons[config_name])
            self.browse_buttons[config_name].textChanged.connect(self.update_path_label)
        
        button_layout = QHBoxLayout()
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setObjectName("cancel_button")
        self.cancel_button.clicked.connect(self.path_dialog.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.clear_all_button = QPushButton("Clear All Data")
        self.clear_all_button.setObjectName("clear_all_button")
        self.clear_all_button.clicked.connect(self.clear_all_data)
        button_layout.addWidget(self.clear_all_button)
        
        self.save_button = QPushButton("Save & Load Data")
        self.save_button.setObjectName("primary_button")
        self.save_button.clicked.connect(self.save_paths)
        button_layout.addWidget(self.save_button)
        
        # Setup main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(title_label)
        main_layout.addWidget(note_label)
        main_layout.addSpacing(10)
        
        # Add a container widget for the form
        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        main_layout.addWidget(form_widget)
        
        # Add some space before buttons
        main_layout.addSpacing(10)
        main_layout.addLayout(button_layout)
        
        self.path_dialog.setLayout(main_layout)
        
        # Pre-populate path fields with existing paths
        for config_name in config_display_names.keys():
            existing_path = load_path_from_db(config_name, self.db_path)
            if existing_path:
                self.browse_buttons[config_name].setText(existing_path)
        
        self.path_dialog.exec_()
        
    def save_paths(self):
        """Save paths and load data"""
        paths_to_save = {}
        
        for config_type, browse_button in self.browse_buttons.items():
            folder_path = browse_button.text()
            if folder_path:
                paths_to_save[config_type] = folder_path
        
        if not paths_to_save:
            QMessageBox.warning(self, "No Paths", "No paths have been selected.")
            return
        
        # Get the goldlist path to share with CarSYS and MagGlass
        goldlist_path = paths_to_save.get("goldlist")
        
        for config_type, folder_path in paths_to_save.items():
            try:
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    QMessageBox.warning(self, "Load Error", f"No valid Excel files found in the {folder_path} directory for {config_type}.")
                    continue
                
                # Clear existing data
                self.clear_data(config_type)
                
                # Special case: if goldlist is being updated, also clear CarSYS and MagGlass
                if config_type == "goldlist":
                    self.clear_data("CarSys")
                    self.clear_data("mag_glass")
                
                if hasattr(self, 'progress_bar'):
                    self.progress_bar.setVisible(True)
                    self.progress_bar.setMaximum(len(files))
                    self.progress_bar.setValue(0)
                
                # Process files based on configuration type
                data_loaded = False
                for i, (filename, filepath) in enumerate(files.items()):
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=1)
                            if result == "Data loaded successfully":
                                data_loaded = True
                                
                                # If this is goldlist, also load the same file for CarSYS and MagGlass
                                if config_type == 'goldlist':
                                    load_carsys_data_to_db(filepath, table_name='CarSys', db_path=self.db_path)
                                    load_mag_glass_data(filepath, table_name='mag_glass', db_path=self.db_path)
                        elif config_type == 'prequal':
                            df = pd.read_excel(filepath)
                            if df.empty:
                                logging.warning(f"{filename} is empty.")
                                continue
                            data = df.to_dict(orient='records')
                            update_configuration(config_type, folder_path, data, self.db_path)
                            data_loaded = True
                        elif config_type == 'mag_glass':
                            result = load_mag_glass_data(filepath, config_type, db_path=self.db_path)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'CarSys':
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=0)
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
                    
                    if hasattr(self, 'progress_bar'):
                        self.progress_bar.setValue(i + 1)
                
                # Save path to database
                save_path_to_db(config_type, folder_path, self.db_path)
                
                # If this is goldlist, also save the same path for CarSYS and MagGlass
                if config_type == 'goldlist':
                    save_path_to_db('CarSys', folder_path, self.db_path)
                    save_path_to_db('mag_glass', folder_path, self.db_path)
            
            except Exception as e:
                logging.error(f"Error saving path for {config_type}: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save path for {config_type}: {str(e)}")
        
        if hasattr(self, 'progress_bar'):
            self.progress_bar.setVisible(False)
        self.load_configurations()  # Reload all configurations
        self.populate_dropdowns()  # Repopulate dropdowns
        self.check_data_loaded()  # Check if data is loaded
        
        # Show a single success message for the entire operation
        msg = self.create_styled_messagebox("Success", "Paths saved and data loaded successfully!", QMessageBox.Information)
        msg.exec_()
        self.path_dialog.accept()
        
    def update_path_label(self):
        """Update path labels"""
        for config_name, label in self.path_labels.items():
            label.setText(f"{config_name} Path: {self.browse_buttons[config_name].text()}")
            
    def clear_all_data(self):
        """Clear all data with confirmation"""
        msg_box = self.create_styled_messagebox(
            "Clear All Data", 
            "Are you sure you want to clear ALL data?\nThis action cannot be undone.",
            QMessageBox.Question
        )
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.No)
        
        confirmation = msg_box.exec_()
        
        if confirmation == QMessageBox.Yes:
            # Clear all data types
            self.clear_data()
            
            # Clear all path text fields
            for path_field in self.browse_buttons.values():
                path_field.clear()
                
            # Update UI
            self.load_configurations()
            self.populate_dropdowns()
            self.check_data_loaded()
            
            # Show success message
            success_msg = self.create_styled_messagebox(
                "Data Cleared", 
                "All data has been successfully cleared.",
                QMessageBox.Information
            )
            success_msg.exec_()
            
            # Log the action
            self.log_action(self.current_user, "Cleared all data from the database")
            
    def export_data(self):
        """Export data functionality"""
        try:
            # Create export dialog
            dialog = QDialog(self)
            dialog.setWindowTitle("Export Data")
            dialog.setGeometry(200, 200, 400, 300)
            
            layout = QVBoxLayout()
            
            # Title
            title_label = QLabel("Export Data")
            title_label.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px;")
            layout.addWidget(title_label)
            
            # Export options
            options_layout = QVBoxLayout()
            
            # CSV Export
            csv_button = QPushButton("Export to CSV")
            csv_button.clicked.connect(lambda: self.export_to_csv())
            options_layout.addWidget(csv_button)
            
            # JSON Export
            json_button = QPushButton("Export to JSON")
            json_button.clicked.connect(lambda: self.export_to_json())
            options_layout.addWidget(json_button)
            
            layout.addLayout(options_layout)
            
            # Close button
            close_button = QPushButton("Close")
            close_button.clicked.connect(dialog.accept)
            layout.addWidget(close_button)
            
            dialog.setLayout(layout)
            dialog.exec_()
            
        except Exception as e:
            logging.error(f"Error in export dialog: {e}")
            QMessageBox.critical(self, "Error", f"Failed to open export dialog: {str(e)}")
            
    def export_to_csv(self):
        """Export data to CSV"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save CSV File", "", "CSV Files (*.csv)"
            )
            if file_path:
                # Export all data types
                all_data = {}
                for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'carsys']:
                    if self.data.get(config_type):
                        all_data[config_type] = self.data[config_type]
                
                # Convert to CSV format
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    for data_type, data_list in all_data.items():
                        if data_list:
                            # Write header for each data type
                            csvfile.write(f"\n=== {data_type.upper()} ===\n")
                            if data_list:
                                fieldnames = data_list[0].keys()
                                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                                writer.writeheader()
                                writer.writerows(data_list)
                
                QMessageBox.information(self, "Success", f"Data exported to {file_path}")
                self.log_action(self.current_user, f"Exported data to CSV: {file_path}")
                
        except Exception as e:
            logging.error(f"Error exporting to CSV: {e}")
            QMessageBox.critical(self, "Error", f"Failed to export to CSV: {str(e)}")
            
    def export_to_json(self):
        """Export data to JSON"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save JSON File", "", "JSON Files (*.json)"
            )
            if file_path:
                # Export all data
                export_data = {
                    'export_date': datetime.now().isoformat(),
                    'data': self.data
                }
                
                with open(file_path, 'w', encoding='utf-8') as jsonfile:
                    json.dump(export_data, jsonfile, indent=2, ensure_ascii=False)
                
                QMessageBox.information(self, "Success", f"Data exported to {file_path}")
                self.log_action(self.current_user, f"Exported data to JSON: {file_path}")
                
        except Exception as e:
            logging.error(f"Error exporting to JSON: {e}")
            QMessageBox.critical(self, "Error", f"Failed to export to JSON: {str(e)}")

    def refresh_lists(self):
        """Refresh data lists"""
        # Implementation would be copied from original
        pass

    def load_configurations(self):
        """Load configurations from database"""
        logging.debug("Loading configurations...")
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'carsys']:
            data = load_configuration(config_type, self.db_path)
            self.data[config_type] = data if data else []
            logging.debug(f"Loaded {len(data)} items for {config_type}")
        if 'prequal' in self.data:
            self.populate_dropdowns()

    def check_data_loaded(self):
        """Check if data is loaded"""
        # Implementation would be copied from original
        pass

    def apply_selected_theme(self):
        theme = self.theme_dropdown.currentText()
        palette = get_theme_palette(theme)
        self.current_theme = theme
        self.save_settings({"theme": self.current_theme})

        # Main stylesheet
        style = f"""
        QMainWindow {{
            background: {palette['bg']};
        }}
        QWidget#ModernCard {{
            background-color: {palette['card']};
            border-radius: 12px;
            border: 1px solid {'#23262f' if theme == 'Dark' else '#e0e0e0'};
        }}
        QLabel {{
            color: {palette['text']};
        }}
        QPushButton {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {palette['accent']}, stop:1 {palette['accent2']});
            border: none;
            border-radius: 8px;
            color: {palette['button_fg']};
            font-weight: 600;
            padding: 12px 24px;
            font-size: 14px;
        }}
        QPushButton[style='secondary'] {{
            background: {palette['card']};
            color: {palette['text']};
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
        }}
        QPushButton[style='danger'] {{
            background: {palette['danger']};
            color: #fff;
        }}
        QPushButton[style='primary'] {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {palette['accent']}, stop:1 {palette['accent2']});
            color: #fff;
        }}
        QComboBox {{
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-radius: 8px;
            padding: 8px 12px;
            background: {palette['card']};
            color: {palette['text']};
            font-size: 14px;
        }}
        QComboBox QAbstractItemView {{
            background: {palette['card']};
            color: {palette['text']};
            selection-background-color: {palette['accent']};
            selection-color: #fff;
        }}
        QLineEdit {{
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-radius: 8px;
            padding: 10px 12px;
            background: {palette['card']};
            color: {palette['text']};
            font-size: 14px;
        }}
        QTextBrowser {{
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-radius: 8px;
            background: {palette['card']};
            color: {palette['text']};
            font-size: 13px;
        }}
        QTabWidget::pane {{
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-radius: 8px;
            background: {palette['card']};
        }}
        QTabBar::tab {{
            background: {palette['bg']};
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 8px 16px;
            margin-right: 2px;
            font-weight: 600;
            color: {palette['text']};
        }}
        QTabBar::tab:selected {{
            background: {palette['card']};
            border-color: {palette['accent']};
            color: {palette['accent']};
        }}
        QTabBar::tab:hover:!selected {{
            background: {'#23262f' if theme == 'Dark' else '#e9ecef'};
        }}
        QStatusBar {{
            background: {palette['card']};
            border-top: 1px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            color: {palette['text']};
            font-size: 12px;
        }}
        QToolBar {{
            background: {palette['card']};
            border-bottom: 1px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            spacing: 8px;
            padding: 8px;
        }}
        QProgressBar {{
            border: 2px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            border-radius: 8px;
            text-align: center;
            background: {palette['bg']};
            font-weight: 600;
        }}
        QProgressBar::chunk {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {palette['accent']}, stop:1 {palette['accent2']});
            border-radius: 6px;
        }}
        QSlider::groove:horizontal {{
            border: 1px solid {'#23262f' if theme == 'Dark' else '#e9ecef'};
            height: 8px;
            background: {palette['bg']};
            border-radius: 4px;
        }}
        QSlider::handle:horizontal {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {palette['accent']}, stop:1 {palette['accent2']});
            border: 2px solid {palette['accent']};
            width: 18px;
            margin: -5px 0;
            border-radius: 9px;
        }}
        QSlider::handle:horizontal:hover {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {palette['accent2']}, stop:1 {palette['accent']});
        }}
        """
        self.setStyleSheet(style)

        # Update user label and status color
        if hasattr(self, 'user_label'):
            if theme == "Dark":
                self.user_label.setStyleSheet("font-size: 15px; font-weight: 600; color: #fff; font-family: 'Segoe UI', Arial, sans-serif;")
            else:
                self.user_label.setStyleSheet("font-size: 15px; font-weight: 600; color: #20567C; font-family: 'Segoe UI', Arial, sans-serif;")
        if hasattr(self, 'user_status'):
            self.user_status.setStyleSheet(f"font-size: 12px; color: {palette['success']};")

        # Set title and subtitle label colors for dark/light theme
        if hasattr(self, 'findChild'):
            brand_label = self.findChild(QLabel, "brand_label")
            subtitle_label = self.findChild(QLabel, "subtitle_label")
            if theme == "Dark":
                if brand_label:
                    brand_label.setStyleSheet("color: #fff; font-size: 28px; font-weight: bold;")
                if subtitle_label:
                    subtitle_label.setStyleSheet("color: #b0b8c1; font-size: 14px;")
            else:
                if brand_label:
                    brand_label.setStyleSheet("color: #20567C; font-size: 28px; font-weight: bold;")
                if subtitle_label:
                    subtitle_label.setStyleSheet("color: #3a4a5a; font-size: 14px;")

    def apply_saved_theme(self):
        saved_settings = getattr(self, 'load_settings', lambda: {"theme": "Light"})()
        theme = saved_settings.get("theme", "Light")
        index = self.theme_dropdown.findText(theme)
        if index != -1:
            self.theme_dropdown.setCurrentIndex(index)
        self.apply_selected_theme()

    def closeEvent(self, event):
        """Handle close event"""
        self.log_action(self.current_user, "Application closed")
        event.accept()

    def keyPressEvent(self, event):
        """Handle key press events"""
        if event.key() == Qt.Key_F1 and event.modifiers() == (Qt.ControlModifier | Qt.ShiftModifier):
            dialog = AdminOptionsDialog(self)
            dialog.exec_()
        else:
            super().keyPressEvent(event)

    # Add to ModernAnalyzerApp:
    def save_settings(self, settings):
        import json
        with open(self.settings_file, 'w') as file:
            json.dump(settings, file)

    def load_settings(self):
        import os, json
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r') as file:
                return json.load(file)
        return {"theme": "Light"}

    def get_valid_excel_files(self, folder_path):
        import re, os, logging
        file_pattern = re.compile(r'(.+).xlsx$', re.IGNORECASE)
        skip_pattern = re.compile(r'.*X\.X\.xlsx$', re.IGNORECASE)
        valid_files = {}
        for file_name in os.listdir(folder_path):
            if file_pattern.match(file_name) and not skip_pattern.match(file_name):
                full_path = os.path.join(folder_path, file_name)
                valid_files[file_name] = full_path
        if not valid_files:
            logging.error("No valid Excel files found.")
        return valid_files

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
        except Exception as e:
            logging.error(f"Failed to clear data: {e}")
            QMessageBox.critical(self, "Error", "Failed to clear database.")
        self.load_configurations()

    def refresh_lists(self):
        self.log_action(self.current_user, "Clicked Refresh Lists button")
        any_data_loaded = False
        last_processed_path = ""
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'CarSys']:
            folder_path = load_path_from_db(config_type, self.db_path)
            if not folder_path and config_type in ['mag_glass', 'CarSys']:
                folder_path = load_path_from_db('goldlist', self.db_path)
                if folder_path:
                    save_path_to_db(config_type, folder_path, self.db_path)
            if folder_path:
                last_processed_path = folder_path
                self.clear_data(config_type)
                import logging
                logging.info(f"Cleared existing data for {config_type}")
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    if config_type != 'CarSys':
                        QMessageBox.warning(self, "Load Error", f"No valid Excel files found in the directory for {config_type}.")
                    continue
                data_loaded = False
                if hasattr(self, 'progress_bar'):
                    self.progress_bar.setVisible(True)
                    self.progress_bar.setMaximum(len(files))
                    self.progress_bar.setValue(0)
                for i, (filename, filepath) in enumerate(files.items()):
                    try:
                        if config_type in ['blacklist', 'goldlist']:
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=1)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'mag_glass':
                            result = load_mag_glass_data(filepath, config_type, db_path=self.db_path)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        elif config_type == 'CarSys':
                            result = load_excel_data_to_db(filepath, config_type, db_path=self.db_path, sheet_index=0)
                            if result == "Data loaded successfully":
                                data_loaded = True
                        else:
                            import pandas as pd
                            df = pd.read_excel(filepath)
                            if df.empty:
                                import logging
                                logging.warning(f"{filename} is empty.")
                                continue
                            data = df.to_dict(orient='records')
                            update_configuration(config_type, folder_path, data, self.db_path)
                            data_loaded = True
                    except Exception as e:
                        import logging
                        logging.error(f"Error loading {filename}: {str(e)}")
                        if config_type != 'CarSys':
                            QMessageBox.critical(self, "Load Error", f"Failed to load {filename}: {str(e)}")
                    if hasattr(self, 'progress_bar'):
                        self.progress_bar.setValue(i + 1)
                if hasattr(self, 'progress_bar'):
                    self.progress_bar.setVisible(False)
                if data_loaded:
                    any_data_loaded = True
                    self.load_configurations()
                    self.populate_dropdowns()
                    self.check_data_loaded()
                    if hasattr(self, 'status_bar'):
                        self.status_bar.showMessage(f"Data refreshed from: {folder_path}")
            else:
                import logging
                logging.warning(f"No saved path found for {config_type}")
        if any_data_loaded:
            msg = self.create_styled_messagebox("Success", "All data refreshed successfully!", QMessageBox.Information)
            msg.exec_()
        # Always repopulate dropdowns and check enabled state after refresh
        self.populate_dropdowns()
        self.check_data_loaded()

    def load_configurations(self):
        import logging
        logging.debug("Loading configurations...")
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'carsys']:
            data = load_configuration(config_type, self.db_path)
            self.data[config_type] = data if data else []
            logging.debug(f"Loaded {len(data)} items for {config_type}")
        # Debug: print first few prequal items and their types
        prequal_sample = self.data['prequal'][:3]
        logging.debug(f"Sample prequal data: {prequal_sample}")
        for i, item in enumerate(prequal_sample):
            logging.debug(f"prequal[{i}] type: {type(item)}")
        if 'prequal' in self.data:
            self.populate_dropdowns()
        self.check_data_loaded()

    def check_data_loaded(self):
        if not self.data['prequal']:
            self.make_dropdown.setDisabled(True)
            self.model_dropdown.setDisabled(True)
            self.year_dropdown.setDisabled(True)
        else:
            self.make_dropdown.setDisabled(False)
            self.model_dropdown.setDisabled(False)
            self.year_dropdown.setDisabled(False)

    def populate_dropdowns(self):
        # Direct port of original Analyzer+ logic
        valid_prequal_data = [item for item in self.data['prequal'] if self.has_valid_prequal(item)]

        years = []
        for item in valid_prequal_data:
            try:
                if 'Year' in item and pd.notna(item['Year']):
                    year = int(float(item['Year']))
                    years.append(year)
            except (ValueError, TypeError):
                continue
        unique_years = sorted(set(years), reverse=True)
        self.year_dropdown.clear()
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.addItems([str(year) for year in unique_years])
        self.make_dropdown.clear()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")
        makes = set()
        for item in valid_prequal_data:
            try:
                if 'Make' in item and item['Make'] and isinstance(item['Make'], str):
                    make = item['Make'].strip()
                    if make and make.lower() != 'unknown':
                        makes.add(make)
            except (AttributeError, KeyError):
                continue
        if '' in makes:
            makes.remove('')
        self.make_dropdown.addItems(sorted(makes))
        logging.debug(f"Makes populated: {sorted(makes)}")

    def has_valid_prequal(self, item):
        # Match original app logic - only check Calibration Pre-Requisites
        return item.get('Calibration Pre-Requisites') not in [None, 'N/A']

    def display_blacklist(self, selected_make):
        """Display blacklist data"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='blacklist'")
            if not cursor.fetchone():
                self.blacklist_panel_widget.setPlainText("No blacklist data found. Please load data first.")
                return
            
            # Query blacklist data - use correct column names from schema
            if selected_make == "All":
                query = """
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                """
            else:
                query = f"""
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist WHERE carMake = '{selected_make}'
                """
            
            df = pd.read_sql_query(query, conn)
            if not df.empty:
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                if getattr(self, 'current_theme', 'Light') == 'Dark':
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                        background: #23262f;
                        color: #f1f1f1;
                    }}
                    .table th {{
                        background-color: #dc3545;
                        color: #fff;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #444;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #444;
                        color: #f1f1f1;
                        background: #23262f;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #2d2f3a;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                else:
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                    }}
                    .table th {{
                        background-color: #dc3545;
                        color: white;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #ddd;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #ddd;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #f8f9fa;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                self.blacklist_panel_widget.setHtml(styled_html)
                self.blacklist_panel_widget.setOpenExternalLinks(True)
            else:
                self.blacklist_panel_widget.setPlainText(f"No blacklist results found for make: {selected_make}")
        except Exception as e:
            logging.error(f"Failed to execute blacklist query: {query}\nError: {e}")
            self.blacklist_panel_widget.setPlainText(f"An error occurred while fetching blacklist data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def display_goldlist(self, selected_make):
        """Display goldlist data"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='goldlist'")
            if not cursor.fetchone():
                self.goldlist_panel_widget.setPlainText("No goldlist data found. Please load data first.")
                return
            
            # Query goldlist data - use correct column names from schema
            if selected_make == "All":
                query = """
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
                """
            else:
                query = f"""
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist WHERE carMake = '{selected_make}'
                """
            
            df = pd.read_sql_query(query, conn)
            if not df.empty:
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                if getattr(self, 'current_theme', 'Light') == 'Dark':
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                        background: #23262f;
                        color: #f1f1f1;
                    }}
                    .table th {{
                        background-color: #ffc107;
                        color: #23262f;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #444;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #444;
                        color: #f1f1f1;
                        background: #23262f;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #2d2f3a;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                else:
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                    }}
                    .table th {{
                        background-color: #ffc107;
                        color: black;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #ddd;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #ddd;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #f8f9fa;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                self.goldlist_panel_widget.setHtml(styled_html)
                self.goldlist_panel_widget.setOpenExternalLinks(True)
            else:
                self.goldlist_panel_widget.setPlainText(f"No goldlist results found for make: {selected_make}")
        except Exception as e:
            logging.error(f"Failed to execute goldlist query: {query}\nError: {e}")
            self.goldlist_panel_widget.setPlainText(f"An error occurred while fetching goldlist data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def display_mag_glass(self, selected_make):
        """Display Mag Glass data"""
        try:
            conn = self.get_db_connection()
            
            # Check if table exists
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mag_glass'")
            if not cursor.fetchone():
                self.mag_glass_panel_widget.setPlainText("No Mag Glass data found. Please load data first.")
                return
            
            # Prepare the query based on the selected make - use correct quoted column names
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
            
            if not df.empty:
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                if getattr(self, 'current_theme', 'Light') == 'Dark':
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                        background: #23262f;
                        color: #f1f1f1;
                    }}
                    .table th {{
                        background-color: #4f8cff;
                        color: #fff;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #444;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #444;
                        color: #f1f1f1;
                        background: #23262f;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #2d2f3a;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                else:
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                    }}
                    .table th {{
                        background-color: #667eea;
                        color: white;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #ddd;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #ddd;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #f8f9fa;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                self.mag_glass_panel_widget.setHtml(styled_html)
                self.mag_glass_panel_widget.setOpenExternalLinks(True)
            else:
                self.mag_glass_panel_widget.setPlainText(f"No Mag Glass results found for make: {selected_make}")
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.mag_glass_panel_widget.setPlainText(f"An error occurred while fetching the data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def display_carsys_data(self, selected_make):
        """Display CarSys data"""
        try:
            conn = self.get_db_connection()
            
            # Check if table exists
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='carsys'")
            if not cursor.fetchone():
                self.carsys_panel_widget.setPlainText("No CarSys data found. Please load data first.")
                return
            
            # If no make is selected, display all results
            if selected_make == "Select Make" or selected_make == "All":
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                """
                df = pd.read_sql_query(query, conn)
            else:
                # If a make is selected, filter the results by the selected make (case/whitespace-insensitive)
                query = """
                SELECT genericSystemName, dtcSys, carMake, comments
                FROM carsys
                WHERE TRIM(UPPER(carMake)) = TRIM(UPPER(?))
                """
                df = pd.read_sql_query(query, conn, params=(selected_make,))
            
            # Replace NaN values with empty strings
            df.fillna("", inplace=True)
            
            # Display the results in the CarSys panel
            if df.empty:
                self.carsys_panel_widget.setPlainText(f"No CarSys results found for make: {selected_make}")
            else:
                # Rename columns for display to match the original Excel column names
                df = df.rename(columns={
                    'genericSystemName': 'Generic System Name',
                    'dtcSys': 'DTCsys',
                    'carMake': 'CarMake',
                    'comments': 'Comments'
                })
                html_table = df.to_html(index=False, escape=False, classes='table table-striped')
                if getattr(self, 'current_theme', 'Light') == 'Dark':
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                        background: #23262f;
                        color: #f1f1f1;
                    }}
                    .table th {{
                        background-color: #4f8cff;
                        color: #fff;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #444;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #444;
                        color: #f1f1f1;
                        background: #23262f;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #2d2f3a;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                else:
                    styled_html = f"""
                    <html>
                    <head>
                    <style>
                    .table {{
                        border-collapse: collapse;
                        width: 100%;
                        font-family: Arial, sans-serif;
                        font-size: 12px;
                    }}
                    .table th {{
                        background-color: #667eea;
                        color: white;
                        padding: 8px;
                        text-align: left;
                        border: 1px solid #ddd;
                    }}
                    .table td {{
                        padding: 6px;
                        border: 1px solid #ddd;
                    }}
                    .table-striped tr:nth-child(even) {{
                        background-color: #f8f9fa;
                    }}
                    </style>
                    </head>
                    <body>
                    {html_table}
                    </body>
                    </html>
                    """
                self.carsys_panel_widget.setHtml(styled_html)
                self.carsys_panel_widget.setOpenExternalLinks(True)
        
        except Exception as e:
            logging.error(f"Failed to display CarSys data: {e}")
            self.carsys_panel_widget.setPlainText(f"An error occurred while fetching the CarSys data: {str(e)}")
        
        finally:
            conn.close()

    def on_tab_changed(self, index):
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        
        print(f"[DEBUG] Tab changed to index {index}")
        print(f"[DEBUG] Selected make: {selected_make}, model: {selected_model}, year: {selected_year}")
        
        if index == 0:  # Prequals
            # Prequals uses year, make, and model
            print("[DEBUG] Calling handle_prequal_search")
            self.handle_prequal_search(selected_make, selected_model, selected_year)
        elif index == 1:  # Blacklist
            # Blacklist only uses make - show all if "Select Make" or "All", otherwise filter by make
            if selected_make in ["Select Make", "All"]:
                print("[DEBUG] Calling display_blacklist with 'All'")
                self.display_blacklist("All")
            else:
                print(f"[DEBUG] Calling display_blacklist with '{selected_make}'")
                self.display_blacklist(selected_make)
        elif index == 2:  # Goldlist
            # Goldlist only uses make - show all if "Select Make" or "All", otherwise filter by make
            if selected_make in ["Select Make", "All"]:
                print("[DEBUG] Calling display_goldlist with 'All'")
                self.display_goldlist("All")
            else:
                print(f"[DEBUG] Calling display_goldlist with '{selected_make}'")
                self.display_goldlist(selected_make)
        elif index == 3:  # Mag Glass
            # Mag Glass only uses make - show all if "Select Make" or "All", otherwise filter by make
            if selected_make in ["Select Make", "All"]:
                print("[DEBUG] Calling display_mag_glass with 'All'")
                self.display_mag_glass("All")
            else:
                print(f"[DEBUG] Calling display_mag_glass with '{selected_make}'")
                self.display_mag_glass(selected_make)
        elif index == 4:  # CarSys
            # CarSys only uses make - show all if "Select Make" or "All", otherwise filter by make
            if selected_make in ["Select Make", "All"]:
                print("[DEBUG] Calling display_carsys_data with 'All'")
                self.display_carsys_data("All")
            else:
                print(f"[DEBUG] Calling display_carsys_data with '{selected_make}'")
                self.display_carsys_data(selected_make)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Analyzer+ Professional")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Professional Vehicle Analysis")
    
    # Create and show the main window
    ex = ModernAnalyzerApp()
    ex.show()
    
    sys.exit(app.exec_()) 
