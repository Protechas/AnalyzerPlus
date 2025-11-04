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

# DEBUG: Print which Python is being used
print(f"="*70)
print(f"PYTHON EXECUTABLE: {sys.executable}")
print(f"PYTHON VERSION: {sys.version}")
print(f"="*70)

import pandas as pd
from PyQt5.QtCore import Qt, QUrl, QTimer, QPropertyAnimation, QEasingCurve, QRect, QRectF, pyqtProperty, pyqtSignal
from PyQt5.QtGui import QDesktopServices, QIcon, QKeySequence, QFont, QPalette, QColor, QPixmap, QPainter, QLinearGradient, QPen, QBrush, QTextCursor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QListWidget, QLabel, QMessageBox, QFileDialog, QInputDialog,
    QShortcut, QSizePolicy, QFormLayout, QScrollArea, QGridLayout, 
    QGroupBox, QCheckBox, QRadioButton, QButtonGroup, QStackedWidget, 
    QStyle, QFrame, QPushButton, QComboBox, QLineEdit, QTextBrowser,
    QProgressBar, QSlider, QTabWidget, QSplitter, QStatusBar, QToolBar,
    QDialog, QSpinBox
)
from PyQt5.QtWidgets import QStyleFactory
from PyQt5.QtWidgets import QGraphicsDropShadowEffect, QGraphicsOpacityEffect
from modern_components import (
    ModernDialog, ModernButton, ModernComboBox, ModernTextBrowser,
    ModernLineEdit, ModernProgressBar, ModernSlider, ModernTabWidget,
    ModernSplitter, ModernStatusBar, ModernToolBar
)
from multi_vehicle_compare import MultiVehicleCompareDialog
from database_utils import get_prequal_data, get_unique_makes, get_unique_models, get_unique_years

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
        self.current_font_size = 13
        self.setStyleSheet("""
            QTextBrowser {
                border: 2px solid #e9ecef;
                border-radius: 8px;
                background: white;
                line-height: 1.5;
            }
            QTextBrowser:focus {
                border-color: #667eea;
            }
        """)
        # Set initial font
        font = QFont()
        font.setPointSize(self.current_font_size)
        self.setFont(font)
        self.document().setDefaultFont(font)
    
    def set_font_size(self, size):
        """Change the actual font size"""
        self.current_font_size = size
        
        # Update widget font
        font = QFont()
        font.setPointSize(size)
        self.setFont(font)
        
        # Update document default font
        self.document().setDefaultFont(font)
        
        # Change font size of all existing text using QTextCursor
        cursor = QTextCursor(self.document())
        cursor.select(QTextCursor.Document)
        char_format = cursor.charFormat()
        char_format.setFontPointSize(size)
        cursor.mergeCharFormat(char_format)

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
            CREATE TABLE IF NOT EXISTS manufacturer_chart (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Year TEXT,
                Make TEXT,
                Model TEXT,
                Calibration_Type TEXT,
                Protech_Generic_System_Name TEXT,
                SME_Generic_System_Name TEXT,
                SME_Calibration_Type TEXT,
                Feature TEXT,
                Service_Information_Hyperlink TEXT,
                Calibration_Pre_Requisites TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
        logging.error(error_message)
        error_messages.append(error_message)
        QMessageBox.critical(parent, "File Access Error", error_message)
        return error_message
        error_message = f"File not found: {excel_path}"
        logging.error(error_message)
        if parent:
            QMessageBox.critical(parent, "File Not Found", error_message)
        return error_message
    except Exception as read_error:
        error_message = f"Error reading Excel file {excel_path}: {str(read_error)}"
        logging.error(error_message)
        error_messages.append(error_message)
        if parent:
            QMessageBox.critical(parent, "File Read Error", error_message)
            return error_message
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
            # For prequal, just store as JSON - ignore comment columns
            # Remove any columns that contain 'comment' in their name (case insensitive)
            comment_columns = [col for col in df.columns if 'comment' in col.lower()]
            if comment_columns:
                df = df.drop(columns=comment_columns)
                logging.info(f"Removed comment columns from prequal data: {comment_columns}")
            
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
        elif table_name == 'manufacturer_chart':
            # For manufacturer chart, create table with expected columns and load all data
            # Define expected columns for manufacturer chart (normalized)
            expected_columns = [
                'Year', 'Make', 'Model', 'Calibration_Type', 'Protech_Generic_System_Name',
                'SME_Generic_System_Name', 'SME_Calibration_Type', 'Feature',
                'Service_Information_Hyperlink', 'Calibration_Pre_Requisites'
            ]
            
            # Normalize the dataframe columns to match expected format
            df.columns = [col.strip() for col in df.columns]
            
            # Create the table with proper schema
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Check if table exists, if not create it
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
            if not cursor.fetchone():
                logging.info("Creating manufacturer_chart table")
            
                # Create table with all expected columns
                create_table_sql = """
                CREATE TABLE manufacturer_chart (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year TEXT,
                    Make TEXT,
                    Model TEXT,
                    Calibration_Type TEXT,
                    Protech_Generic_System_Name TEXT,
                    SME_Generic_System_Name TEXT,
                    SME_Calibration_Type TEXT,
                    Feature TEXT,
                    Service_Information_Hyperlink TEXT,
                    Calibration_Pre_Requisites TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
                cursor.execute(create_table_sql)
                logging.info("Created manufacturer_chart table")
            
            # Log available columns for debugging
            logging.info(f"Available columns in Excel file: {list(df.columns)}")
            
            # Debug: Check if this looks like an Acura file by checking the data
            if len(df) > 0:
                sample_make = str(df.iloc[0].get('Make', df.iloc[0].get('make', '')))
                if 'acura' in sample_make.lower():
                    logging.info(f"This appears to be an Acura file. Sample make: '{sample_make}'")
                    logging.info(f"All column names: {[col for col in df.columns]}")
            
            # Create a mapping dictionary to handle various column name formats
            column_mapping = {}
            for col in df.columns:
                col_lower = col.lower().strip()
                col_clean = col_lower.replace(' ', '').replace('-', '').replace('_', '')
                
                # Map exact matches first
                if col_lower == 'year':
                    column_mapping['Year'] = col
                elif col_lower == 'make':
                    column_mapping['Make'] = col
                elif col_lower == 'model':
                    column_mapping['Model'] = col
                elif col_lower == 'calibration type':
                    column_mapping['Calibration_Type'] = col
                elif col_lower == 'protech generic system name':
                    column_mapping['Protech_Generic_System_Name'] = col
                elif col_lower == 'sme generic system name':
                    column_mapping['SME_Generic_System_Name'] = col
                elif col_lower == 'sme calibration type':
                    column_mapping['SME_Calibration_Type'] = col
                elif col_lower == 'feature':
                    column_mapping['Feature'] = col
                elif col_lower == 'service information hyperlink':
                    column_mapping['Service_Information_Hyperlink'] = col
                elif col_lower == 'calibration pre-requisites':
                    column_mapping['Calibration_Pre_Requisites'] = col
                # Fallback to partial matches
                elif 'year' in col_lower:
                    column_mapping['Year'] = col
                elif 'make' in col_lower:
                    column_mapping['Make'] = col
                elif 'model' in col_lower:
                    column_mapping['Model'] = col
                elif 'calibration' in col_lower and 'type' in col_lower:
                    column_mapping['Calibration_Type'] = col
                elif 'protech' in col_lower and 'system' in col_lower and 'name' in col_lower:
                    column_mapping['Protech_Generic_System_Name'] = col
                elif 'sme' in col_lower and 'system' in col_lower and 'name' in col_lower:
                    column_mapping['SME_Generic_System_Name'] = col
                elif 'sme' in col_lower and 'calibration' in col_lower and 'type' in col_lower:
                    column_mapping['SME_Calibration_Type'] = col
                elif 'feature' in col_lower:
                    column_mapping['Feature'] = col
                elif 'service' in col_lower and 'hyperlink' in col_lower:
                    column_mapping['Service_Information_Hyperlink'] = col
                elif 'prerequisites' in col_lower or 'pre-requisites' in col_lower:
                    column_mapping['Calibration_Pre_Requisites'] = col
            
            logging.info(f"Column mapping: {column_mapping}")
            
            # Insert data - map columns as best as possible
            for index, row in df.iterrows():
                # Extract values using the mapping
                year = str(row.get(column_mapping.get('Year', 'Year'), '')).strip() if pd.notna(row.get(column_mapping.get('Year', 'Year'), '')) else ''
                make = str(row.get(column_mapping.get('Make', 'Make'), '')).strip() if pd.notna(row.get(column_mapping.get('Make', 'Make'), '')) else ''
                model = str(row.get(column_mapping.get('Model', 'Model'), '')).strip() if pd.notna(row.get(column_mapping.get('Model', 'Model'), '')) else ''
                calibration_type = str(row.get(column_mapping.get('Calibration_Type', 'Calibration Type'), '')).strip() if pd.notna(row.get(column_mapping.get('Calibration_Type', 'Calibration Type'), '')) else ''
                protech_system = str(row.get(column_mapping.get('Protech_Generic_System_Name', 'Protech Generic System Name'), '')).strip() if pd.notna(row.get(column_mapping.get('Protech_Generic_System_Name', 'Protech Generic System Name'), '')) else ''
                sme_system = str(row.get(column_mapping.get('SME_Generic_System_Name', 'SME Generic System Name'), '')).strip() if pd.notna(row.get(column_mapping.get('SME_Generic_System_Name', 'SME Generic System Name'), '')) else ''
                sme_calibration = str(row.get(column_mapping.get('SME_Calibration_Type', 'SME Calibration Type'), '')).strip() if pd.notna(row.get(column_mapping.get('SME_Calibration_Type', 'SME Calibration Type'), '')) else ''
                feature = str(row.get(column_mapping.get('Feature', 'Feature'), '')).strip() if pd.notna(row.get(column_mapping.get('Feature', 'Feature'), '')) else ''
                service_link = str(row.get(column_mapping.get('Service_Information_Hyperlink', 'Service Information Hyperlink'), '')).strip() if pd.notna(row.get(column_mapping.get('Service_Information_Hyperlink', 'Service Information Hyperlink'), '')) else ''
                prerequisites = str(row.get(column_mapping.get('Calibration_Pre_Requisites', 'Calibration Pre-Requisites'), '')).strip() if pd.notna(row.get(column_mapping.get('Calibration_Pre_Requisites', 'Calibration Pre-Requisites'), '')) else ''
                
                # Insert the record
                insert_sql = """
                INSERT INTO manufacturer_chart (
                    Year, Make, Model, Calibration_Type, Protech_Generic_System_Name,
                    SME_Generic_System_Name, SME_Calibration_Type, Feature,
                    Service_Information_Hyperlink, Calibration_Pre_Requisites
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                cursor.execute(insert_sql, (
                    year, make, model, calibration_type, protech_system,
                    sme_system, sme_calibration, feature, service_link, prerequisites
                ))
                
                # Log the first few records for debugging
                if index < 3:
                    logging.info(f"Inserted record {index + 1}: Year='{year}', Make='{make}', Model='{model}'")
                
                # Log every 1000th record to see what data is being loaded
                if index % 1000 == 0:
                    logging.info(f"Processing record {index}: Year='{year}', Make='{make}', Model='{model}'")
                
                # Log specifically if we find Acura data
                if make.upper() == 'ACURA':
                    logging.info(f"Found Acura record {index}: Year='{year}', Make='{make}', Model='{model}'")
            
            conn.commit()
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
            "prequal": "Prequals (Longsheets)",
            "manufacturer_chart": "Manufacturer Chart"
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
                
                # Special handling for manufacturer_chart - use simple pandas approach like prequals
                logging.info(f"Processing config_type: {config_type}")
                if config_type == 'manufacturer_chart':
                    logging.info("Loading manufacturer chart data using simple pandas approach...")
                    
                    # Clear existing manufacturer chart data
                    conn = sqlite3.connect(self.parent.db_path)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM manufacturer_chart")
                    conn.commit()
                    conn.close()
                    logging.info("Cleared existing manufacturer chart data")
                    
                    # Process each file like prequals
                    for i, (filename, filepath) in enumerate(files.items()):
                        try:
                            logging.info(f"Processing manufacturer chart file: {filename}")
                            
                            # Check if file is accessible (not a SharePoint placeholder)
                            import os
                            if not os.path.exists(filepath):
                                logging.warning(f"File not found or not accessible: {filename}")
                                continue
                            
                            # Check file size - if it's 0 or very small, it might be a placeholder
                            file_size = os.path.getsize(filepath)
                            if file_size < 1024:  # Less than 1KB
                                logging.warning(f"File appears to be a placeholder or empty (size: {file_size} bytes): {filename}")
                                continue
                            
                            if 'acura' in filename.lower():
                                logging.info(f"Found potential Acura file: {filename}")
                            
                            # Check for "Model Version" sheet first, then fall back to first sheet
                            try:
                                # Try to read the "Model Version" sheet first
                                df = pd.read_excel(filepath, sheet_name="Model Version")
                                logging.info(f"Found 'Model Version' sheet in {filename}")
                            except Exception as sheet_error:
                                # If "Model Version" sheet doesn't exist, use the first sheet
                                try:
                                    df = pd.read_excel(filepath)
                                    logging.info(f"Using first sheet for {filename}")
                                except Exception as read_error:
                                    logging.error(f"Cannot read Excel file {filename}: {str(read_error)}")
                                    continue
                            
                            if df.empty:
                                logging.warning(f"{filename} is empty.")
                                continue
                            
                            # Log the columns found in this file
                            logging.info(f"Columns in {filename}: {list(df.columns)}")
                            
                            # Convert to records and save to manufacturer_chart table
                            data = df.to_dict(orient='records')
                            self.save_manufacturer_chart_data(data, self.parent.db_path)
                            data_loaded = True
                            logging.info(f"Loaded manufacturer chart data from: {filename}")
                            
                        except Exception as e:
                            logging.error(f"Error loading manufacturer chart from {filename}: {str(e)}")
                            import traceback
                            logging.error(traceback.format_exc())
                        self.parent.progress_bar.setValue(i + 1)
                else:
                    # Process individual files for other data types
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
                        except Exception as e:
                            logging.error(f"Error loading {filename}: {str(e)}")
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

    def load_manufacturer_chart_data(self, filepath, db_path):
        """Load manufacturer chart data from Excel file with specified columns"""
        try:
            # Read Excel file
            df = pd.read_excel(filepath)
            
            if df.empty:
                return "No data found in file"
            
            # Normalize column names (case-insensitive, remove special characters)
            df.columns = [normalize_col(col) for col in df.columns]
            
            # Define expected column names (normalized)
            expected_columns = [
                'year', 'make', 'model', 'calibrationtype', 'protechgenericsystemname',
                'smegenericsystemname', 'smecalibrationtype', 'feature', 
                'serviceinformationhyperlink', 'calibrationprerequisites'
            ]
            
            # Check if required columns exist (with variations)
            missing_columns = []
            for expected_col in expected_columns:
                if not any(expected_col in col.lower() for col in df.columns):
                    missing_columns.append(expected_col)
            
            if missing_columns:
                logging.warning(f"Missing columns in manufacturer chart: {missing_columns}")
                # Continue with available columns
            
            # Create table if it doesn't exist
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Create manufacturer_chart table with all possible columns
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS manufacturer_chart (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year TEXT,
                    Make TEXT,
                    Model TEXT,
                    Calibration_Type TEXT,
                    Protech_Generic_System_Name TEXT,
                    SME_Generic_System_Name TEXT,
                    SME_Calibration_Type TEXT,
                    Feature TEXT,
                    Service_Information_Hyperlink TEXT,
                    Calibration_Pre_Requisites TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Clear existing data
            cursor.execute("DELETE FROM manufacturer_chart")
            
            # Map normalized columns to database columns
            column_mapping = {
                'year': 'Year',
                'make': 'Make', 
                'model': 'Model',
                'calibrationtype': 'Calibration_Type',
                'protechgenericsystemname': 'Protech_Generic_System_Name',
                'smegenericsystemname': 'SME_Generic_System_Name',
                'smecalibrationtype': 'SME_Calibration_Type',
                'feature': 'Feature',
                'serviceinformationhyperlink': 'Service_Information_Hyperlink',
                'calibrationprerequisites': 'Calibration_Pre_Requisites'
            }
            
            # Log normalized column names for debugging
            logging.info(f"Normalized DataFrame columns: {df.columns.tolist()}")
            
            # Prepare data for insertion
            records = []
            for _, row in df.iterrows():
                record = {}
                for normalized_col, db_col in column_mapping.items():
                    # Since columns are already normalized, do exact match
                    if normalized_col in df.columns:
                        value = row[normalized_col]
                        if pd.isna(value):
                            record[db_col] = ''  # Use empty string for consistency
                        else:
                            record[db_col] = str(value).strip()
                    else:
                        record[db_col] = ''
                        logging.warning(f"Column '{normalized_col}' not found in DataFrame")
                
                records.append(record)
            
            # Log first record for debugging
            if records:
                logging.info(f"First CMC record to insert: {records[0]}")
            
            # Insert data
            if records:
                placeholders = ', '.join(['?' for _ in column_mapping.values()])
                columns = ', '.join(column_mapping.values())
                cursor.executemany(
                    f"INSERT INTO manufacturer_chart ({columns}) VALUES ({placeholders})",
                    [tuple(record.values()) for record in records]
                )
                
                conn.commit()
                conn.close()
                logging.info(f"Successfully inserted {len(records)} CMC records")
                return "Data loaded successfully"
            else:
                conn.close()
                return "No valid records found"
                
        except Exception as e:
            logging.error(f"Error loading manufacturer chart data: {str(e)}")
            return f"Error: {str(e)}"

    def save_manufacturer_chart_data(self, data, db_path):
        """Save manufacturer chart data to database using simple approach like prequals"""
        try:
            # Create table if it doesn't exist
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Create manufacturer_chart table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS manufacturer_chart (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year TEXT,
                    Make TEXT,
                    Model TEXT,
                    Calibration_Type TEXT,
                    Protech_Generic_System_Name TEXT,
                    SME_Generic_System_Name TEXT,
                    SME_Calibration_Type TEXT,
                    Feature TEXT,
                    Service_Information_Hyperlink TEXT,
                    Calibration_Pre_Requisites TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Insert data using simple approach
            for record in data:
                # Map columns flexibly (case-insensitive)
                year = record.get('Year', record.get('year', ''))
                make = record.get('Make', record.get('make', ''))
                model = record.get('Model', record.get('model', ''))
                calibration_type = record.get('Calibration Type', record.get('calibration type', record.get('CalibrationType', '')))
                protech_system = record.get('Protech Generic System Name', record.get('protech generic system name', ''))
                sme_system = record.get('SME Generic System Name', record.get('sme generic system name', ''))
                sme_calibration = record.get('SME Calibration Type', record.get('sme calibration type', ''))
                feature = record.get('Feature', record.get('feature', ''))
                service_link = record.get('Service Information Hyperlink', record.get('service information hyperlink', ''))
                prerequisites = record.get('Calibration Pre-Requisites', record.get('calibration pre-requisites', ''))
                
                cursor.execute('''
                    INSERT INTO manufacturer_chart 
                    (Year, Make, Model, Calibration_Type, Protech_Generic_System_Name, 
                     SME_Generic_System_Name, SME_Calibration_Type, Feature, 
                     Service_Information_Hyperlink, Calibration_Pre_Requisites)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (year, make, model, calibration_type, protech_system, 
                      sme_system, sme_calibration, feature, service_link, prerequisites))
            
            conn.commit()
            conn.close()
            logging.info(f"Saved {len(data)} manufacturer chart records to database")
            
        except Exception as e:
            logging.error(f"Error saving manufacturer chart data: {str(e)}")
            raise

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
        
        # Lock state variables
        self.year_locked = False
        self.make_locked = False
        self.model_locked = False
        self.locked_year = None
        self.locked_make = None
        self.locked_model = None
        
        # Region data
        self.region_makes = {
            'German': ['BMW', 'MINI', 'Rolls-Royce', 'Volkswagen', 'Audi', 'Porsche', 'Fiat', 'Alfa Romeo', 'Jaguar', 'Land Rover', 'Volvo'],
            'Asian': ['Honda', 'Acura', 'Toyota', 'Lexus', 'Nissan', 'Infinity', 'Mitsubishi', 'Mazda', 'Subaru', 'Kia', 'Hyundai', 'Genesis'],
            'US': ['Buick', 'Cadillac', 'Chevrolet', 'GMC', 'Ford', 'Lincoln', 'Mercury', 'Chrysler', 'Dodge', 'Jeep', 'Ram']
        }
        self.current_region = 'ALL'  # Default to ALL
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

        # Region toggle section
        region_card = ModernCard()
        region_card.setFixedSize(200, 60)
        region_card.setObjectName("region_card")  # Add object name for styling
        region_layout = QVBoxLayout(region_card)
        region_layout.setContentsMargins(10, 6, 10, 6)
        
        # Region toggle button - styled as a toggle switch
        self.region_toggle_button = QPushButton("ALL/REGION")
        self.region_toggle_button.setCheckable(True)
        self.region_toggle_button.setFixedSize(140, 32)
        self.region_toggle_button.setStyleSheet("""
            QPushButton {
                background: #28a745;
                color: white;
                border: 1px solid #28a745;
                border-radius: 16px;
                font-size: 12px;
                font-weight: bold;
                font-family: 'Segoe UI', Arial, sans-serif;
                text-align: center;
                padding: 4px 8px;
            }
            QPushButton:hover {
                background: #218838;
                border-color: #218838;
            }
            QPushButton:checked {
                background: #007bff;
                border-color: #007bff;
            }
            QPushButton:checked:hover {
                background: #0056b3;
                border-color: #0056b3;
            }
        """)
        self.region_toggle_button.clicked.connect(self.toggle_region_mode)
        region_layout.addWidget(self.region_toggle_button)
        
        # Region dropdown (initially hidden)
        self.region_dropdown = ModernComboBox()
        self.region_dropdown.addItems(["Asian", "German", "US"])
        self.region_dropdown.setStyleSheet("""
            QComboBox {
                background: #fff;
                color: #20567C;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                min-width: 80px;
            }
        """)
        self.region_dropdown.currentTextChanged.connect(self.on_region_changed)
        self.region_dropdown.hide()
        self.region_dropdown.setMaximumHeight(0)  # Ensure it takes no space when hidden
        region_layout.addWidget(self.region_dropdown)
        
        header_layout.addWidget(region_card)

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
        # self.add_toolbar_button("Refresh Lists", self.refresh_lists, "refresh_button")  # Hidden for now
        self.add_toolbar_button("Compare Vehicles", self.open_compare_dialog, "compare_button")
        clear_button = ModernButton("Clear Filters", style="secondary")
        clear_button.clicked.connect(self.clear_filters)
        self.toolbar.addWidget(clear_button)
        self.always_on_top_button = ModernButton("Pin to Top", style="secondary")
        self.always_on_top_button.setCheckable(True)
        self.always_on_top_button.clicked.connect(self.toggle_always_on_top)
        self.toolbar.addWidget(self.always_on_top_button)
        self.toolbar.addSeparator()
        
        # Font size control
        font_size_label = QLabel("Font Size:")
        self.toolbar.addWidget(font_size_label)
        self.font_size_spinbox = QSpinBox()
        self.font_size_spinbox.setRange(8, 32)
        self.font_size_spinbox.setValue(13)
        self.font_size_spinbox.setSuffix(" pt")
        self.font_size_spinbox.setStyleSheet("""
            QSpinBox {
                background: #fff;
                color: #20567C;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 14px;
                min-width: 80px;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                background: #19507a;
                border-radius: 3px;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background: #1e5d8a;
            }
        """)
        self.font_size_spinbox.valueChanged.connect(self.update_font_size)
        self.toolbar.addWidget(self.font_size_spinbox)
        
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

        # Shared style for field containers (white rounded box with label and input)
        field_container_style = (
            "QWidget { background: #ffffff; border: 2px solid #e9ecef; border-radius: 12px; padding: 6px 10px; }"
            "QLabel { font-weight: 600; font-size: 14px; color: #20567C; margin-right: 8px; }"
        )
        
        # Lock button styles
        self.lock_button_style = (
            "QPushButton { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 4px; min-width: 20px; max-width: 20px; }"
            "QPushButton:hover { background: #e9ecef; }"
            "QPushButton:pressed { background: #dee2e6; }"
        )
        
        self.lock_button_style_locked = (
            "QPushButton { background: #28a745; border: 1px solid #28a745; border-radius: 4px; padding: 4px; min-width: 20px; max-width: 20px; color: white; }"
            "QPushButton:hover { background: #218838; }"
            "QPushButton:pressed { background: #1e7e34; }"
        )

        # Year container (label + dropdown + lock button inside one white box)
        year_container = QWidget()
        year_container.setStyleSheet(field_container_style)
        year_layout = QHBoxLayout(year_container)
        year_layout.setSpacing(8)
        year_layout.setContentsMargins(10, 6, 10, 6)
        year_label = QLabel("Year:")
        year_layout.addWidget(year_label)
        self.year_dropdown = ModernComboBox()
        self.year_dropdown.setFixedWidth(120)
        self.year_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.currentIndexChanged.connect(self.on_year_selected)
        year_layout.addWidget(self.year_dropdown)
        
        # Year lock button
        self.year_lock_button = QPushButton("🔓")
        self.year_lock_button.setToolTip("Lock Year selection")
        self.year_lock_button.setStyleSheet(self.lock_button_style)
        self.year_lock_button.clicked.connect(lambda: self.toggle_lock('year'))
        year_layout.addWidget(self.year_lock_button)
        search_layout.addWidget(year_container)

        # Make container (label + dropdown + lock button)
        make_container = QWidget()
        make_container.setStyleSheet(field_container_style)
        make_layout = QHBoxLayout(make_container)
        make_layout.setSpacing(8)
        make_layout.setContentsMargins(10, 6, 10, 6)
        make_label = QLabel("Make:")
        make_layout.addWidget(make_label)
        self.make_dropdown = ModernComboBox()
        self.make_dropdown.setFixedWidth(120)
        self.make_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        self.make_dropdown.currentIndexChanged.connect(self.update_model_dropdown)
        make_layout.addWidget(self.make_dropdown)
        
        # Make lock button
        self.make_lock_button = QPushButton("🔓")
        self.make_lock_button.setToolTip("Lock Make selection")
        self.make_lock_button.setStyleSheet(self.lock_button_style)
        self.make_lock_button.clicked.connect(lambda: self.toggle_lock('make'))
        make_layout.addWidget(self.make_lock_button)
        search_layout.addWidget(make_container)

        # Model container (label + dropdown + lock button)
        model_container = QWidget()
        model_container.setStyleSheet(field_container_style)
        model_layout = QHBoxLayout(model_container)
        model_layout.setSpacing(8)
        model_layout.setContentsMargins(10, 6, 10, 6)
        model_label = QLabel("Model:")
        model_layout.addWidget(model_label)
        self.model_dropdown = ModernComboBox()
        self.model_dropdown.setFixedWidth(130)
        self.model_dropdown.setStyleSheet("padding: 6px 12px; font-size: 14px;")
        self.model_dropdown.addItem("Select Model")
        self.model_dropdown.currentIndexChanged.connect(self.handle_model_change)
        model_layout.addWidget(self.model_dropdown)
        
        # Model lock button
        self.model_lock_button = QPushButton("🔓")
        self.model_lock_button.setToolTip("Lock Model selection")
        self.model_lock_button.setStyleSheet(self.lock_button_style)
        self.model_lock_button.clicked.connect(lambda: self.toggle_lock('model'))
        model_layout.addWidget(self.model_lock_button)
        search_layout.addWidget(model_container)

        # Search container (label + line edit + button) - expands with window
        search_container = QWidget()
        search_container.setStyleSheet(field_container_style)
        search_container_layout = QHBoxLayout(search_container)
        search_container_layout.setSpacing(8)
        search_container_layout.setContentsMargins(10, 6, 10, 6)
        
        # Fixed width label
        search_label = QLabel("Search:")
        search_label.setFixedWidth(100)
        search_label.setStyleSheet("font-weight: bold;")
        search_container_layout.addWidget(search_label)
        
        # Expanding search bar
        self.search_bar = ModernLineEdit("Enter DTC code or description (searches Blacklist & Goldlist only)")
        self.search_bar.setMinimumWidth(150)
        self.search_bar.setStyleSheet("padding: 6px 12px; font-size: 12px;")
        self.search_bar.returnPressed.connect(self.perform_search)
        search_container_layout.addWidget(self.search_bar)
        
        # Search button
        search_button = ModernButton("🔍 Search", style="secondary")
        search_button.setFixedWidth(150)
        search_button.clicked.connect(self.perform_search)
        search_container_layout.addWidget(search_button)
        
        search_layout.addWidget(search_container, 1)  # Stretch factor of 1 to expand

        main_layout.addWidget(search_card)

    def create_main_content(self, main_layout):
        """Create main content area with multi-toggle tab row and horizontal split display"""
        content_card = ModernCard()
        content_layout = QVBoxLayout(content_card)

        # Multi-toggle tab row
        self.tab_buttons = {}
        self.tab_panels = {}
        tab_info = [
            ("CMC", "prequals_panel"),  # Use prequals_panel for both CMC and Prequals
            ("Blacklist", "blacklist_panel"),
            ("Goldlist", "goldlist_panel"),
            ("Mag Glass", "mag_glass_panel")
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
        self.tab_panels = {
            "prequals_panel": self.prequals_panel,
            "blacklist_panel": self.blacklist_panel,
            "goldlist_panel": self.goldlist_panel,
            "mag_glass_panel": self.mag_glass_panel
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
                self.update_prequals_cmc_display()
            elif attr == "blacklist_panel":
                self.display_blacklist(self.make_dropdown.currentText())
            elif attr == "goldlist_panel":
                self.display_goldlist(self.make_dropdown.currentText())
            elif attr == "mag_glass_panel":
                self.display_mag_glass(self.make_dropdown.currentText())

    def create_prequals_panel(self):
        card = ModernCard()
        card_layout = QVBoxLayout(card)
        
        # Header with toggle switch
        header_layout = QHBoxLayout()
        
        # Create a custom sliding toggle widget
        class SlidingToggle(QWidget):
            # Define a signal for position changes
            position_changed = pyqtSignal(int)
            
            def __init__(self, parent=None):
                super().__init__(parent)
                self.setFixedSize(140, 30)
                self.current_position = 0  # 0 for CMC, 1 for Prequals
                self.animation = QPropertyAnimation(self, b"slider_position")
                self.animation.setDuration(200)
                self.animation.setEasingCurve(QEasingCurve.OutCubic)
                
            def _get_slider_position(self):
                return self.current_position
                
            def _set_slider_position(self, position):
                self.current_position = position
                self.update()
                
            slider_position = pyqtProperty(float, _get_slider_position, _set_slider_position)
            
            def paintEvent(self, event):
                painter = QPainter(self)
                painter.setRenderHint(QPainter.Antialiasing)
                
                # Draw background
                painter.setPen(QPen(QColor("#20567C"), 2))
                painter.setBrush(QBrush(QColor("#f8f9fa")))
                painter.drawRoundedRect(self.rect(), 15, 15)
                
                # Calculate slider position
                slider_width = self.width() // 2
                slider_x = int(self.current_position * slider_width)
                
                # Draw sliding highlight
                painter.setPen(Qt.NoPen)
                painter.setBrush(QBrush(QColor("#20567C")))
                painter.drawRoundedRect(slider_x, 2, slider_width, self.height() - 4, 13, 13)
                
                # Draw text
                painter.setPen(QColor("#6c757d"))
                painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
                
                # CMC text
                cmc_rect = QRectF(0, 0, slider_width, self.height())
                if self.current_position == 0:
                    painter.setPen(QColor("white"))
                painter.drawText(cmc_rect, Qt.AlignCenter, "CMC")
                
                # Prequals text
                prequals_rect = QRectF(slider_width, 0, slider_width, self.height())
                if self.current_position == 1:
                    painter.setPen(QColor("white"))
                else:
                    painter.setPen(QColor("#6c757d"))
                painter.drawText(prequals_rect, Qt.AlignCenter, "Prequals")
            
            def mousePressEvent(self, event):
                if event.button() == Qt.LeftButton:
                    # Determine which half was clicked
                    if event.x() < self.width() // 2:
                        self.slide_to(0)
                    else:
                        self.slide_to(1)
            
            def slide_to(self, position):
                if position != self.current_position:
                    self.animation.setStartValue(self.current_position)
                    self.animation.setEndValue(position)
                    self.animation.start()
                    self.current_position = position
                    # Emit signal
                    print(f"DEBUG: Emitting position_changed signal with position {position}")
                    self.position_changed.emit(position)
        
        # Create the sliding toggle
        self.sliding_toggle = SlidingToggle()
        self.sliding_toggle.current_position = 0  # Start with CMC selected
        # Connect the signal to the callback
        self.sliding_toggle.position_changed.connect(self.on_toggle_changed)
        
        # Add the toggle to the header
        header_layout.addWidget(self.sliding_toggle)
        # Hide the toggle for now
        self.sliding_toggle.hide()
        header_layout.addStretch()
        header_layout.addStretch()
        
        # Title label (will be updated based on toggle)
        self.prequals_cmc_title = QLabel("Manufacturer Chart Data")
        self.prequals_cmc_title.setObjectName("panel_title_label")
        self.prequals_cmc_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #20567C;")
        header_layout.addWidget(self.prequals_cmc_title)
        header_layout.addStretch()
        
        expand_button = ModernButton("Expand View", style="secondary")
        expand_button.clicked.connect(self.expand_prequals_cmc_panel)
        header_layout.addWidget(expand_button)
        card_layout.addLayout(header_layout)
        
        self.left_panel = ModernTextBrowser()
        self.left_panel.setStyleSheet("background: #f8fafc; border-radius: 8px; font-size: 14px; padding: 8px;")
        card_layout.addWidget(self.left_panel)
        
        # Initialize with CMC data
        self.current_data_mode = "CMC"
        
        # Set initial display to CMC data
        self.display_cmc_data("Select Year", "Select Make", "Select Model")
        
        return card

    def on_toggle_changed(self, position):
        """Handle sliding toggle position change"""
        print(f"DEBUG: on_toggle_changed called with position {position}")
        if position == 0:  # CMC
            self.current_data_mode = "CMC"
            self.prequals_cmc_title.setText("Manufacturer Chart Data")
            # Update the main tab button text
            if "prequals_panel" in self.tab_buttons:
                print(f"DEBUG: Setting tab text to CMC")
                self.tab_buttons["prequals_panel"].setText("CMC")
        else:  # Prequals
            self.current_data_mode = "Prequals"
            self.prequals_cmc_title.setText("Prequalification Data")
            # Update the main tab button text
            if "prequals_panel" in self.tab_buttons:
                print(f"DEBUG: Setting tab text to Prequals")
                self.tab_buttons["prequals_panel"].setText("Prequals")
        
        # Update the display
        self.update_prequals_cmc_display()
        
        # Update tab button styles to reflect the change
        self.update_tab_button_styles()

    def update_prequals_cmc_display(self):
        """Update the display based on current toggle state"""
        if self.current_data_mode == "CMC":
            self.display_cmc_data(self.year_dropdown.currentText(), self.make_dropdown.currentText(), self.model_dropdown.currentText())
        else:
            self.handle_prequal_search(self.make_dropdown.currentText(), self.model_dropdown.currentText(), self.year_dropdown.currentText())

    def expand_prequals_cmc_panel(self):
        """Expand the current panel content"""
        title = "Manufacturer Chart Data" if self.current_data_mode == "CMC" else "Prequalification Data"
        self.pop_out_panel(title, self.left_panel.toHtml())

    def update_tab_button_styles(self):
        """Update tab button styles after text changes"""
        for attr, btn in self.tab_buttons.items():
            btn.setStyleSheet(self.get_tab_button_style(selected=btn.isChecked()))

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
        # Don't update if year is locked
        if self.year_locked:
            return
            
        selected_year = self.year_dropdown.currentText()
        current_make = self.make_dropdown.currentText()
        current_model = self.model_dropdown.currentText()
        
        print(f"[DEBUG] on_year_selected: Year selected: '{selected_year}'")
        print(f"[DEBUG] on_year_selected: Current make: '{current_make}'")
        print(f"[DEBUG] on_year_selected: Current model: '{current_model}'")
        
        # Update dropdowns to show only available options based on current selections
        self.update_dropdowns_with_locks()
        
        # Update panels
        self.update_visible_panels()

    def update_model_dropdown(self):
        # Don't update if make is locked
        if self.make_locked:
            return
            
        selected_year = self.year_dropdown.currentText().strip()
        selected_make = self.make_dropdown.currentText().strip()
        current_model = self.model_dropdown.currentText().strip()
        
        print(f"[DEBUG] update_model_dropdown: Year='{selected_year}', Make='{selected_make}', Current Model='{current_model}'")
        
        # Update dropdowns to show only available options based on current selections
        self.update_dropdowns_with_locks()
        
        # Update panels
        self.update_visible_panels()

    def handle_model_change(self, index):
        # Don't update if model is locked
        if self.model_locked:
            return
            
        selected_model = self.model_dropdown.currentText().strip()
        selected_make = self.make_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        
        print(f"[DEBUG] handle_model_change: Model='{selected_model}', Make='{selected_make}', Year='{selected_year}'")
        
        # Update dropdowns to show only available options based on current selections
        self.update_dropdowns_with_locks()
        
        # Update panels
        self.update_visible_panels()
        
        import logging
        logging.debug(f"Model selected: {selected_model}")
        
        # Update display based on current toggle state
        self.update_prequals_cmc_display()

    def populate_models(self, year_text, make_text):
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")
        
        print(f"[DEBUG] populate_models: Year: '{year_text}', Make: '{make_text}'")
        
        if year_text == "Select Year" or make_text == "Select Make" or make_text == "All":
            print(f"[DEBUG] populate_models: Early return - invalid selection")
            return
        try:
            year_int = int(year_text)
            matching_models = set()
            
            print(f"[DEBUG] populate_models: Searching for models with year {year_int} and make '{make_text}'")
            
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Year' in item and 'Make' in item and 'Model' in item and pd.notna(item['Year']) and item['Make'] == make_text):
                        item_year = int(float(item['Year']))
                        if item_year == year_int:
                            model_name = str(item['Model'])
                            matching_models.add(model_name)
                            if len(matching_models) <= 5:  # Show first 5 models being added
                                print(f"[DEBUG] populate_models: Adding model: '{model_name}' (type: {type(item['Model'])})")
                except (ValueError, TypeError, KeyError):
                    continue
            
            print(f"[DEBUG] populate_models: Found {len(matching_models)} matching models")
            
            if matching_models:
                sorted_models = sorted(matching_models)
                self.model_dropdown.addItems(sorted_models)
                print(f"[DEBUG] populate_models: Added models: {sorted_models[:5]}...")  # Show first 5
                import logging
                logging.info(f"Added {len(matching_models)} models for Year: {year_int}, Make: {make_text}")
            else:
                print(f"[DEBUG] populate_models: No models found!")
                import logging
                logging.warning(f"No models found for Year: {year_int}, Make: {make_text}")
        except (ValueError, TypeError) as e:
            print(f"[DEBUG] populate_models: Error: {e}")
            import logging
            logging.error(f"Error in populate_models: {e}")

    def clear_filters(self):
        """Clear all filters"""
        # Only clear dropdowns that are not locked
        if not self.year_locked:
            self.year_dropdown.setCurrentIndex(0)
        if not self.make_locked:
            self.make_dropdown.setCurrentIndex(0)
        if not self.model_locked:
            self.model_dropdown.setCurrentIndex(0)
            self.search_bar.clear()
            self.clear_display_panels()
            self.update_visible_panels()  # Update panels after clearing
            self.log_action(self.current_user, "Cleared all filters")
            self.log_action(self.current_user, "Cleared all filters")

    def toggle_always_on_top(self, checked):
        """Toggle always on top behavior"""
        if checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
            self.show()
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
            self.show()
    
    def update_font_size(self, size):
        """Update font size for all text display browsers"""
        # List of text browser attribute names
        browser_names = [
            'text_browser',
            'left_panel',
            'blacklist_panel_widget',
            'goldlist_panel_widget',
            'mag_glass_panel_widget',
            'results_display'
        ]
        
        # Update each text browser if it exists
        for name in browser_names:
            if hasattr(self, name):
                browser = getattr(self, name)
                if hasattr(browser, 'set_font_size'):
                    browser.set_font_size(size)
    



    def toggle_lock(self, field_type):
        """Toggle lock state for a field"""
        if field_type == 'year':
            self.year_locked = not self.year_locked
            if self.year_locked:
                self.locked_year = self.year_dropdown.currentText()
                self.year_lock_button.setText("🔒")
                self.year_lock_button.setStyleSheet(self.lock_button_style_locked)
                self.year_lock_button.setToolTip("Unlock Year selection")
            else:
                self.locked_year = None
                self.year_lock_button.setText("🔓")
                self.year_lock_button.setStyleSheet(self.lock_button_style)
        elif field_type == 'make':
            self.make_locked = not self.make_locked
            if self.make_locked:
                self.locked_make = self.make_dropdown.currentText()
                self.make_lock_button.setText("🔒")
                self.make_lock_button.setStyleSheet(self.lock_button_style_locked)
                self.make_lock_button.setToolTip("Unlock Make selection")
            else:
                self.locked_make = None
                self.make_lock_button.setText("🔓")
                self.make_lock_button.setStyleSheet(self.lock_button_style)
                self.make_lock_button.setToolTip("Lock Make selection")
        
        elif field_type == 'model':
            self.model_locked = not self.model_locked
            if self.model_locked:
                self.locked_model = self.model_dropdown.currentText()
                self.model_lock_button.setText("🔒")
                self.model_lock_button.setStyleSheet(self.lock_button_style_locked)
                self.model_lock_button.setToolTip("Unlock Model selection")
            else:
                self.locked_model = None
                self.model_lock_button.setText("🔓")
                self.model_lock_button.setStyleSheet(self.lock_button_style)
                self.model_lock_button.setToolTip("Lock Model selection")
        
        # Update other dropdowns based on locked values
        self.update_dropdowns_with_locks()
    def update_dropdowns_with_locks(self):
        """Update dropdowns while respecting locked values"""
        # Prevent recursive updates
        if getattr(self, '_updating_dropdowns', False):
            return
        self._updating_dropdowns = True
        try:
            # Get current values (use locked values if locked, otherwise use current dropdown text)
            year_to_use = self.locked_year if self.year_locked else self.year_dropdown.currentText()
            make_to_use = self.locked_make if self.make_locked else self.make_dropdown.currentText()
            model_to_use = self.locked_model if self.model_locked else self.model_dropdown.currentText()
            
            # Always update dropdowns to show only available options based on current selections
            # Update years dropdown based on current make and model selections
            if not self.year_locked:
                self.update_years_based_on_selections(make_to_use, model_to_use)
            
            # Update makes dropdown based on current year and model selections
            if not self.make_locked:
                self.update_makes_based_on_selections(year_to_use, model_to_use)
            
            # Update models dropdown based on current year and make selections
            if not self.model_locked:
                self.update_models_based_on_selections(year_to_use, make_to_use)
            
            # Simply disable/enable dropdowns based on lock state
            self.year_dropdown.setEnabled(not self.year_locked)
            self.make_dropdown.setEnabled(not self.make_locked)
            self.model_dropdown.setEnabled(not self.model_locked)
        finally:
            self._updating_dropdowns = False



    def perform_search(self):
        """Perform search based on current filters"""
        dtc_code = self.search_bar.text().strip()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        
        print(f"[DEBUG] Search criteria: Year={selected_year}, Make={selected_make}, Model={selected_model}, DTC={dtc_code}")
        
        # Print first 3 prequal records loaded
        if hasattr(self, 'data') and 'prequal' in self.data and self.data['prequal']:
            print(f"[DEBUG] First 3 prequal records loaded: {self.data['prequal'][:3]}")
        else:
            print("[DEBUG] No prequal data loaded.")
        
        # Update displays based on visible tabs
        self.update_displays_based_on_visible_tabs(dtc_code, selected_make, selected_model, selected_year)
        
        # Log search actions
        if dtc_code:
            self.log_action(self.current_user, f"DTC search performed: {dtc_code}")
        if selected_year != "Select Year" and selected_make != "Select Make" and selected_model != "Select Model":
            self.log_action(self.current_user, f"Vehicle search performed: {selected_year} {selected_make} {selected_model}")

    def clear_display_panels(self):
        """Clear all display panels"""
        self.left_panel.clear()
        self.blacklist_panel_widget.clear()
        self.goldlist_panel_widget.clear()
        self.mag_glass_panel_widget.clear()



    def update_displays_based_on_visible_tabs(self, dtc_code, selected_make, selected_model, selected_year):
        """Update displays based on which tabs are currently visible"""
        # Check which tabs are visible and update their content accordingly
        if self.prequals_panel.isVisible():
            self.update_prequals_cmc_display()
        if self.blacklist_panel.isVisible():
            if dtc_code:
                self.search_blacklist_dtc(dtc_code, selected_make)
            else:
                self.display_blacklist(selected_make)
        if self.goldlist_panel.isVisible():
            if dtc_code:
                self.search_goldlist_dtc(dtc_code, selected_make)
            else:
                self.display_goldlist(selected_make)
        if self.mag_glass_panel.isVisible():
            self.display_mag_glass(selected_make)

    def handle_prequal_search(self, selected_make, selected_model, selected_year):
        """Handle prequalification search"""
        print(f"[DEBUG] handle_prequal_search: Make='{selected_make}', Model='{selected_model}', Year='{selected_year}'")
        
        # Dictionary to hold unique System Acronyms
        unique_results = {}

        # Convert selected_model to a string
        selected_model_str = str(selected_model)

        # Debug: Show sample data for the selected make and year
        sample_data = [item for item in self.data['prequal'] 
                      if item.get('Make') == selected_make and str(item.get('Year', '')) == str(selected_year)]
        if sample_data:
            print(f"[DEBUG] handle_prequal_search: Found {len(sample_data)} records for {selected_make} {selected_year}")
            print(f"[DEBUG] handle_prequal_search: Sample models in data: {[str(item.get('Model', '')) for item in sample_data[:5]]}")
            print(f"[DEBUG] handle_prequal_search: Looking for model: '{selected_model_str}'")
        else:
            print(f"[DEBUG] handle_prequal_search: No data found for {selected_make} {selected_year}")

        # Filtering data based on selections
        filtered_results = [item for item in self.data['prequal']
                            if (selected_make == "All" or item['Make'] == selected_make) and
                            (selected_model == "Select Model" or str(item['Model']) == selected_model_str) and
                            (selected_year == "Select Year" or str(int(float(item['Year']))) == selected_year)]

        print(f"[DEBUG] handle_prequal_search: Filtered results: {len(filtered_results)}")
        
        if filtered_results:
            print(f"[DEBUG] handle_prequal_search: Sample filtered result: {filtered_results[0]}")
        else:
            # Debug: Show why filtering failed
            print(f"[DEBUG] handle_prequal_search: No results found. Checking each condition:")
            make_matches = [item for item in self.data['prequal'] if item.get('Make') == selected_make]
            print(f"[DEBUG] handle_prequal_search: Make matches: {len(make_matches)}")
            
            if make_matches:
                year_matches = [item for item in make_matches if str(item.get('Year', '')) == str(selected_year)]
                print(f"[DEBUG] handle_prequal_search: Year matches: {len(year_matches)}")
                
                if year_matches:
                    model_matches = [item for item in year_matches if str(item.get('Model', '')) == selected_model_str]
                    print(f"[DEBUG] handle_prequal_search: Model matches: {len(model_matches)}")
                    if not model_matches:
                        print(f"[DEBUG] handle_prequal_search: Model comparison failed. Available models: {list(set([str(item.get('Model', '')) for item in year_matches]))}")

        # Populating the dictionary with unique entries based on System Acronym
        for item in filtered_results:
            system_acronym = item.get('Protech Generic System Name.1', 'N/A')  # Match original app column name
            if system_acronym not in unique_results:
                unique_results[system_acronym] = item  # Store the entire item for display

        print(f"[DEBUG] handle_prequal_search: Unique results: {len(unique_results)}")

        # Now, pass the unique items to be displayed
        self.display_results(list(unique_results.values()), context='prequal')



    def display_cmc_data(self, selected_year, selected_make, selected_model):
        """Display Manufacturer Chart data"""
        conn = None
        try:
            # Check if year, make, and model are selected
            if (selected_year == "Select Year" or selected_make == "Select Make" or 
                selected_model == "Select Model"):
                self.left_panel.setPlainText("Please select Year, Make, and Model to view Manufacturer Chart data.")
                return
            
            conn = self.get_db_connection()
            
            # Check if table exists
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='manufacturer_chart'")
            if not cursor.fetchone():
                self.left_panel.setPlainText("No Manufacturer Chart data found. Please load data first.")
                return
            
            # Query for manufacturer chart data with year, make, and model filters
            # First, let's see what columns are actually in the table
            cursor.execute("PRAGMA table_info(manufacturer_chart)")
            columns_info = cursor.fetchall()
            logging.debug(f"Manufacturer chart table columns: {[col[1] for col in columns_info]}")
            
            # Get all data first to see what we have
            query = "SELECT * FROM manufacturer_chart"
            df = pd.read_sql_query(query, conn)
            
            # Debug: Check for Acura data specifically
            acura_query = "SELECT DISTINCT Make FROM manufacturer_chart WHERE UPPER(Make) LIKE '%ACURA%'"
            acura_cursor = conn.cursor()
            acura_cursor.execute(acura_query)
            acura_makes = acura_cursor.fetchall()
            logging.debug(f"Acura makes found in database: {acura_makes}")
            
            # Check for any data with 'acura' in the make field
            all_makes_query = "SELECT DISTINCT Make FROM manufacturer_chart WHERE Make IS NOT NULL ORDER BY Make"
            all_makes_cursor = conn.cursor()
            all_makes_cursor.execute(all_makes_query)
            all_makes_in_db = [row[0] for row in all_makes_cursor.fetchall()]
            logging.debug(f"All makes in database: {all_makes_in_db[:20]}")  # Show first 20
            
            logging.debug(f"Total manufacturer chart records: {len(df)}")
            if not df.empty:
                logging.debug(f"Sample manufacturer chart data: {df.head()}")
                logging.debug(f"Available columns: {list(df.columns)}")
            
            # Now filter by year, make, and model
            if not df.empty:
                # Normalize the search values
                search_year = str(selected_year).strip().upper()
                search_make = str(selected_make).strip().upper()
                search_model = str(selected_model).strip().upper()
                
                # Debug: Show what data is available
                logging.debug(f"Searching for: Year='{search_year}', Make='{search_make}', Model='{search_model}'")
                
                # Show sample of available data - handle None values properly
                unique_years = df['Year'].dropna().unique()
                unique_makes = df['Make'].dropna().unique()
                unique_models = df['Model'].dropna().unique()
                logging.debug(f"Available years: {list(unique_years)[:10]}")  # Show first 10
                all_makes = sorted([make for make in unique_makes if make is not None])
                logging.debug(f"Available makes (first 20): {all_makes[:20]}")
                logging.debug(f"Total unique makes: {len(all_makes)}")
                # Check specifically for Acura
                acura_present = 'Acura' in all_makes
                logging.debug(f"Is 'Acura' in the data: {acura_present}")
                if acura_present:
                    acura_data = df[df['Make'].fillna('').str.upper() == 'ACURA']
                    logging.debug(f"Acura records found: {len(acura_data)}")
                    if len(acura_data) > 0:
                        logging.debug(f"Acura years: {acura_data['Year'].dropna().unique().tolist()}")
                        logging.debug(f"Acura models: {acura_data['Model'].dropna().unique().tolist()}")
                else:
                    # Check if there are any variations of Acura
                    acura_variants = [make for make in all_makes if make and 'acura' in make.lower()]
                    logging.debug(f"Acura variants found: {acura_variants}")
                logging.debug(f"Available models: {list(unique_models)[:10]}")  # Show first 10
                
                # Filter the dataframe - handle None values and float years properly
                # Convert Year to int first to remove .0, then to string
                def normalize_year(val):
                    if pd.isna(val) or val == '':
                        return ''
                    try:
                        # Try to convert to float first, then to int to remove decimals
                        return str(int(float(val)))
                    except:
                        return str(val).strip()
                
                filtered_df = df[
                    (df['Year'].apply(normalize_year).str.upper() == search_year) &
                    (df['Make'].fillna('').astype(str).str.strip().str.upper() == search_make) &
                    (df['Model'].fillna('').astype(str).str.strip().str.upper() == search_model)
                ]
                
                logging.debug(f"Filtered records for {selected_year} {selected_make} {selected_model}: {len(filtered_df)}")
                df = filtered_df
            
            # Replace NaN values with empty strings
            df.fillna("", inplace=True)
            
            # Display the results in the left panel
            if df.empty:
                self.left_panel.setPlainText(f"No Manufacturer Chart data found for {selected_year} {selected_make} {selected_model}")
            else:
                # Format the data similar to prequals display
                html_content = self.format_cmc_data_for_display(df)
                self.left_panel.setHtml(html_content)
                self.left_panel.setOpenExternalLinks(True)
        
        except Exception as e:
            logging.error(f"Failed to display CMC data: {e}")
            self.left_panel.setPlainText(f"An error occurred while fetching the Manufacturer Chart data: {str(e)}")
        
        finally:
            if conn:
                conn.close()

    def format_cmc_data_for_display(self, df):
        """Format CMC data for display similar to prequals format"""
        html_content = """
        <html>
        <head>
        <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 9px; line-height: 1.4; color: #333; }
        .record { 
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); 
            border: 2px solid #667eea; 
            border-radius: 10px; 
            padding: 12px; 
            margin-bottom: 15px; 
            box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
            position: relative; 
            overflow: hidden; 
        }
        .record-header {
            background: #667eea;
            color: white;
            padding: 6px 10px;
            margin: -12px -12px 8px -12px;
            border-radius: 8px 8px 0 0;
            font-weight: bold;
            font-size: 11px;
        }
        .field { margin-bottom: 2px; }
        .field-label { 
            font-weight: bold; 
            color: #20567C; 
            display: inline-block; 
            width: 180px; 
            font-size: 9px;
        }
        .field-value { color: #333; font-size: 9px; }
        .service-link { color: #0066cc; text-decoration: none; }
        .service-link:hover { text-decoration: underline; }
        </style>
        </head>
        <body>
        """
        
        for index, row in df.iterrows():
            # Get the feature name for the header
            feature = row.get('Feature', 'Unknown Feature')
            if pd.isna(feature) or feature == '':
                feature = 'Unknown Feature'
            
            html_content += f"""
            <div class="record">
                <div class="record-header">{feature}</div>
            """
            
            # Display all available columns dynamically
            for col in df.columns:
                if col.lower() not in ['id', 'created_at']:  # Skip internal columns
                    value = row.get(col, 'N/A')
                    if pd.isna(value) or value == '':
                        value = 'N/A'
                    
                    # Handle special cases
                    if 'hyperlink' in col.lower() or 'link' in col.lower():
                        if value and value != 'N/A' and value.startswith(('http://', 'https://')):
                            html_content += f"""
                <div class="field">
                    <span class="field-label">{col.replace('_', ' ').title()}:</span>
                    <span class="field-value"><a href="{value}" class="service-link" target="_blank">View Service Information</a></span>
                </div>
                            """
                        else:
                            html_content += f"""
                <div class="field">
                    <span class="field-label">{col.replace('_', ' ').title()}:</span>
                    <span class="field-value">N/A</span>
                </div>
                            """
                    elif 'prerequisites' in col.lower() or 'requirements' in col.lower():
                        if value and value != 'N/A':
                            # Convert line breaks to HTML breaks
                            value = str(value).replace('\n', '<br>')
                            html_content += f"""
                <div class="field">
                    <span class="field-label">{col.replace('_', ' ').title()}:</span>
                    <span class="field-value">{value}</span>
                </div>
                            """
                        else:
                            html_content += f"""
                <div class="field">
                    <span class="field-label">{col.replace('_', ' ').title()}:</span>
                    <span class="field-value">N/A</span>
                </div>
                            """
                    else:
                        html_content += f"""
                <div class="field">
                    <span class="field-label">{col.replace('_', ' ').title()}:</span>
                    <span class="field-value">{value}</span>
                </div>
                        """
            
            html_content += """
            </div>
            """
        
        html_content += """
        </body>
        </html>
        """
        
        return html_content

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
    
    def format_side_by_side_data(self, data, data_type):
        """Format data for side-by-side display"""
        if not data:
            return '<span class="no-data">No data</span>'
        
        html = ""
        
        if data_type == 'prequal':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 15px; padding: 12px; background: #fff; border-radius: 6px; border-left: 4px solid #1976d2;">
                        <strong style="color: #1976d2;">Record {i+1}</strong><br/>
                        <strong>Component:</strong> {item.get('Parent Component', 'N/A')}<br/>
                        <strong>System:</strong> {item.get('Protech Generic System Name', 'N/A')}<br/>
                        <strong>Calibration Type:</strong> {item.get('Calibration Type', 'N/A')}<br/>
                        <strong>Prerequisites:</strong> {item.get('Calibration Pre-Requisites (Short Hand)', 'N/A')}<br/>
                        <strong>Point of Impact:</strong> {item.get('Point of Impact #', 'N/A')}
                    </div>
                """
        elif data_type in ['blacklist', 'goldlist']:
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 15px; padding: 12px; background: #fff; border-radius: 6px; border-left: 4px solid #1976d2;">
                        <strong style="color: #1976d2;">Record {i+1}</strong><br/>
                        <strong>DTC Code:</strong> {item.get('dtcCode', 'N/A')}<br/>
                        <strong>Description:</strong> {item.get('dtcDescription', 'N/A')}<br/>
                        <strong>Make:</strong> {item.get('carMake', 'N/A')}
                    </div>
                """
        elif data_type == 'mag_glass':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 15px; padding: 12px; background: #fff; border-radius: 6px; border-left: 4px solid #1976d2;">
                        <strong style="color: #1976d2;">Record {i+1}</strong><br/>
                        <strong>Make:</strong> {item.get('Make', 'N/A')}<br/>
                        <strong>Model:</strong> {item.get('Model', 'N/A')}<br/>
                        <strong>Year:</strong> {item.get('Year', 'N/A')}
                    </div>
                """
        elif data_type == 'carsys':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 15px; padding: 12px; background: #fff; border-radius: 6px; border-left: 4px solid #1976d2;">
                        <strong style="color: #1976d2;">Record {i+1}</strong><br/>
                        <strong>Make:</strong> {item.get('Make', 'N/A')}<br/>
                        <strong>Model:</strong> {item.get('Model', 'N/A')}<br/>
                        <strong>System:</strong> {item.get('System', 'N/A')}
        </div>
        """
        
        return html

    def clear_search_bar(self):
        """Clear search bar"""
        self.search_bar.clear()

    def change_opacity(self, value):
        """Change window opacity"""
        self.setWindowOpacity(value / 100.0)

    def open_admin(self):
        """Open admin panel"""
        dialog = ManageDataListsDialog(self)
        dialog.exec_()
    
    def open_compare_dialog(self):
        """Open multi-vehicle comparison dialog"""
        dialog = MultiVehicleCompareDialog(self)
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
                            df = pd.read_excel(filepath, sheet_name=0)
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
        
        # Load prequal data using the new database utilities
        prequal_data = get_prequal_data(self.db_path)
        self.data['prequal'] = prequal_data if prequal_data else []
        logging.debug(f"Loaded {len(self.data['prequal'])} items for prequal")
        
        # Load other data types
        for config_type in ['blacklist', 'goldlist', 'mag_glass', 'carsys']:
            data = load_configuration(config_type, self.db_path)
            self.data[config_type] = data if data else []
            logging.debug(f"Loaded {len(data)} items for {config_type}")
        
        if self.data['prequal']:
            self.populate_dropdowns()
        else:
            logging.warning("No prequal data found!")

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
                cursor.execute("DROP TABLE IF EXISTS manufacturer_chart")
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
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'CarSys', 'manufacturer_chart']:
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
                        elif config_type == 'manufacturer_chart':
                            # For manufacturer chart, we need to load ALL files in the directory
                            # This is different from other data types that process one file at a time
                            continue  # Skip individual file processing - we'll handle all files at once
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
        
        # Check if manufacturer_chart data exists in database (don't reload from files)
        manufacturer_chart_path = load_path_from_db('manufacturer_chart', self.db_path)
        if manufacturer_chart_path:
            try:
                # Check if we have data in the database
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM manufacturer_chart")
                count = cursor.fetchone()[0]
                conn.close()
                
                if count > 0:
                    logging.info(f"Manufacturer chart data already in database ({count} records)")
                    any_data_loaded = True
                else:
                    logging.warning("No manufacturer chart data in database. Please use 'Save and Load' to load data.")
                    
            except Exception as e:
                logging.error(f"Error checking manufacturer chart data: {str(e)}")
        
        if any_data_loaded:
            msg = self.create_styled_messagebox("Success", "All data refreshed successfully!", QMessageBox.Information)
            msg.exec_()
        # Always repopulate dropdowns and check enabled state after refresh
        self.populate_dropdowns()
        self.check_data_loaded()

    def load_manufacturer_chart_data(self, filepath, db_path):
        """Load manufacturer chart data from Excel file with specified columns"""
        try:
            # Read Excel file
            df = pd.read_excel(filepath)
            
            if df.empty:
                return "No data found in file"
            
            # Normalize column names (case-insensitive, remove special characters)
            df.columns = [normalize_col(col) for col in df.columns]
            
            # Define expected column names (normalized)
            expected_columns = [
                'year', 'make', 'model', 'calibrationtype', 'protechgenericsystemname',
                'smegenericsystemname', 'smecalibrationtype', 'feature', 
                'serviceinformationhyperlink', 'calibrationprerequisites'
            ]
            
            # Check if required columns exist (with variations)
            missing_columns = []
            for expected_col in expected_columns:
                if not any(expected_col in col.lower() for col in df.columns):
                    missing_columns.append(expected_col)
            
            if missing_columns:
                logging.warning(f"Missing columns in manufacturer chart: {missing_columns}")
                # Continue with available columns
            
            # Create table if it doesn't exist
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Create manufacturer_chart table with all possible columns
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS manufacturer_chart (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year TEXT,
                    Make TEXT,
                    Model TEXT,
                    Calibration_Type TEXT,
                    Protech_Generic_System_Name TEXT,
                    SME_Generic_System_Name TEXT,
                    SME_Calibration_Type TEXT,
                    Feature TEXT,
                    Service_Information_Hyperlink TEXT,
                    Calibration_Pre_Requisites TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Clear existing data
            cursor.execute("DELETE FROM manufacturer_chart")
            
            # Map normalized columns to database columns
            column_mapping = {
                'year': 'Year',
                'make': 'Make', 
                'model': 'Model',
                'calibrationtype': 'Calibration_Type',
                'protechgenericsystemname': 'Protech_Generic_System_Name',
                'smegenericsystemname': 'SME_Generic_System_Name',
                'smecalibrationtype': 'SME_Calibration_Type',
                'feature': 'Feature',
                'serviceinformationhyperlink': 'Service_Information_Hyperlink',
                'calibrationprerequisites': 'Calibration_Pre_Requisites'
            }
            
            # Log normalized column names for debugging
            logging.info(f"Normalized DataFrame columns: {df.columns.tolist()}")
            
            # Prepare data for insertion
            records = []
            for _, row in df.iterrows():
                record = {}
                for normalized_col, db_col in column_mapping.items():
                    # Since columns are already normalized, do exact match
                    if normalized_col in df.columns:
                        value = row[normalized_col]
                        if pd.isna(value):
                            record[db_col] = ''  # Use empty string for consistency
                        else:
                            record[db_col] = str(value).strip()
                    else:
                        record[db_col] = ''
                        logging.warning(f"Column '{normalized_col}' not found in DataFrame")
                
                records.append(record)
            
            # Log first record for debugging
            if records:
                logging.info(f"First CMC record to insert: {records[0]}")
            
            # Insert data
            if records:
                placeholders = ', '.join(['?' for _ in column_mapping.values()])
                columns = ', '.join(column_mapping.values())
                cursor.executemany(
                    f"INSERT INTO manufacturer_chart ({columns}) VALUES ({placeholders})",
                    [tuple(record.values()) for record in records]
                )
                
                conn.commit()
                conn.close()
                logging.info(f"Successfully inserted {len(records)} CMC records")
                return "Data loaded successfully"
            else:
                conn.close()
                return "No valid records found"
                
        except Exception as e:
            logging.error(f"Error loading manufacturer chart data: {str(e)}")
            return f"Error: {str(e)}"

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
        """Populate dropdowns with data from both prequal and manufacturer chart"""
        # Get data from both sources
        prequal_years = []
        prequal_makes = []
        manufacturer_years = []
        manufacturer_makes = []
        
        # Get prequal data if available
        if self.data['prequal']:
            prequal_years = get_unique_years(self.data['prequal'])
            prequal_makes = get_unique_makes(self.data['prequal'])
            logging.debug(f"Prequal - Found years: {prequal_years}")
            logging.debug(f"Prequal - Found makes: {prequal_makes}")
        
        # Get manufacturer chart data if available
        try:
            from database_utils import get_unique_years_from_manufacturer_chart, get_unique_makes_from_manufacturer_chart
            manufacturer_years = get_unique_years_from_manufacturer_chart(self.db_path)
            manufacturer_makes = get_unique_makes_from_manufacturer_chart(self.db_path)
            logging.debug(f"Manufacturer Chart - Found years: {manufacturer_years}")
            logging.debug(f"Manufacturer Chart - Found makes: {manufacturer_makes}")
        except Exception as e:
            logging.error(f"Error getting manufacturer chart data: {e}")
        
        # Combine and deduplicate years and makes
        all_years = list(set(prequal_years + manufacturer_years))
        all_makes = list(set(prequal_makes + manufacturer_makes))
        
        # Sort years in descending order (newest first)
        # Handle float years like '2017.0' by converting to int first
        def year_sort_key(year_str):
            try:
                return int(float(year_str))
            except (ValueError, TypeError):
                return 0
        
        all_years.sort(key=year_sort_key, reverse=True)
        # Sort makes alphabetically
        all_makes.sort()
        
        logging.debug(f"Combined - Found years: {all_years}")
        logging.debug(f"Combined - Found makes: {all_makes}")
        
        # Clear and populate year dropdown
        self.year_dropdown.clear()
        self.year_dropdown.addItem("Select Year")
        for year in all_years:
            self.year_dropdown.addItem(year)
            
        # Clear and populate make dropdown
        self.make_dropdown.clear()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")  # Keep the "All" option
        for make in all_makes:
            self.make_dropdown.addItem(make)
            
        # Clear model dropdown
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")

    def has_valid_prequal(self, item):
        """Check if an item has valid prequal data"""
        if not isinstance(item, dict):
            return False
            
        # Check required fields
        required_fields = ['Year', 'Make', 'Model']
        for field in required_fields:
            if field not in item or not item[field] or str(item[field]).lower() == 'nan':
                return False
                
        # Check if year is valid
        try:
            year = int(float(item['Year']))
            if year < 1900 or year > 2100:  # Reasonable year range
                return False
        except (ValueError, TypeError):
            return False
            
        # Check if make and model are valid strings
        if not isinstance(item['Make'], str) or not item['Make'].strip():
            return False
        if not isinstance(item['Model'], str) or not item['Model'].strip():
            return False
            
        # Check if make is not unknown
        if item['Make'].lower() in ['unknown', 'nan']:
            return False
            
        return True

    def search_blacklist_dtc(self, dtc_code, selected_make):
        """Search blacklist DTC codes"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='blacklist'")
            if not cursor.fetchone():
                self.blacklist_panel_widget.setPlainText("No blacklist data found. Please load data first.")
                return
            
            # Query blacklist data with DTC code search
            if selected_make == "All":
                query = f"""
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                WHERE dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%'
                """
            else:
                query = f"""
                SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%') AND carMake = '{selected_make}'
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
                self.blacklist_panel_widget.setPlainText(f"No blacklist results found for DTC code: {dtc_code}")
        except Exception as e:
            logging.error(f"Failed to execute blacklist search query: {query}\nError: {e}")
            self.blacklist_panel_widget.setPlainText(f"An error occurred while searching blacklist data: {str(e)}")
        finally:
            if conn:
                conn.close()

    def search_goldlist_dtc(self, dtc_code, selected_make):
        """Search goldlist DTC codes"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='goldlist'")
            if not cursor.fetchone():
                self.goldlist_panel_widget.setPlainText("No goldlist data found. Please load data first.")
                return
            
            # Query goldlist data with DTC code search
            if selected_make == "All":
                query = f"""
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
                WHERE dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%'
                """
            else:
                query = f"""
                SELECT 'goldlist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM goldlist
                WHERE (dtcCode LIKE '%{dtc_code}%' OR dtcDescription LIKE '%{dtc_code}%') AND carMake = '{selected_make}'
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
                self.goldlist_panel_widget.setPlainText(f"No goldlist results found for DTC code: {dtc_code}")
        except Exception as e:
            logging.error(f"Failed to execute goldlist search query: {query}\nError: {e}")
            self.goldlist_panel_widget.setPlainText(f"An error occurred while searching goldlist data: {str(e)}")
        finally:
            if conn:
                conn.close()

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

    def update_visible_panels(self):
        dtc_code = self.search_bar.text().strip()
        for attr, btn in self.tab_buttons.items():
            if btn.isChecked():
                if attr == "prequals_panel":
                    # Update display based on current toggle state
                    self.update_prequals_cmc_display()
                elif attr == "blacklist_panel":
                    if dtc_code:
                        self.search_blacklist_dtc(dtc_code, self.make_dropdown.currentText())
                    else:
                        self.display_blacklist(self.make_dropdown.currentText())
                elif attr == "goldlist_panel":
                    if dtc_code:
                        self.search_goldlist_dtc(dtc_code, self.make_dropdown.currentText())
                    else:
                        self.display_goldlist(self.make_dropdown.currentText())
                elif attr == "mag_glass_panel":
                    self.display_mag_glass(self.make_dropdown.currentText())
                elif attr == "cmc_panel":
                    self.display_cmc_data(self.year_dropdown.currentText(), self.make_dropdown.currentText(), self.model_dropdown.currentText())

    def update_years_for_locked_model(self, locked_model):
        """Update year dropdown to only years that contain the locked model."""
        try:
            valid_years = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Model' in item and 'Year' in item and 
                        pd.notna(item.get('Year')) and str(item.get('Model', '')).strip() == str(locked_model).strip()):
                        valid_years.add(int(float(item['Year'])))
                except (ValueError, TypeError, KeyError):
                    continue
            valid_years = sorted(valid_years, reverse=True)

            # Only rebuild if year isn't locked
            if not self.year_locked:
                current_year = self.year_dropdown.currentText()
                self.year_dropdown.blockSignals(True)
                self.year_dropdown.clear()
                self.year_dropdown.addItem("Select Year")
                self.year_dropdown.addItems([str(y) for y in valid_years])
                self.year_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_year in [str(y) for y in valid_years]:
                    idx = self.year_dropdown.findText(current_year)
                    if idx != -1:
                        self.year_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating years for locked model '{locked_model}': {e}")

    def update_makes_for_locked_model(self, locked_model):
        """Update make dropdown to only makes that contain the locked model."""
        try:
            valid_makes = set()
            for item in self.data['prequal']:
                    try:
                        if (self.has_valid_prequal(item) and 'Model' in item and 'Make' in item and 
                        str(item.get('Model', '')).strip() == str(locked_model).strip() and 
                        item.get('Make', '').strip() and item.get('Make', '').strip().lower() != 'unknown'):
                            valid_makes.add(item['Make'].strip())
                    except (ValueError, TypeError, KeyError):
                        continue
            valid_makes = sorted(valid_makes)

            # Only rebuild if make isn't locked
            if not self.make_locked:
                current_make = self.make_dropdown.currentText()
                self.make_dropdown.blockSignals(True)
                self.make_dropdown.clear()
                self.make_dropdown.addItem("Select Make")
                self.make_dropdown.addItem("All")
                self.make_dropdown.addItems(valid_makes)
                self.make_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_make in valid_makes:
                    idx = self.make_dropdown.findText(current_make)
                    if idx != -1:
                        self.make_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating makes for locked model '{locked_model}': {e}")

    def update_makes_for_locked_year(self, locked_year):
        """Update makes dropdown to only makes available for the locked year."""
        try:
            year_int = int(locked_year)
            valid_makes = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Year' in item and 'Make' in item and 
                        pd.notna(item.get('Year')) and int(float(item['Year'])) == year_int and
                        item.get('Make', '').strip() and item.get('Make', '').strip().lower() != 'unknown'):
                        valid_makes.add(item['Make'].strip())
                except (ValueError, TypeError, KeyError):
                        continue
            valid_makes = sorted(valid_makes)

            # Only rebuild if make isn't locked
            if not self.make_locked:
                current_make = self.make_dropdown.currentText()
                self.make_dropdown.blockSignals(True)
                self.make_dropdown.clear()
                self.make_dropdown.addItem("Select Make")
                self.make_dropdown.addItem("All")
                self.make_dropdown.addItems(valid_makes)
                self.make_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_make in valid_makes:
                    idx = self.make_dropdown.findText(current_make)
                    if idx != -1:
                        self.make_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating makes for locked year '{locked_year}': {e}")

    def update_models_for_locked_year(self, locked_year):
        """Update models dropdown to only models available for the locked year."""
        try:
            year_int = int(locked_year)
            valid_models = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Year' in item and 'Model' in item and 
                        pd.notna(item.get('Year')) and int(float(item['Year'])) == year_int):
                        valid_models.add(str(item['Model']))
                except (ValueError, TypeError, KeyError):
                    continue
            valid_models = sorted(valid_models)

            # Only rebuild if model isn't locked
            if not self.model_locked:
                current_model = self.model_dropdown.currentText()
                self.model_dropdown.blockSignals(True)
                self.model_dropdown.clear()
                self.model_dropdown.addItem("Select Model")
                self.model_dropdown.addItems(valid_models)
                self.model_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_model in valid_models:
                    idx = self.model_dropdown.findText(current_model)
                    if idx != -1:
                        self.model_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating models for locked year '{locked_year}': {e}")

    def update_years_for_locked_make(self, locked_make):
        """Update year dropdown to only years available for the locked make."""
        try:
            valid_years = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Make' in item and 'Year' in item and 
                        pd.notna(item.get('Year')) and str(item.get('Make', '')).strip() == str(locked_make).strip()):
                        valid_years.add(int(float(item['Year'])))
                except (ValueError, TypeError, KeyError):
                    continue
            valid_years = sorted(valid_years, reverse=True)

            # Only rebuild if year isn't locked
            if not self.year_locked:
                current_year = self.year_dropdown.currentText()
                self.year_dropdown.blockSignals(True)
                self.year_dropdown.clear()
                self.year_dropdown.addItem("Select Year")
                self.year_dropdown.addItems([str(y) for y in valid_years])
                self.year_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_year in [str(y) for y in valid_years]:
                    idx = self.year_dropdown.findText(current_year)
                    if idx != -1:
                        self.year_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating years for locked make '{locked_make}': {e}")

    def update_models_for_locked_make(self, locked_make):
        """Update models dropdown to only models available for the locked make."""
        try:
            valid_models = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 'Make' in item and 'Model' in item and 
                        str(item.get('Make', '')).strip() == str(locked_make).strip()):
                        valid_models.add(str(item['Model']))
                except (ValueError, TypeError, KeyError):
                    continue
            valid_models = sorted(valid_models)

            # Only rebuild if model isn't locked
            if not self.model_locked:
                current_model = self.model_dropdown.currentText()
                self.model_dropdown.blockSignals(True)
                self.model_dropdown.clear()
                self.model_dropdown.addItem("Select Model")
                self.model_dropdown.addItems(valid_models)
                self.model_dropdown.blockSignals(False)
                
                # Restore selection if still present
                if current_model in valid_models:
                    idx = self.model_dropdown.findText(current_model)
                    if idx != -1:
                        self.model_dropdown.setCurrentIndex(idx)
        except Exception as e:
            logging.error(f"Error updating models for locked make '{locked_make}': {e}")

    def update_years_based_on_selections(self, make_to_use, model_to_use):
        """Update year dropdown to show only years that match current make and model selections."""
        try:
            valid_years = set()
            for item in self.data['prequal']:
                try:
                    if not self.has_valid_prequal(item) or pd.isna(item.get('Year')):
                        continue
                    
                    # Check if this item matches our current selections
                    make_matches = (make_to_use in ["Select Make", "All", ""] or 
                                  str(item.get('Make', '')).strip() == str(make_to_use).strip())
                    model_matches = (model_to_use in ["Select Model", ""] or 
                                   str(item.get('Model', '')).strip() == str(model_to_use).strip())
                    
                    if make_matches and model_matches:
                        # Apply region filtering - only include years for makes in current region
                        make = item.get('Make', '').strip()
                        if self.is_make_in_current_region(make):
                            valid_years.add(int(float(item['Year'])))
                except (ValueError, TypeError, KeyError):
                    continue
            
            # Get years from manufacturer chart data
            try:
                if make_to_use not in ["Select Make", "All", ""]:
                    # Get all years for the selected make from manufacturer chart
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    if model_to_use not in ["Select Model", ""]:
                        # If model is selected, get years for that specific make and model
                        cursor.execute("SELECT DISTINCT Year FROM manufacturer_chart WHERE Make = ? AND Model = ? AND Year IS NOT NULL AND Year != '' ORDER BY Year DESC", (make_to_use, model_to_use))
                    else:
                        # If only make is selected, get all years for that make
                        cursor.execute("SELECT DISTINCT Year FROM manufacturer_chart WHERE Make = ? AND Year IS NOT NULL AND Year != '' ORDER BY Year DESC", (make_to_use,))
                    
                    manufacturer_years = cursor.fetchall()
                    conn.close()
                    
                    for year_row in manufacturer_years:
                        try:
                            year_str = str(year_row[0])
                            year_int = int(float(year_str))
                            valid_years.add(year_int)
                        except (ValueError, TypeError):
                            continue
            except Exception as e:
                logging.error(f"Error getting manufacturer chart years: {e}")
            
            valid_years = sorted(valid_years, reverse=True)
            
            # Store current selection
            current_year = self.year_dropdown.currentText()
            
            # Rebuild dropdown
            self.year_dropdown.blockSignals(True)
            self.year_dropdown.clear()
            self.year_dropdown.addItem("Select Year")
            for year in valid_years:
                self.year_dropdown.addItem(str(year))
            self.year_dropdown.blockSignals(False)
            
            # Try to restore selection
            if current_year != "Select Year":
                index = self.year_dropdown.findText(current_year)
                if index >= 0:
                    self.year_dropdown.setCurrentIndex(index)
                    
        except Exception as e:
            print(f"Error updating years based on selections: {e}")
    
    def update_makes_based_on_selections(self, year_to_use, model_to_use):
        """Update make dropdown to show only makes that match current year and model selections."""
        try:
            valid_makes = set()
            for item in self.data['prequal']:
                try:
                    if not self.has_valid_prequal(item):
                        continue
                    
                    # Check if this item matches our current selections
                    # Normalize year comparison to handle float values (e.g., 2024.0 vs 2024)
                    year_matches = year_to_use in ["Select Year", ""]
                    if not year_matches and 'Year' in item and pd.notna(item['Year']):
                        try:
                            item_year = int(float(item['Year']))
                            selected_year = int(year_to_use)
                            year_matches = (item_year == selected_year)
                        except (ValueError, TypeError):
                            pass
                    
                    model_matches = (model_to_use in ["Select Model", ""] or 
                                   str(item.get('Model', '')).strip() == str(model_to_use).strip())
                    
                    if year_matches and model_matches and 'Make' in item and item['Make'].strip():
                        make = item['Make'].strip()
                        # Apply region filtering
                        if self.is_make_in_current_region(make):
                            valid_makes.add(make)
                except (ValueError, TypeError, KeyError):
                    continue
            
            # Get makes from manufacturer chart data
            try:
                if year_to_use not in ["Select Year", ""]:
                    # Get all makes for the selected year from manufacturer chart
                    # Convert year to float format for database comparison (e.g., "2021" -> "2021.0")
                    year_float = f"{float(year_to_use)}"
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    if model_to_use not in ["Select Model", ""]:
                        # If model is selected, get makes for that specific year and model
                        cursor.execute("SELECT DISTINCT Make FROM manufacturer_chart WHERE Year = ? AND Model = ? AND Make IS NOT NULL AND Make != '' ORDER BY Make", (year_float, model_to_use))
                    else:
                        # If only year is selected, get all makes for that year
                        cursor.execute("SELECT DISTINCT Make FROM manufacturer_chart WHERE Year = ? AND Make IS NOT NULL AND Make != '' ORDER BY Make", (year_float,))
                    
                    manufacturer_makes = cursor.fetchall()
                    conn.close()
                    
                    for make_row in manufacturer_makes:
                        make = make_row[0].strip()
                        if make and self.is_make_in_current_region(make):
                            valid_makes.add(make)
            except Exception as e:
                logging.error(f"Error getting manufacturer chart makes: {e}")
            
            valid_makes = sorted(valid_makes)
            
            # Store current selection
            current_make = self.make_dropdown.currentText()
            
            # Rebuild dropdown
            self.make_dropdown.blockSignals(True)
            self.make_dropdown.clear()
            self.make_dropdown.addItem("Select Make")
            for make in valid_makes:
                self.make_dropdown.addItem(make)
            self.make_dropdown.blockSignals(False)
            
            # Try to restore selection
            if current_make not in ["Select Make", "All"]:
                    index = self.make_dropdown.findText(current_make)
                    if index >= 0:
                        self.make_dropdown.setCurrentIndex(index)
                    
        except Exception as e:
            print(f"Error updating makes based on selections: {e}")
    
    def update_models_based_on_selections(self, year_to_use, make_to_use):
        """Update model dropdown to show only models that match current year and make selections."""
        try:
            valid_models = set()
            for item in self.data['prequal']:
                try:
                    if not self.has_valid_prequal(item):
                        continue
                    
                    # Check if this item matches our current selections
                    # Handle year comparison with proper type conversion
                    try:
                        item_year = int(float(item.get('Year', '')))
                        selected_year = int(float(year_to_use)) if year_to_use not in ["Select Year", ""] else None
                        year_matches = (selected_year is None or item_year == selected_year)
                    except (ValueError, TypeError):
                        year_matches = (year_to_use in ["Select Year", ""] or 
                                      str(item.get('Year', '')).strip() == str(year_to_use).strip())
                    
                    make_matches = (make_to_use in ["Select Make", "All", ""] or 
                                  str(item.get('Make', '')).strip() == str(make_to_use).strip())
                    
                    if year_matches and make_matches and 'Model' in item and item['Model'].strip():
                        # Apply region filtering - only include models for makes in current region
                        make = item.get('Make', '').strip()
                        if self.is_make_in_current_region(make):
                            valid_models.add(item['Model'].strip())
                except (ValueError, TypeError, KeyError):
                    continue
            
            # Get models from manufacturer chart data
            try:
                from database_utils import get_unique_models_from_manufacturer_chart
                if year_to_use not in ["Select Year", ""] and make_to_use not in ["Select Make", "All", ""]:
                    manufacturer_models = get_unique_models_from_manufacturer_chart(year_to_use, make_to_use, self.db_path)
                    for model in manufacturer_models:
                        if model.strip():
                            valid_models.add(model.strip())
            except Exception as e:
                logging.error(f"Error getting manufacturer chart models: {e}")
            
            valid_models = sorted(valid_models)
            
            # Store current selection
            current_model = self.model_dropdown.currentText()
            
            # Rebuild dropdown
            self.model_dropdown.blockSignals(True)
            self.model_dropdown.clear()
            self.model_dropdown.addItem("Select Model")
            for model in valid_models:
                self.model_dropdown.addItem(model)
            self.model_dropdown.blockSignals(False)
            
            # Try to restore selection
            if current_model not in ["Select Model", ""]:
                index = self.model_dropdown.findText(current_model)
                if index >= 0:
                    self.model_dropdown.setCurrentIndex(index)
                    
        except Exception as e:
            print(f"Error updating models based on selections: {e}")

    def toggle_region_mode(self):
        """Toggle between ALL and REGION mode"""
        if self.region_toggle_button.isChecked():
            # Switch to REGION mode
            self.region_toggle_button.setText("ALL/REGION")
            self.region_dropdown.show()
            self.region_dropdown.setMaximumHeight(16777215)  # Reset to default max height
            self.current_region = self.region_dropdown.currentText()
            # Restore card shadow for REGION mode
            region_card = self.findChild(QWidget, "region_card")
            if region_card:
                region_card.setStyleSheet("""
                    QWidget#region_card {
                        background: white;
                        border-radius: 8px;
                        border: 1px solid #e0e0e0;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    }
                """)
        else:
            # Switch to ALL mode
            self.region_toggle_button.setText("ALL/REGION")
            self.region_dropdown.hide()
            self.region_dropdown.setMaximumHeight(0)  # Ensure it takes no space
            self.current_region = 'ALL'
            # Remove card shadow for ALL mode
            region_card = self.findChild(QWidget, "region_card")
            if region_card:
                region_card.setStyleSheet("""
                    QWidget#region_card {
                        background: white;
                        border-radius: 8px;
                        border: 1px solid #e0e0e0;
                        box-shadow: none;
                    }
                """)
        
        # Update dropdowns to reflect region filtering
        self.update_dropdowns_with_locks()
        
    def on_region_changed(self):
        """Handle region dropdown selection change"""
        if self.region_toggle_button.isChecked():
            self.current_region = self.region_dropdown.currentText()
            # Update dropdowns to reflect new region
            self.update_dropdowns_with_locks()
    
    def get_region_filtered_makes(self):
        """Get makes filtered by current region selection"""
        if self.current_region == 'ALL':
            return None  # No filtering
        return self.region_makes.get(self.current_region, [])
    
    def is_make_in_current_region(self, make):
        """Check if a make is in the current region"""
        if self.current_region == 'ALL':
            return True
        return make in self.region_makes.get(self.current_region, [])

class VehicleCompareDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Vehicle Comparison")
        self.setFixedSize(1200, 800)
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface with modern styling matching the main app"""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Compact header
        header_layout = QHBoxLayout()
        
        title_label = QLabel("Vehicle Comparison")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: 700;
                color: #333;
            }
        """)
        
        subtitle_label = QLabel("Compare prequalification data between two vehicles")
        subtitle_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #666;
                font-style: italic;
            }
        """)
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(subtitle_label)
        layout.addLayout(header_layout)
        
        # Compact vehicle selection section
        selection_layout = QHBoxLayout()
        selection_layout.setSpacing(15)
        
        # Vehicle 1 selection (compact)
        vehicle1_group = QGroupBox("Vehicle 1")
        vehicle1_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #1976d2;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
                background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #1976d2;
            }
        """)
        vehicle1_layout = QGridLayout(vehicle1_group)
        vehicle1_layout.setSpacing(5)
        
        self.vehicle1_year = ModernComboBox()
        self.vehicle1_year.addItem("Select Year")
        self.vehicle1_make = ModernComboBox()
        self.vehicle1_make.addItem("Select Make")
        self.vehicle1_model = ModernComboBox()
        self.vehicle1_model.addItem("Select Model")
        
        vehicle1_layout.addWidget(QLabel("Year:"), 0, 0)
        vehicle1_layout.addWidget(self.vehicle1_year, 0, 1)
        vehicle1_layout.addWidget(QLabel("Make:"), 1, 0)
        vehicle1_layout.addWidget(self.vehicle1_make, 1, 1)
        vehicle1_layout.addWidget(QLabel("Model:"), 2, 0)
        vehicle1_layout.addWidget(self.vehicle1_model, 2, 1)
        
        # VS label (compact)
        vs_label = QLabel("VS")
        vs_label.setStyleSheet("""
            QLabel {
                font-size: 20px;
                font-weight: 700;
                color: #dc3545;
                padding: 10px;
                background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
                color: white;
                border-radius: 8px;
                min-width: 50px;
                text-align: center;
            }
        """)
        
        # Vehicle 2 selection (compact)
        vehicle2_group = QGroupBox("Vehicle 2")
        vehicle2_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #7b1fa2;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
                background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #7b1fa2;
            }
        """)
        vehicle2_layout = QGridLayout(vehicle2_group)
        vehicle2_layout.setSpacing(5)
        
        self.vehicle2_year = ModernComboBox()
        self.vehicle2_year.addItem("Select Year")
        self.vehicle2_make = ModernComboBox()
        self.vehicle2_make.addItem("Select Make")
        self.vehicle2_model = ModernComboBox()
        self.vehicle2_model.addItem("Select Model")
        
        vehicle2_layout.addWidget(QLabel("Year:"), 0, 0)
        vehicle2_layout.addWidget(self.vehicle2_year, 0, 1)
        vehicle2_layout.addWidget(QLabel("Make:"), 1, 0)
        vehicle2_layout.addWidget(self.vehicle2_make, 1, 1)
        vehicle2_layout.addWidget(QLabel("Model:"), 2, 0)
        vehicle2_layout.addWidget(self.vehicle2_model, 2, 1)
        
        selection_layout.addWidget(vehicle1_group)
        selection_layout.addWidget(vs_label, alignment=Qt.AlignCenter)
        selection_layout.addWidget(vehicle2_group)
        
        layout.addLayout(selection_layout)
        
        # Compact compare button
        compare_button = ModernButton("Compare Vehicles", style="primary")
        compare_button.setStyleSheet("""
            ModernButton {
                font-size: 14px;
                font-weight: 600;
                padding: 10px 20px;
                margin: 10px 0;
            }
        """)
        compare_button.clicked.connect(self.compare_vehicles)
        layout.addWidget(compare_button, alignment=Qt.AlignCenter)
        
        # Results section (maximized)
        results_group = QGroupBox("Comparison Results")
        results_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        results_layout = QVBoxLayout(results_group)
        results_layout.setSpacing(5)
        
        # Results display with modern styling
        self.results_display = ModernTextBrowser()
        self.results_display.setStyleSheet("""
            ModernTextBrowser {
                background: #ffffff;
                border: 2px solid #e9ecef;
                border-radius: 10px;
                padding: 15px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                line-height: 1.5;
            }
        """)
        self.results_display.setMinimumHeight(500)
        results_layout.addWidget(self.results_display)
        
        layout.addWidget(results_group, 1)  # Give results more space
        
        # Populate dropdowns
        self.populate_dropdowns()
        
        # Connect signals
        self.vehicle1_year.currentTextChanged.connect(lambda: self.on_vehicle1_changed('year'))
        self.vehicle1_make.currentTextChanged.connect(lambda: self.on_vehicle1_changed('make'))
        self.vehicle2_year.currentTextChanged.connect(lambda: self.on_vehicle2_changed('year'))
        self.vehicle2_make.currentTextChanged.connect(lambda: self.on_vehicle2_changed('make'))
        
    def populate_dropdowns(self):
        """Populate dropdowns with available data"""
        if not hasattr(self.parent, 'data') or 'prequal' not in self.parent.data:
            return
            
        # Get unique years, makes, and models
        years = set()
        makes = set()
        models = set()
        
        for item in self.parent.data['prequal']:
            if self.parent.has_valid_prequal(item):
                if 'Year' in item and pd.notna(item['Year']):
                    years.add(int(float(item['Year'])))
                if 'Make' in item and item['Make'].strip():
                    makes.add(item['Make'].strip())
                if 'Model' in item and item['Model'].strip():
                    models.add(item['Model'].strip())
        
        # Populate year dropdowns
        sorted_years = sorted(years, reverse=True)
        for year in sorted_years:
            self.vehicle1_year.addItem(str(year))
            self.vehicle2_year.addItem(str(year))
        
        # Populate make dropdowns
        sorted_makes = sorted(makes)
        for make in sorted_makes:
            self.vehicle1_make.addItem(make)
            self.vehicle2_make.addItem(make)
    
    def on_vehicle1_changed(self, field):
        """Handle vehicle 1 selection changes"""
        if field == 'year':
            self.update_vehicle1_models()
        elif field == 'make':
            self.update_vehicle1_models()
    
    def on_vehicle2_changed(self, field):
        """Handle vehicle 2 selection changes"""
        if field == 'year':
            self.update_vehicle2_models()
        elif field == 'make':
            self.update_vehicle2_models()
    
    def update_vehicle1_models(self):
        """Update vehicle 1 model dropdown"""
        year = self.vehicle1_year.currentText()
        make = self.vehicle1_make.currentText()
        
        self.vehicle1_model.clear()
        self.vehicle1_model.addItem("Select Model")
        
        if year != "Select Year" and make != "Select Make":
            models = set()
            for item in self.parent.data['prequal']:
                if (self.parent.has_valid_prequal(item) and 
                    str(item.get('Year', '')).strip() == year and
                    str(item.get('Make', '')).strip() == make and
                    'Model' in item and item['Model'].strip()):
                    models.add(item['Model'].strip())
            
            for model in sorted(models):
                self.vehicle1_model.addItem(model)
    
    def update_vehicle2_models(self):
        """Update vehicle 2 model dropdown"""
        year = self.vehicle2_year.currentText()
        make = self.vehicle2_make.currentText()
        
        self.vehicle2_model.clear()
        self.vehicle2_model.addItem("Select Model")
        
        if year != "Select Year" and make != "Select Make":
            models = set()
            for item in self.parent.data['prequal']:
                if (self.parent.has_valid_prequal(item) and 
                    str(item.get('Year', '')).strip() == year and
                    str(item.get('Make', '')).strip() == make and
                    'Model' in item and item['Model'].strip()):
                    models.add(item['Model'].strip())
            
            for model in sorted(models):
                self.vehicle2_model.addItem(model)
    
    def compare_vehicles(self):
        """Compare the selected vehicles"""
        vehicle1 = {
            'year': self.vehicle1_year.currentText(),
            'make': self.vehicle1_make.currentText(),
            'model': self.vehicle1_model.currentText()
        }
        
        vehicle2 = {
            'year': self.vehicle2_year.currentText(),
            'make': self.vehicle2_make.currentText(),
            'model': self.vehicle2_model.currentText()
        }
        
        # Validate selections
        if (vehicle1['year'] == "Select Year" or vehicle1['make'] == "Select Make" or vehicle1['model'] == "Select Model" or
            vehicle2['year'] == "Select Year" or vehicle2['make'] == "Select Make" or vehicle2['model'] == "Select Model"):
            self.results_display.setPlainText("Please select complete vehicle information for both vehicles.")
            return
        
        # Get prequalification data for both vehicles
        vehicle1_data = {'prequal': self.get_prequal_data(vehicle1)}
        vehicle2_data = {'prequal': self.get_prequal_data(vehicle2)}
        
        # Check if any data was found
        has_data = vehicle1_data['prequal'] or vehicle2_data['prequal']
        if not has_data:
            self.results_display.setPlainText("No prequalification data found for the selected vehicles.")
            return
        
        # Generate comparison
        comparison = self.generate_detailed_comparison(vehicle1, vehicle2, vehicle1_data, vehicle2_data)
        self.results_display.setHtml(comparison)
    
    def get_prequal_data(self, vehicle):
        """Get prequal data for a specific vehicle"""
        data = []
        for item in self.parent.data['prequal']:
            if (self.parent.has_valid_prequal(item) and
                str(item.get('Year', '')).strip() == vehicle['year'] and
                str(item.get('Make', '')).strip() == vehicle['make'] and
                str(item.get('Model', '')).strip() == vehicle['model']):
                data.append(item)
        return data
    
    def get_blacklist_data(self, vehicle):
        """Get blacklist data for a specific vehicle"""
        data = []
        for item in self.parent.data['blacklist']:
            if (str(item.get('Make', '')).strip() == vehicle['make']):
                data.append(item)
        return data
    
    def get_goldlist_data(self, vehicle):
        """Get goldlist data for a specific vehicle"""
        data = []
        for item in self.parent.data['goldlist']:
            if (str(item.get('Make', '')).strip() == vehicle['make']):
                data.append(item)
        return data
    
    def get_mag_glass_data(self, vehicle):
        """Get mag glass data for a specific vehicle"""
        data = []
        for item in self.parent.data['mag_glass']:
            if (str(item.get('Make', '')).strip() == vehicle['make']):
                data.append(item)
        return data
    
    def get_carsys_data(self, vehicle):
        """Get carsys data for a specific vehicle"""
        data = []
        for item in self.parent.data['carsys']:
            if (str(item.get('Make', '')).strip() == vehicle['make']):
                data.append(item)
        return data
    
    def generate_comparison(self, vehicle1, vehicle2, vehicle1_data, vehicle2_data):
        """Generate HTML comparison between two vehicles"""
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #20567C; color: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }}
                .vehicle-info {{ background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }}
                .comparison-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                .comparison-table th, .comparison-table td {{ 
                    border: 1px solid #dee2e6; padding: 10px; text-align: left; 
                }}
                .comparison-table th {{ background: #e9ecef; font-weight: bold; }}
                .highlight {{ background: #fff3cd; }}
                .different {{ background: #f8d7da; }}
                .same {{ background: #d4edda; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Vehicle Comparison</h1>
                <p>Comparing {vehicle1['year']} {vehicle1['make']} {vehicle1['model']} vs {vehicle2['year']} {vehicle2['make']} {vehicle2['model']}</p>
            </div>
        """
        
        # Vehicle information summary
        html += f"""
            <div class="vehicle-info">
                <h3>Vehicle Information</h3>
                <table class="comparison-table">
                    <tr>
                        <th>Attribute</th>
                        <th>Vehicle 1</th>
                        <th>Vehicle 2</th>
                        <th>Status</th>
                    </tr>
                    <tr>
                        <td>Year</td>
                        <td>{vehicle1['year']}</td>
                        <td>{vehicle2['year']}</td>
                        <td class="{'same' if vehicle1['year'] == vehicle2['year'] else 'different'}">
                            {'Same' if vehicle1['year'] == vehicle2['year'] else 'Different'}
                        </td>
                    </tr>
                    <tr>
                        <td>Make</td>
                        <td>{vehicle1['make']}</td>
                        <td>{vehicle2['make']}</td>
                        <td class="{'same' if vehicle1['make'] == vehicle2['make'] else 'different'}">
                            {'Same' if vehicle1['make'] == vehicle2['make'] else 'Different'}
                        </td>
                    </tr>
                    <tr>
                        <td>Model</td>
                        <td>{vehicle1['model']}</td>
                        <td>{vehicle2['model']}</td>
                        <td class="{'same' if vehicle1['model'] == vehicle2['model'] else 'different'}">
                            {'Same' if vehicle1['model'] == vehicle2['model'] else 'Different'}
                        </td>
                    </tr>
                </table>
            </div>
        """
        
        # Data comparison
        html += f"""
            <div class="vehicle-info">
                <h3>Data Comparison</h3>
                <table class="comparison-table">
                    <tr>
                        <th>Data Type</th>
                        <th>Vehicle 1 Records</th>
                        <th>Vehicle 2 Records</th>
                    </tr>
                    <tr>
                        <td>Prequalification Data</td>
                        <td>{len(vehicle1_data)} records</td>
                        <td>{len(vehicle2_data)} records</td>
                    </tr>
                </table>
            </div>
        """
        
        # Detailed comparison of prequal data
        if vehicle1_data or vehicle2_data:
            html += """
                <div class="vehicle-info">
                    <h3>Prequalification Data Comparison</h3>
                    <table class="comparison-table">
                        <tr>
                            <th>Component</th>
                            <th>Vehicle 1</th>
                            <th>Vehicle 2</th>
                        </tr>
            """
            
            # Get unique components
            components = set()
            for item in vehicle1_data + vehicle2_data:
                if 'Parent Component' in item:
                    components.add(item['Parent Component'])
            
            for component in sorted(components):
                vehicle1_components = [item for item in vehicle1_data if item.get('Parent Component') == component]
                vehicle2_components = [item for item in vehicle2_data if item.get('Parent Component') == component]
                
                html += f"""
                    <tr>
                        <td>{component}</td>
                        <td>{'✓' if vehicle1_components else '✗'} ({len(vehicle1_components)} records)</td>
                        <td>{'✓' if vehicle2_components else '✗'} ({len(vehicle2_components)} records)</td>
                    </tr>
                """
            
            html += "</table></div>"
        
        html += "</body></html>"
        return html
    
    def generate_detailed_comparison(self, vehicle1, vehicle2, vehicle1_data, vehicle2_data):
        """Generate detailed HTML comparison matching main window format"""
        # Extract vehicle information for proper string formatting
        vehicle1_year = vehicle1['year']
        vehicle1_make = vehicle1['make']
        vehicle1_model = vehicle1['model']
        vehicle2_year = vehicle2['year']
        vehicle2_make = vehicle2['make']
        vehicle2_model = vehicle2['model']
        # Build HTML string step by step to avoid f-string issues
        html = """
        <html>
        <head>
            <style>
                body { 
                    font-family: 'Segoe UI', Arial, sans-serif; 
                    margin: 0; 
                    padding: 10px; 
                    background: #ffffff; 
                    color: #333; 
                }
                .vehicle-header { 
                    background: #f0f0f0; 
                    color: #000; 
                    padding: 8px; 
                    border-radius: 4px; 
                    font-weight: 600; 
                    font-size: 14px; 
                    margin-bottom: 10px; 
                    text-align: center; 
                    border: 1px solid #ccc; 
                }
                .content-area { 
                    background: #ffffff; 
                    border: 1px solid #ddd; 
                    border-radius: 4px; 
                    padding: 8px; 
                    min-height: 400px; 
                    overflow-y: auto; 
                }
            </style>
        </head>
        <body>
        """
        
        # Get all unique system names from both vehicles
        v1_prequal = vehicle1_data.get('prequal', [])
        v2_prequal = vehicle2_data.get('prequal', [])
        
        # Create dictionaries to map system names to their data
        v1_systems = {}
        v2_systems = {}
        
        for record in v1_prequal:
            system_name = record.get('Protech Generic System Name', 'Unknown System')
            v1_systems[system_name] = record
            
        for record in v2_prequal:
            system_name = record.get('Protech Generic System Name', 'Unknown System')
            v2_systems[system_name] = record
        
        # Get all unique system names
        all_systems = set(v1_systems.keys()) | set(v2_systems.keys())
        all_systems = sorted(list(all_systems))
        
        if all_systems:
            # Create the main comparison table that spans full width
            html += """
            <div style="margin-bottom: 20px;">
                <div style="background: #f5f5f5; padding: 10px; border-radius: 4px; margin-bottom: 10px; font-weight: bold; text-align: center;">
                    System Comparison: """ + f"{vehicle1_year} {vehicle1_make} {vehicle1_model}" + """ vs """ + f"{vehicle2_year} {vehicle2_make} {vehicle2_model}" + """
                </div>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr style="background: #f0f0f0; font-weight: bold;">
                        <td style="padding: 8px; border: 1px solid #ddd; width: 50%; text-align: center;">Vehicle 1: """ + f"{vehicle1_year} {vehicle1_make} {vehicle1_model}" + """</td>
                        <td style="padding: 8px; border: 1px solid #ddd; width: 50%; text-align: center;">Vehicle 2: """ + f"{vehicle2_year} {vehicle2_make} {vehicle2_model}" + """</td>
                    </tr>
            """
            
            for system_name in all_systems:
                v1_has_system = system_name in v1_systems
                v2_has_system = system_name in v2_systems
                
                html += f'<tr><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">'
                
                if v1_has_system:
                    # Format Vehicle 1 system data
                    record = v1_systems[system_name]
                    html += f'<div style="margin-bottom: 10px;"><strong>{system_name}</strong></div>'
                    html += self.format_single_record(record)
                else:
                    html += f'<div style="text-align: center; padding: 20px; color: #999; font-style: italic;">System not available: {system_name}</div>'
                
                html += '</td><td style="padding: 8px; border: 1px solid #ddd; vertical-align: top;">'
                
                if v2_has_system:
                    # Format Vehicle 2 system data
                    record = v2_systems[system_name]
                    html += f'<div style="margin-bottom: 10px;"><strong>{system_name}</strong></div>'
                    html += self.format_single_record(record)
                else:
                    html += f'<div style="text-align: center; padding: 20px; color: #999; font-style: italic;">System not available: {system_name}</div>'
                
                html += '</td></tr>'
            
            html += '</table></div>'
        else:
            html += '<div style="text-align: center; padding: 40px; color: #666; font-style: italic;">*No prequalification data found for either vehicle.*</div>'
        
        html += """
        </body>
        </html>
        """
        return html
    
    def format_comparison_data(self, results):
        """Format data for comparison using main window style"""
        if not results:
            return '<div style="text-align: center; padding: 40px; color: #666; font-style: italic;">*No prequalification data found for the selected criteria.*</div>'
        
        # Convert to markdown format (same as main window)
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
        
        # Convert markdown to HTML for display (same as main window)
        html_text = self.convert_markdown_to_html(markdown_text)
        return html_text
    
    def format_single_record(self, record):
        """Format a single prequalification record for display"""
        if not record:
            return '<div style="color: #999; font-style: italic;">No data available</div>'
        
        html = '<div style="background: #fafafa; padding: 10px; border-radius: 4px; margin-bottom: 8px;">'
        
        # Display key fields
        key_fields = [
            ('Parent Component', 'Parent Component'),
            ('Calibration Type', 'Calibration Type'),
            ('Calibration Pre-Requisites', 'Calibration Pre-Requisites'),
            ('Calibration Pre-Requisites (Short Hand)', 'Short Hand'),
            ('Point of Impact #', 'Point of Impact'),
            ('OEM ADAS System Name', 'OEM System Name'),
            ('Component Generic Acronyms', 'Component Acronyms')
        ]
        
        for field_key, display_name in key_fields:
            if field_key in record and record[field_key] and str(record[field_key]) != 'nan':
                value = str(record[field_key])
                html += f'<div style="margin-bottom: 6px;"><strong>{display_name}:</strong> {value}</div>'
        
        # Add service information link if available
        if 'Service Information Hyperlink' in record and record['Service Information Hyperlink'] and str(record['Service Information Hyperlink']) != 'nan':
            link = record['Service Information Hyperlink']
            html += f'<div style="margin-bottom: 6px;"><strong>Service Info:</strong> <a href="{link}" style="color: #0066cc;">View Service Information</a></div>'
        
        html += '</div>'
        return html
    
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
    
    def format_side_by_side_data(self, data, data_type, vehicle_type='vehicle1'):
        """Format data for side-by-side display"""
        if not data:
            return '<span class="no-data">No data</span>'
        
        # Set colors based on vehicle type
        if vehicle_type == 'vehicle1':
            border_color = '#1976d2'  # Blue
            header_color = '#1976d2'
        else:
            border_color = '#7b1fa2'  # Purple
            header_color = '#7b1fa2'
        
        html = ""
        if data_type == 'prequal':
            for i, item in enumerate(data):
                html += f"""
                    <div style=\"margin-bottom: 20px; padding: 16px; background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border-radius: 10px; border: 2px solid {border_color}; box-shadow: 0 4px 12px rgba(0,0,0,0.1); position: relative; overflow: hidden;\">
                        <div style=\"position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, {header_color}, {border_color});\"></div>
                        <div style=\"background: {header_color}; color: white; padding: 8px 12px; margin: -16px -16px 12px -16px; border-radius: 8px 8px 0 0; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px;\">
                            <span style=\"font-size: 16px;\">📋</span> Record {i+1}
                        </div>
                        <div style=\"display: grid; gap: 8px;\">
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🔧 Component:</strong> {item.get('Parent Component', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">⚙️ System:</strong> {item.get('Protech Generic System Name', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🎯 Calibration Type:</strong> {item.get('Calibration Type', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">📋 Prerequisites:</strong> {item.get('Calibration Pre-Requisites (Short Hand)', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">📍 Point of Impact:</strong> {item.get('Point of Impact #', 'N/A')}
                            </div>
                        </div>
                    </div>
                """
        elif data_type in ['blacklist', 'goldlist']:
            for i, item in enumerate(data):
                html += f"""
                    <div style=\"margin-bottom: 20px; padding: 16px; background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border-radius: 10px; border: 2px solid {border_color}; box-shadow: 0 4px 12px rgba(0,0,0,0.1); position: relative; overflow: hidden;\">
                        <div style=\"position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, {header_color}, {border_color});\"></div>
                        <div style=\"background: {header_color}; color: white; padding: 8px 12px; margin: -16px -16px 12px -16px; border-radius: 8px 8px 0 0; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px;\">
                            <span style=\"font-size: 16px;\">🚨</span> Record {i+1}
                        </div>
                        <div style=\"display: grid; gap: 8px;\">
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🔢 DTC Code:</strong> {item.get('dtcCode', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">📝 Description:</strong> {item.get('dtcDescription', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🚗 Make:</strong> {item.get('carMake', 'N/A')}
                            </div>
                        </div>
                    </div>
                """
        elif data_type == 'mag_glass':
            for i, item in enumerate(data):
                html += f"""
                    <div style=\"margin-bottom: 20px; padding: 16px; background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border-radius: 10px; border: 2px solid {border_color}; box-shadow: 0 4px 12px rgba(0,0,0,0.1); position: relative; overflow: hidden;\">
                        <div style=\"position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, {header_color}, {border_color});\"></div>
                        <div style=\"background: {header_color}; color: white; padding: 8px 12px; margin: -16px -16px 12px -16px; border-radius: 8px 8px 0 0; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px;\">
                            <span style=\"font-size: 16px;\">🔍</span> Record {i+1}
                        </div>
                        <div style=\"display: grid; gap: 8px;\">
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🚗 Make:</strong> {item.get('Make', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🏷️ Model:</strong> {item.get('Model', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">📅 Year:</strong> {item.get('Year', 'N/A')}
                            </div>
                        </div>
                    </div>
                """
        elif data_type == 'carsys':
            for i, item in enumerate(data):
                html += f"""
                    <div style=\"margin-bottom: 20px; padding: 16px; background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border-radius: 10px; border: 2px solid {border_color}; box-shadow: 0 4px 12px rgba(0,0,0,0.1); position: relative; overflow: hidden;\">
                        <div style=\"position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, {header_color}, {border_color});\"></div>
                        <div style=\"background: {header_color}; color: white; padding: 8px 12px; margin: -16px -16px 12px -16px; border-radius: 8px 8px 0 0; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px;\">
                            <span style=\"font-size: 16px;\">💻</span> Record {i+1}
                        </div>
                        <div style=\"display: grid; gap: 8px;\">
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🚗 Make:</strong> {item.get('Make', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">🏷️ Model:</strong> {item.get('Model', 'N/A')}
                            </div>
                            <div style=\"background: rgba(0,0,0,0.02); padding: 8px; border-radius: 6px; border-left: 3px solid {border_color};\">
                                <strong style=\"color: {header_color};\">⚙️ System:</strong> {item.get('System', 'N/A')}
                            </div>
                        </div>
                    </div>
                """
        return html

    def format_data_preview(self, data, data_type):
        """Format a preview of the data for display"""
        if not data:
            return '<span class="no-data">No data</span>'
        
        html = ""
        
        if data_type == 'prequal':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 10px; padding: 8px; background: #f8f9fa; border-radius: 4px;">
                        <strong>Record {i+1}:</strong><br/>
                        <strong>Component:</strong> {item.get('Parent Component', 'N/A')}<br/>
                        <strong>System:</strong> {item.get('Protech Generic System Name', 'N/A')}<br/>
                        <strong>Calibration Type:</strong> {item.get('Calibration Type', 'N/A')}<br/>
                        <strong>Prerequisites:</strong> {item.get('Calibration Pre-Requisites (Short Hand)', 'N/A')}
                    </div>
                """
        elif data_type in ['blacklist', 'goldlist']:
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 10px; padding: 8px; background: #f8f9fa; border-radius: 4px;">
                        <strong>Record {i+1}:</strong><br/>
                        <strong>DTC Code:</strong> {item.get('dtcCode', 'N/A')}<br/>
                        <strong>Description:</strong> {item.get('dtcDescription', 'N/A')}<br/>
                        <strong>Make:</strong> {item.get('carMake', 'N/A')}
                    </div>
                """
        elif data_type == 'mag_glass':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 10px; padding: 8px; background: #f8f9fa; border-radius: 4px;">
                        <strong>Record {i+1}:</strong><br/>
                        <strong>Make:</strong> {item.get('Make', 'N/A')}<br/>
                        <strong>Model:</strong> {item.get('Model', 'N/A')}<br/>
                        <strong>Year:</strong> {item.get('Year', 'N/A')}
                    </div>
                """
        elif data_type == 'carsys':
            for i, item in enumerate(data):
                html += f"""
                    <div style="margin-bottom: 10px; padding: 8px; background: #f8f9fa; border-radius: 4px;">
                        <strong>Record {i+1}:</strong><br/>
                        <strong>Make:</strong> {item.get('Make', 'N/A')}<br/>
                        <strong>Model:</strong> {item.get('Model', 'N/A')}<br/>
                        <strong>System:</strong> {item.get('System', 'N/A')}
                    </div>
                """
        
        return html

# ... existing code ...

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
