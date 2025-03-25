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
import pandas as pd  # Make sure this import is available
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices, QIcon, QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLineEdit, QComboBox, QTextBrowser, QListWidget, QLabel, 
    QPushButton, QToolBar, QStatusBar, QMessageBox, QDialog, 
    QFileDialog, QSplitter, QFrame, QProgressBar, QInputDialog,
    QShortcut, QSizePolicy, QSlider, QFormLayout
)

# Configure logging
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

class UserLoginDialog(QDialog):
    def __init__(self, parent=None, theme='Dark'):
        super().__init__(parent)
        self.setWindowTitle("User Login")
        self.setGeometry(100, 100, 350, 240)  # Increased height to accommodate new buttons
        self.setModal(True)
        self.theme = theme
        self.result_code = QDialog.Rejected  # Default result code
        self.setup_ui()
        self.apply_theme()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Add logo or icon (optional)
        title_label = QLabel("Analyzer+")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(16)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)
        
        # Add instructions
        instruction_label = QLabel("Please enter your PIN to login:")
        instruction_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(instruction_label)
        
        # PIN input field
        self.pin_input = QLineEdit()
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)  # Limit to 4 digits for PIN
        self.pin_input.setAlignment(Qt.AlignCenter)
        self.pin_input.setPlaceholderText("Enter 4-digit PIN")  # Add placeholder text
        self.pin_input.returnPressed.connect(self.accept)  # Allow enter key to submit
        layout.addWidget(self.pin_input)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        # Cancel button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        # Login button
        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.accept)
        button_layout.addWidget(self.login_button)
        
        layout.addLayout(button_layout)
        
        # Add "Create Account" option
        signup_layout = QHBoxLayout()
        signup_label = QLabel("Don't have a PIN?")
        signup_layout.addWidget(signup_label)
        
        self.signup_button = QPushButton("Create Account")
        self.signup_button.clicked.connect(self.on_signup_clicked)
        signup_layout.addWidget(self.signup_button)
        
        layout.addLayout(signup_layout)
        
        # Add "Reset PIN" option
        reset_layout = QHBoxLayout()
        reset_label = QLabel("Forgot your PIN?")
        reset_layout.addWidget(reset_label)
        
        self.reset_button = QPushButton("Reset PIN")
        self.reset_button.clicked.connect(self.on_reset_clicked)
        reset_layout.addWidget(self.reset_button)
        
        layout.addLayout(reset_layout)
        
        self.setLayout(layout)
    
    def on_signup_clicked(self):
        self.result_code = 2  # Use code 2 for signup
        self.done(2)
    
    def on_reset_clicked(self):
        self.result_code = 3  # Use code 3 for reset PIN
        self.done(3)
    
    def apply_theme(self):
        # Apply the same theme/style as the main application
        if self.theme == "Dark":
            self.apply_dark_theme()
        else:
            self.apply_color_theme(self.get_color_for_theme(self.theme))
    
    def get_color_for_theme(self, theme):
        # Return appropriate colors based on theme name
        theme_colors = {
            "Red": ("#DC143C", "#A52A2A", "#fff"),
            "Blue": ("#4169E1", "#1E3A5F", "#fff"),
            "Light": ("#f0f0f0", "#d0d0d0", "black"),
            "Green": ("#006400", "#003a00", "#fff"),
            "Yellow": ("#ffd700", "#b29500", "black"),
            "Pink": ("#ff69b4", "#b2477d", "black"),
            "Purple": ("#800080", "#4b004b", "#fff"),
            "Teal": ("#008080", "#004b4b", "#fff"),
            "Cyan": ("#00ffff", "#00b2b2", "black"),
            "Orange": ("#ff8c00", "#ff4500", "black")
        }
        
        return theme_colors.get(theme, ("#2b2b2b", "#2b2b2b", "#ddd"))  # Default to dark theme colors
    
    def apply_dark_theme(self):
        style = """
        QDialog {
            background-color: #2b2b2b;
            color: #ddd;
        }
        QLabel {
            color: #ddd;
            font-weight: bold;
        }
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        QLineEdit {
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            background-color: #333;
            color: white;
            font-weight: bold;
            height: 24px;
        }
        QLineEdit:hover {
            border: 2px solid #aaa;
        }
        """
        self.setStyleSheet(style)
    
    def apply_color_theme(self, colors):
        color_start, color_end, text_color = colors
        style = f"""
        QDialog {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
        }}
        QLabel {{
            color: {text_color};
            font-weight: bold;
        }}
        QPushButton {{
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }}
        QPushButton:hover {{
            background-color: #555;
        }}
        QPushButton:pressed {{
            background-color: #666;
            border-style: inset;
        }}
        QLineEdit {{
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            background-color: #333;
            color: white;
            font-weight: bold;
            height: 24px;
        }}
        QLineEdit:hover {{
            border: 2px solid #aaa;
        }}
        """
        self.setStyleSheet(style)
    
    def get_pin(self):
        return self.pin_input.text()
        
    def exec_(self):
        """Override exec_ to return our custom result codes"""
        result = super().exec_()
        
        # If the dialog was accepted through the standard accept() method
        # (like when login button is clicked), pass that through
        if result == QDialog.Accepted:
            return QDialog.Accepted
            
        # Otherwise return our custom result code
        return self.result_code

class PinManagementDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Users and PINs")
        self.setGeometry(100, 100, 400, 300)
        self.db_path = parent.db_path
        self.setup_ui()
        self.load_users_and_pins()
        # Apply theme from parent app
        if parent:
            self.apply_theme(parent.current_theme)

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

        # Add a debug button to view all users - helpful for administrators
        self.debug_button = QPushButton("Show All Users")
        self.debug_button.clicked.connect(self.show_all_users)
        button_layout.addWidget(self.debug_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def apply_theme(self, theme):
        """Apply the application theme to this dialog"""
        if theme == "Dark":
            base_style = self.parent().get_dialog_dark_style()
            list_style = """
            QListWidget {
                background-color: #333;
                border: 2px solid #555;
                border-radius: 10px;
                color: white;
                padding: 5px;
            }
            QListWidget::item {
                height: 25px;
                color: white;
            }
            QListWidget::item:selected {
                background-color: #4f70e0;
                border-radius: 5px;
            }
            QListWidget::item:hover {
                background-color: #555;
                border-radius: 5px;
            }
            """
            self.setStyleSheet(base_style + list_style)
        else:
            # For color themes, get the appropriate colors
            colors = self.parent().get_colors_for_theme(theme)
            if colors:
                color_start, color_end, text_color = colors
                base_style = self.parent().get_dialog_color_style(color_start, color_end, text_color)
                list_style = f"""
                QListWidget {{
                    background-color: #333;
                    border: 2px solid #555;
                    border-radius: 10px;
                    color: white;
                    padding: 5px;
                }}
                QListWidget::item {{
                    height: 25px;
                    color: white;
                }}
                QListWidget::item:selected {{
                    background-color: #4f70e0;
                    border-radius: 5px;
                }}
                QListWidget::item:hover {{
                    background-color: #555;
                    border-radius: 5px;
                }}
                """
                self.setStyleSheet(base_style + list_style)
                
    def show_all_users(self):
        """Debug function to show all users and their PINs in a messagebox"""
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT pin, name FROM leader_log')
        users = cursor.fetchall()
        conn.close()
        
        if not users:
            msg = self.parent().create_styled_messagebox("User List", "No users found in the database.", QMessageBox.Information)
            msg.exec_()
            return
            
        user_list_text = "All registered users:\n\n"
        for pin, name in users:
            user_list_text += f"Username: {name}, PIN: {pin}\n"
            
        msg = self.parent().create_styled_messagebox("User List", user_list_text, QMessageBox.Information)
        msg.exec_()
        self.parent().log_action(self.parent().current_user, "Viewed all user accounts (debug)")

    def load_users_and_pins(self):
        self.user_list.clear()
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT pin, name FROM leader_log')
        users = cursor.fetchall()
        for pin, name in users:
            self.user_list.addItem(f"{pin} - {name}")

    def add_user(self):
        # Create a custom dialog for adding a new user
        add_dialog = QDialog(self)
        add_dialog.setWindowTitle("Add User")
        add_dialog.setMinimumWidth(300)
        
        dialog_layout = QVBoxLayout()
        
        # PIN input
        pin_label = QLabel("Enter new PIN:")
        dialog_layout.addWidget(pin_label)
        
        pin_input = QLineEdit()
        pin_input.setEchoMode(QLineEdit.Password)
        pin_input.setMaxLength(4)  # Limit PIN to 4 digits
        dialog_layout.addWidget(pin_input)
        
        # Button layout
        button_layout = QHBoxLayout()
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(add_dialog.reject)
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(add_dialog.accept)
        button_layout.addWidget(cancel_button)
        button_layout.addWidget(ok_button)
        dialog_layout.addLayout(button_layout)
        
        add_dialog.setLayout(dialog_layout)
        
        # Apply the same theme as the parent dialog
        if self.parent():
            if self.parent().current_theme == "Dark":
                add_dialog.setStyleSheet(self.parent().get_dialog_dark_style())
            else:
                colors = self.parent().get_colors_for_theme(self.parent().current_theme)
                if colors:
                    color_start, color_end, text_color = colors
                    add_dialog.setStyleSheet(self.parent().get_dialog_color_style(color_start, color_end, text_color))
        
        # Show the dialog
        if add_dialog.exec_() == QDialog.Accepted:
            pin = pin_input.text()
            if pin:
                # Now ask for the user name
                name_dialog = QDialog(self)
                name_dialog.setWindowTitle("Add User")
                name_dialog.setMinimumWidth(300)
                
                name_layout = QVBoxLayout()
                
                name_label = QLabel("Enter user name:")
                name_layout.addWidget(name_label)
                
                name_input = QLineEdit()
                name_layout.addWidget(name_input)
                
                name_button_layout = QHBoxLayout()
                name_cancel_button = QPushButton("Cancel")
                name_cancel_button.clicked.connect(name_dialog.reject)
                name_ok_button = QPushButton("OK")
                name_ok_button.clicked.connect(name_dialog.accept)
                name_button_layout.addWidget(name_cancel_button)
                name_button_layout.addWidget(name_ok_button)
                name_layout.addLayout(name_button_layout)
                
                name_dialog.setLayout(name_layout)
                
                # Apply the same theme
                if self.parent():
                    if self.parent().current_theme == "Dark":
                        name_dialog.setStyleSheet(self.parent().get_dialog_dark_style())
                    else:
                        colors = self.parent().get_colors_for_theme(self.parent().current_theme)
                        if colors:
                            color_start, color_end, text_color = colors
                            name_dialog.setStyleSheet(self.parent().get_dialog_color_style(color_start, color_end, text_color))
                
                if name_dialog.exec_() == QDialog.Accepted:
                    name = name_input.text()
                    if name:
                        self.save_user(pin, name)

    def edit_user(self):
        current_item = self.user_list.currentItem()
        if current_item:
            pin, name = current_item.text().split(" - ")
            
            # Create a custom dialog for editing the PIN
            edit_dialog = QDialog(self)
            edit_dialog.setWindowTitle("Edit User")
            edit_dialog.setMinimumWidth(300)
            
            dialog_layout = QVBoxLayout()
            
            # PIN input
            pin_label = QLabel("Edit PIN:")
            dialog_layout.addWidget(pin_label)
            
            pin_input = QLineEdit(pin)
            pin_input.setEchoMode(QLineEdit.Password)
            pin_input.setMaxLength(4)  # Limit PIN to 4 digits
            dialog_layout.addWidget(pin_input)
            
            # Button layout
            button_layout = QHBoxLayout()
            cancel_button = QPushButton("Cancel")
            cancel_button.clicked.connect(edit_dialog.reject)
            ok_button = QPushButton("OK")
            ok_button.clicked.connect(edit_dialog.accept)
            button_layout.addWidget(cancel_button)
            button_layout.addWidget(ok_button)
            dialog_layout.addLayout(button_layout)
            
            edit_dialog.setLayout(dialog_layout)
            
            # Apply the same theme as the parent dialog
            if self.parent():
                if self.parent().current_theme == "Dark":
                    edit_dialog.setStyleSheet(self.parent().get_dialog_dark_style())
                else:
                    colors = self.parent().get_colors_for_theme(self.parent().current_theme)
                    if colors:
                        color_start, color_end, text_color = colors
                        edit_dialog.setStyleSheet(self.parent().get_dialog_color_style(color_start, color_end, text_color))
            
            # Show the dialog
            if edit_dialog.exec_() == QDialog.Accepted:
                new_pin = pin_input.text()
                if new_pin:
                    # Now edit the name
                    name_dialog = QDialog(self)
                    name_dialog.setWindowTitle("Edit User")
                    name_dialog.setMinimumWidth(300)
                    
                    name_layout = QVBoxLayout()
                    
                    name_label = QLabel("Edit user name:")
                    name_layout.addWidget(name_label)
                    
                    name_input = QLineEdit(name)
                    name_layout.addWidget(name_input)
                    
                    name_button_layout = QHBoxLayout()
                    name_cancel_button = QPushButton("Cancel")
                    name_cancel_button.clicked.connect(name_dialog.reject)
                    name_ok_button = QPushButton("OK")
                    name_ok_button.clicked.connect(name_dialog.accept)
                    name_button_layout.addWidget(name_cancel_button)
                    name_button_layout.addWidget(name_ok_button)
                    name_layout.addLayout(name_button_layout)
                    
                    name_dialog.setLayout(name_layout)
                    
                    # Apply the same theme
                    if self.parent():
                        if self.parent().current_theme == "Dark":
                            name_dialog.setStyleSheet(self.parent().get_dialog_dark_style())
                        else:
                            colors = self.parent().get_colors_for_theme(self.parent().current_theme)
                            if colors:
                                color_start, color_end, text_color = colors
                                name_dialog.setStyleSheet(self.parent().get_dialog_color_style(color_start, color_end, text_color))
                    
                    if name_dialog.exec_() == QDialog.Accepted:
                        new_name = name_input.text()
                        if new_name:
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
        # Apply theme from parent app
        if parent:
            self.apply_theme(parent.current_theme)

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
        
    def apply_theme(self, theme):
        """Apply the application theme to this dialog"""
        if theme == "Dark":
            self.setStyleSheet(self.parent().get_dialog_dark_style())
        else:
            # For color themes, get the appropriate colors
            colors = self.parent().get_colors_for_theme(theme)
            if colors:
                color_start, color_end, text_color = colors
                self.setStyleSheet(self.parent().get_dialog_color_style(color_start, color_end, text_color))

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

def handle_error(func, path, exc_info):
    logging.error(f"Error copying {path}: {exc_info}")

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

        # Handle existing backup - remove it if it exists
        if os.path.exists(backup_path):
            logging.info(f"Backup path already exists: {backup_path}. Removing to overwrite.")
            try:
                # Try to remove the directory tree
                shutil.rmtree(backup_path)
            except (PermissionError, OSError) as e:
                logging.error(f"Error removing existing backup: {e}. Operations may be incomplete.")

        # Create the new backup with the simpler, more compatible approach
        logging.info(f"Creating backup from {folder_path} to {backup_path}.")
        try:
            # Simple approach that works on all Python versions
            shutil.copytree(folder_path, backup_path, symlinks=False)
            logging.info(f"Backup successfully created at {backup_path}")
        except shutil.Error as e:
            # Let's just log a warning for the copy errors, but not fail the operation
            logging.warning(f"Some errors during backup: {e}")
        except OSError as e:
            logging.error(f"OS error during copy operation: {e}")
    except sqlite3.Error as e:
        logging.error(f"Failed to save path to database: {e}")
    finally:
        conn.close()

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

def load_mag_glass_data(excel_path, table_name='mag_glass', db_path='data.db'):
    error_messages = []
    try:
        # Try to find a sheet named "Mag Glass" first
        try:
            # Try to load a sheet specifically named "Mag Glass"
            df = pd.read_excel(excel_path, sheet_name="Mag Glass")
            logging.debug(f"Data loaded from 'Mag Glass' sheet: {df.head()}")
        except Exception as sheet_error:
            # If sheet named "Mag Glass" is not found, log a warning and return error
            error_message = f"Failed to find 'Mag Glass' sheet in {excel_path}: {str(sheet_error)}"
            logging.warning(error_message)
            error_messages.append(error_message)
            return error_message

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

# Keep the original function for backwards compatibility but modify it to use the new function
def load_mag_glass_data_from_5th_sheet(excel_path, table_name='mag_glass', db_path='data.db'):
    # Instead of using the 5th sheet, now we try to find a sheet named "Mag Glass"
    return load_mag_glass_data(excel_path, table_name, db_path)

class App(QMainWindow):
    def __init__(self, db_path='data.db'):
        super().__init__()
        self.db_path = db_path
        self.settings_file = 'settings.json'
        self.current_theme = 'Dark'
        self.current_user = None
        self.connection_pool = None
        self.conn = sqlite3.connect(self.db_path)
        self.adas_authenticated = False  # Add this line to track ADAS authentication
        initialize_db(self.db_path)
        self.current_theme = self.get_last_logged_theme()  # Load the last logged theme
        self.data = {'blacklist': [], 'goldlist': [], 'prequal': [], 'mag_glass': [], 'carsys': []}
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

    def create_styled_messagebox(self, title, text, icon_type=QMessageBox.Warning):
        """Create a message box with consistent styling matching the current theme."""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(text)
        msg_box.setIcon(icon_type)
        
        # Simple clean style to match the standard Windows style
        style = """
        QMessageBox {
            background-color: white;
            color: black;
        }
        QPushButton {
            background-color: #f0f0f0;
            border: 1px solid #c0c0c0;
            border-radius: 3px;
            padding: 5px;
            min-width: 80px;
            color: black;
        }
        QPushButton:hover {
            background-color: #e5e5e5;
        }
        QPushButton:pressed {
            background-color: #d0d0d0;
        }
        QLabel {
            color: black;
        }
        """
        
        msg_box.setStyleSheet(style)
        return msg_box

    def prompt_user_pin(self):
        attempts = 0
        max_attempts = 3
        while attempts < max_attempts:
            attempts += 1
            # Create our custom dialog with the current theme
            login_dialog = UserLoginDialog(self, self.current_theme)
            
            # Show the dialog and get the result
            result = login_dialog.exec_()
            
            if result == QDialog.Accepted:
                pin = login_dialog.get_pin()
                if pin:  # If pin is not empty
                    conn = self.get_db_connection()
                    cursor = conn.cursor()
                    
                    # Debug logging to check the actual PIN being entered
                    logging.debug(f"Login attempt with PIN: {pin}")
                    
                    # Check all users in the database for debugging
                    debug_cursor = conn.cursor()
                    debug_cursor.execute('SELECT pin, name FROM leader_log')
                    all_users = debug_cursor.fetchall()
                    logging.debug(f"All users in database: {all_users}")
                    
                    # Emergency override PIN for troubleshooting (1234)
                    if pin == '1234':
                        self.current_user = "Emergency Admin"
                        logging.debug("Logged in with emergency PIN")
                        self.log_action(self.current_user, "Logged in with emergency PIN")
                        conn.close()
                        
                        # Create a new user with the standard PIN if none exists
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
                        
                    # Standard PIN check
                    cursor.execute('SELECT name FROM leader_log WHERE pin = ?', (pin,))
                    result = cursor.fetchone()
                    if result:
                        self.current_user = result[0]
                        logging.debug(f"Successful login for user: {self.current_user}")
                        self.log_action(self.current_user, "Logged in")
                        conn.close()
                        return True
                    elif pin == '9716':  # Check for the standard PIN
                        self.current_user = "Set Up"
                        logging.debug("Logged in with standard PIN")
                        self.log_action(self.current_user, "Logged in with standard PIN")
                        conn.close()
                        return True
                    else:
                        # Show all users in the database to help diagnose the issue
                        debug_msg = "Available users in database:\n"
                        for db_pin, db_name in all_users:
                            debug_msg += f"User: {db_name}, PIN: {db_pin}\n"
                        
                        logging.debug(debug_msg)
                        
                        # Create a more informative error message
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
                    # Empty PIN was provided
                    msg_box = self.create_styled_messagebox("Access Denied", "PIN cannot be empty.")
                    msg_box.exec_()
            elif result == 2:  # Signup was selected
                signup_dialog = SignupDialog(self)
                # Apply the current theme to the signup dialog
                if hasattr(signup_dialog, 'setStyleSheet'):
                    if self.current_theme == "Dark":
                        signup_dialog.setStyleSheet(self.get_dialog_dark_style())
                    else:
                        colors = self.get_colors_for_theme(self.current_theme)
                        signup_dialog.setStyleSheet(self.get_dialog_color_style(*colors))
                
                if signup_dialog.exec_() == QDialog.Accepted:
                    # User was created, continue with login prompt
                    logging.debug("User was created, continuing with login prompt")
                    attempts -= 1  # Don't count this as an attempt
                    continue
            elif result == 3:  # Reset PIN was selected
                reset_dialog = ResetPinDialog(self)
                # Apply the current theme to the reset dialog
                if hasattr(reset_dialog, 'setStyleSheet'):
                    if self.current_theme == "Dark":
                        reset_dialog.setStyleSheet(self.get_dialog_dark_style())
                    else:
                        colors = self.get_colors_for_theme(self.current_theme)
                        reset_dialog.setStyleSheet(self.get_dialog_color_style(*colors))
                
                if reset_dialog.exec_() == QDialog.Accepted:
                    # PIN was reset, continue with login prompt
                    logging.debug("PIN was reset, continuing with login prompt")
                    attempts -= 1  # Don't count this as an attempt
                    continue
            else:
                # User canceled the login, offer them an exit option
                msg_box = QMessageBox()
                msg_box.setWindowTitle("Exit Application?")
                msg_box.setText("Do you want to exit the application?")
                msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                msg_box.setDefaultButton(QMessageBox.No)
                
                if hasattr(self, 'current_theme'):
                    if self.current_theme == "Dark":
                        msg_box.setStyleSheet(self.get_dialog_dark_style())
                    else:
                        colors = self.get_colors_for_theme(self.current_theme)
                        msg_box.setStyleSheet(self.get_dialog_color_style(*colors))
                
                if msg_box.exec_() == QMessageBox.Yes:
                    logging.debug("User chose to exit the application")
                    sys.exit()
                else:
                    logging.debug("User chose to return to login")
                    attempts -= 1  # Don't count this as an attempt
                    continue
        
        # If we've reached the maximum number of attempts
        msg_box = self.create_styled_messagebox(
            "Login Failed", 
            f"Maximum login attempts ({max_attempts}) reached. The application will now exit."
        )
        msg_box.exec_()
        sys.exit()

    def get_colors_for_theme(self, theme):
        """Return appropriate colors based on theme name for dialogs."""
        theme_colors = {
            "Red": ("#DC143C", "#A52A2A", "#fff"),
            "Blue": ("#4169E1", "#1E3A5F", "#fff"),
            "Light": ("#f0f0f0", "#d0d0d0", "black"),
            "Green": ("#006400", "#003a00", "#fff"),
            "Yellow": ("#ffd700", "#b29500", "black"),
            "Pink": ("#ff69b4", "#b2477d", "black"),
            "Purple": ("#800080", "#4b004b", "#fff"),
            "Teal": ("#008080", "#004b4b", "#fff"),
            "Cyan": ("#00ffff", "#00b2b2", "black"),
            "Orange": ("#ff8c00", "#ff4500", "black")
        }
        return theme_colors.get(theme, ("#2b2b2b", "#2b2b2b", "#ddd"))
    
    def get_dialog_dark_style(self):
        """Return dark style for dialogs."""
        return """
        QDialog {
            background-color: #2b2b2b;
            color: #ddd;
        }
        QLabel {
            color: #ddd;
            font-weight: bold;
        }
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        QLineEdit {
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            background-color: #333;
            color: white;
            font-weight: bold;
            height: 24px;
        }
        QLineEdit:hover {
            border: 2px solid #aaa;
        }
        """
    
    def get_dialog_color_style(self, color_start, color_end, text_color):
        """Return color theme style for dialogs."""
        return f"""
        QDialog {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
        }}
        QLabel {{
            color: {text_color};
            font-weight: bold;
        }}
        QPushButton {{
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 28px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }}
        QPushButton:hover {{
            background-color: #555;
        }}
        QPushButton:pressed {{
            background-color: #666;
            border-style: inset;
        }}
        QLineEdit {{
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            background-color: #333;
            color: white;
            font-weight: bold;
            height: 24px;
        }}
        QLineEdit:hover {{
            border: 2px solid #aaa;
        }}
        """

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
        pin_button = self.add_pin_on_top_button()
        self.toolbar.addWidget(pin_button)
        clear_button = self.add_clear_filters_button()
        self.toolbar.addWidget(clear_button)  # Actually add the button to the toolbar

        # Create a vertical layout for the Leader button and theme dropdown
        button_theme_layout = QVBoxLayout()

        # Remove the Leader button and directly add only the theme dropdown
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
        
        # Add a keyboard shortcut for admin access (Ctrl+Shift+F1)
        admin_shortcut = QShortcut(QKeySequence("Ctrl+Shift+F1"), self)
        admin_shortcut.activated.connect(self.access_admin_panel)
        
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

        # Reorder dropdowns: Year first, then Make, then Model
        self.year_dropdown = QComboBox()
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.currentIndexChanged.connect(self.on_year_selected)
        dropdown_layout.addWidget(self.year_dropdown)

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

        main_layout.addLayout(dropdown_layout)  # Add the dropdown layout to the top of the main layout

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Enter DTC code or description")
        self.search_bar.returnPressed.connect(self.perform_search)
        self.search_bar.textChanged.connect(self.update_suggestions)
        main_layout.addWidget(self.search_bar)

        self.filter_dropdown = QComboBox()
        self.filter_dropdown.addItems(["Select List", "Blacklist", "Goldlist", "Prequals", "Mag Glass", "CarSys", "Gold and Black", "Black/Gold/Mag", "All"])
        self.filter_dropdown.currentIndexChanged.connect(self.filter_changed)  # Connect this to a new method
        main_layout.addWidget(self.filter_dropdown)

        self.suggestions_list = QListWidget(self)
        self.suggestions_list.setMaximumHeight(100)
        self.suggestions_list.hide()
        self.suggestions_list.itemClicked.connect(self.on_suggestion_clicked)
        main_layout.addWidget(self.suggestions_list)

        # Add a horizontal line separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)

        # Create a splitter for horizontal arrangement of panels
        self.splitter = QSplitter(Qt.Horizontal)
        
        # Configure splitter to take available space and properly stretch
        self.splitter.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Create container widgets for panels
        self.left_panel_container = QWidget()
        self.left_panel_layout = QVBoxLayout(self.left_panel_container)
        
        # Add label and pop out button for left panel
        left_panel_header_layout = QHBoxLayout()
        left_panel_header_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins to tighten layout
        self.left_panel_label = QLabel("Prequals")
        self.left_panel_label.setStyleSheet("font-weight: bold;")
        left_panel_header_layout.addWidget(self.left_panel_label)
        left_panel_header_layout.addStretch(1)  # Add stretch to push button to right side
        self.left_panel_pop_out_button = QPushButton("🔍 Expand")  # Using magnifying glass icon with text
        # Remove hardcoded green styling to use theme-based styling
        self.left_panel_pop_out_button.setFixedSize(70, 20)  # Smaller height for better alignment
        self.left_panel_pop_out_button.setToolTip("Open in new window")
        self.left_panel_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Prequals", self.left_panel.toHtml()))
        left_panel_header_layout.addWidget(self.left_panel_pop_out_button)
        self.left_panel_layout.addLayout(left_panel_header_layout)
        
        # Create and configure left panel text browser
        self.left_panel = QTextBrowser()
        self.left_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.left_panel_layout.addWidget(self.left_panel)
        self.splitter.addWidget(self.left_panel_container)

        # Configure right panel container
        self.right_panel_container = QWidget()
        self.right_panel_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.right_panel_layout = QVBoxLayout(self.right_panel_container)
        
        # Add label and pop out button for right panel
        right_panel_header_layout = QHBoxLayout()
        right_panel_header_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins to tighten layout
        self.right_panel_label = QLabel("Gold and Black")
        self.right_panel_label.setStyleSheet("font-weight: bold;")
        right_panel_header_layout.addWidget(self.right_panel_label)
        right_panel_header_layout.addStretch(1)  # Add stretch to push button to right side
        self.right_panel_pop_out_button = QPushButton("🔍 Expand")  # Using magnifying glass icon with text
        # Remove hardcoded green styling to use theme-based styling
        self.right_panel_pop_out_button.setFixedSize(70, 20)  # Smaller height for better alignment
        self.right_panel_pop_out_button.setToolTip("Open in new window")
        self.right_panel_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Gold and Black", self.right_panel.toHtml()))
        right_panel_header_layout.addWidget(self.right_panel_pop_out_button)
        self.right_panel_layout.addLayout(right_panel_header_layout)
        
        # Configure right panel text browser
        self.right_panel = QTextBrowser()
        self.right_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.right_panel_layout.addWidget(self.right_panel)
        self.splitter.addWidget(self.right_panel_container)

        # Configure Mag Glass container
        self.mag_glass_container = QWidget()
        self.mag_glass_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.mag_glass_layout = QVBoxLayout(self.mag_glass_container)
        
        # Add label and pop out button for Mag Glass panel
        mag_glass_header_layout = QHBoxLayout()
        mag_glass_header_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins to tighten layout
        self.mag_glass_label = QLabel("Mag Glass")
        self.mag_glass_label.setStyleSheet("font-weight: bold;")
        mag_glass_header_layout.addWidget(self.mag_glass_label)
        mag_glass_header_layout.addStretch(1)  # Add stretch to push button to right side
        self.mag_glass_pop_out_button = QPushButton("🔍 Expand")  
        # Remove hardcoded green styling to use theme-based styling
        self.mag_glass_pop_out_button.setFixedSize(70, 20)  # Smaller height for better alignment
        self.mag_glass_pop_out_button.setToolTip("Open in new window")
        self.mag_glass_pop_out_button.clicked.connect(lambda: self.pop_out_panel("Mag Glass", self.mag_glass_panel.toHtml()))
        mag_glass_header_layout.addWidget(self.mag_glass_pop_out_button)
        self.mag_glass_layout.addLayout(mag_glass_header_layout)
        
        # Configure Mag Glass panel text browser
        self.mag_glass_panel = QTextBrowser()
        self.mag_glass_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.mag_glass_layout.addWidget(self.mag_glass_panel)
        self.splitter.addWidget(self.mag_glass_container)
        self.mag_glass_container.setVisible(False)

        # Configure CarSys container
        self.carsys_container = QWidget()
        self.carsys_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.carsys_layout = QVBoxLayout(self.carsys_container)
        
        # Add label and pop out button for CarSys panel
        carsys_header_layout = QHBoxLayout()
        carsys_header_layout.setContentsMargins(0, 0, 0, 0)
        self.carsys_label = QLabel("CarSys")
        self.carsys_label.setStyleSheet("font-weight: bold;")
        carsys_header_layout.addWidget(self.carsys_label)
        carsys_header_layout.addStretch(1)
        self.carsys_pop_out_button = QPushButton("🔍 Expand")
        # Remove hardcoded green styling to use theme-based styling
        self.carsys_pop_out_button.setFixedSize(70, 20)
        self.carsys_pop_out_button.setToolTip("Open in new window")
        self.carsys_pop_out_button.clicked.connect(lambda: self.pop_out_panel("CarSys", self.carsys_panel.toHtml()))
        carsys_header_layout.addWidget(self.carsys_pop_out_button)
        self.carsys_layout.addLayout(carsys_header_layout)
        
        # Configure CarSys panel text browser
        self.carsys_panel = QTextBrowser()
        self.carsys_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.carsys_layout.addWidget(self.carsys_panel)
        self.splitter.addWidget(self.carsys_container)  # Add CarSys to splitter so it's part of the UI
        self.carsys_container.setVisible(False)

        # Configure splitter to take up maximum space in the layout
        main_layout.addWidget(self.splitter, 1)  # Add stretch factor of 1 to allow expansion

        # Reset margins to default - the reduced margin was causing overlap
        main_layout.setContentsMargins(main_layout.contentsMargins().left(), 
                                      main_layout.contentsMargins().top(), 
                                      main_layout.contentsMargins().right(), 
                                      main_layout.contentsMargins().bottom())
        
        self.add_hide_show_buttons(main_layout)  # Add hide/show buttons

        # Add a flexible spacer to push content up and keep transparency bar at bottom
        main_layout.addStretch(0)  # Reduced stretch factor to give more space to the splitter

        # Create a fixed panel for transparency controls that stays at the bottom
        transparency_panel = QWidget()
        # Increase the height to create more space
        transparency_panel.setFixedHeight(70)  # Increased from 60 to give more space
        transparency_panel.setContentsMargins(0, 10, 0, 0)  # Add top margin to create space from buttons
        
        # Create layout for the transparency panel
        transparency_layout = QVBoxLayout(transparency_panel)
        transparency_layout.setAlignment(Qt.AlignCenter)
        transparency_layout.setContentsMargins(0, 10, 0, 5)  # Increased top padding
        
        # Center the label
        transparency_label_layout = QHBoxLayout()
        transparency_label_layout.setAlignment(Qt.AlignCenter)
        self.opacity_label = QLabel("Transparency")
        self.opacity_label.setStyleSheet("font-weight: bold;")
        transparency_label_layout.addWidget(self.opacity_label)
        transparency_layout.addLayout(transparency_label_layout)
        
        # Center and make the slider smaller
        slider_layout = QHBoxLayout()
        slider_layout.setAlignment(Qt.AlignCenter)
        self.opacity_slider = QSlider(Qt.Horizontal)
        self.opacity_slider.setMinimum(20)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.setFixedWidth(150)  # Make the slider smaller
        self.opacity_slider.valueChanged.connect(self.change_opacity)
        slider_layout.addWidget(self.opacity_slider)
        transparency_layout.addLayout(slider_layout)
        
        # Add the fixed transparency panel to the main layout
        main_layout.addWidget(transparency_panel)

        self.apply_saved_theme()  # Apply the last saved theme
        self.populate_dropdowns()
        self.apply_dropdown_styles()  # Apply dropdown styles initially

        # Adding tooltips to important UI elements
        self.search_bar.setToolTip("Enter DTC code or description to search")
        self.year_dropdown.setToolTip("Select model year")
        self.make_dropdown.setToolTip("Select vehicle make")
        self.model_dropdown.setToolTip("Select vehicle model")
        self.filter_dropdown.setToolTip("Select list type")
        
        # Add tooltips to hide/show buttons
        self.left_hide_show_button.setToolTip("Hide/Show Prequals panel")
        self.right_hide_show_button.setToolTip("Hide/Show Gold and Black panel")
        self.mag_glass_hide_show_button.setToolTip("Hide/Show Mag Glass panel")
        self.carsys_hide_show_button.setToolTip("Hide/Show CarSys panel")

        # Apply theme-colored styling to expand buttons
        self.apply_theme_to_expand_buttons()

    def pop_out_panel(self, title, content):
        pop_out_window = PopOutWindow(title, content, self.central_widget.styleSheet(), self)
        pop_out_window.show()

    def get_db_connection(self):
        return sqlite3.connect(self.db_path)

    def on_year_selected(self):
        selected_year = self.year_dropdown.currentText()
        selected_filter = self.filter_dropdown.currentText()
        
        # If a year is selected, update the make dropdown to show only makes from that year
        if selected_year != "Select Year":
            try:
                selected_year_int = int(selected_year)
                # Filter prequal data with proper NaN handling
                valid_prequal_data = []
                for item in self.data['prequal']:
                    try:
                        if (self.has_valid_prequal(item) and 
                            'Year' in item and pd.notna(item['Year'])):
                            item_year = int(float(item['Year']))
                            if item_year == selected_year_int:
                                valid_prequal_data.append(item)
                    except (ValueError, TypeError, KeyError):
                        continue  # Skip items with invalid year values
                
                # Get unique makes for this year
                makes = []
                for item in valid_prequal_data:
                    try:
                        if 'Make' in item and item['Make'].strip() and item['Make'].strip().lower() != 'unknown':
                            makes.append(item['Make'].strip())
                    except (AttributeError, KeyError):
                        continue
                
                makes = sorted(set(makes))
                
                # Update the make dropdown
                self.make_dropdown.clear()
                self.make_dropdown.addItem("Select Make")
                self.make_dropdown.addItem("All")
                self.make_dropdown.addItems(makes)
                
                # Reset model dropdown
                self.model_dropdown.clear()
                self.model_dropdown.addItem("Select Model")
                
                logging.debug(f"Updated make dropdown with {len(makes)} makes for year {selected_year}")
                
                # If we already have a make selected (e.g. after changing year), update the models
                current_make = self.make_dropdown.currentText()
                if current_make not in ["Select Make", "All"]:
                    self.update_model_dropdown()
            except (ValueError, TypeError) as e:
                logging.error(f"Error updating makes for year {selected_year}: {e}")
        
        # Check if a specific list that requires updating on year change is selected
        if selected_filter in ["Prequals", "Gold and Black", "Blacklist", "Goldlist", "Mag Glass", "All"]:
            self.perform_search()

    def update_model_dropdown(self):
        selected_year = self.year_dropdown.currentText().strip()
        selected_make = self.make_dropdown.currentText().strip()
        
        # Use the new simpler method to populate models
        self.populate_models(selected_year, selected_make)
        
        # Log that we're updating models
        logging.debug(f"Updated model dropdown for Year: {selected_year}, Make: {selected_make}")

    def handle_model_change(self, index):
        selected_model = self.model_dropdown.currentText().strip()
        logging.debug(f"Model selected: {selected_model}")
        
        # The year is already selected at this point, so we just need to perform the search
        self.perform_search()

    def filter_changed(self):
        selected_filter = self.filter_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()

        # Store current visibility states
        left_visible = self.left_panel_container.isVisible()
        right_visible = self.right_panel_container.isVisible()
        mag_glass_visible = self.mag_glass_container.isVisible()
        carsys_visible = self.carsys_container.isVisible()

        # Handle "All" or "Prequals" with no selected year
        if selected_make == "All":
            self.model_dropdown.setCurrentIndex(0)
            selected_model = "Select Model"

        # Only require year for Prequals filter
        if selected_filter == "Prequals" and selected_year == "Select Year":
            self.left_panel.clear()
            self.left_panel.setPlainText("Please select a Year for Prequals search.")
            # Don't return early, still respect panel visibility

        # Update content in panels based on the selected filter
        if selected_filter == "Mag Glass" or mag_glass_visible:
            self.display_mag_glass(selected_make)
        
        if selected_filter == "CarSys" or carsys_visible:
            dtc_code = self.search_bar.text().strip().upper()
            self.search_carsys_dtc(dtc_code)
        
        # When needed, make specific panels visible based on filter selection
        if selected_filter == "Mag Glass" and not mag_glass_visible:
            self.mag_glass_container.setVisible(True)
            self.update_button_text()
        
        if selected_filter == "CarSys" and not carsys_visible:
            self.carsys_container.setVisible(True)
            self.update_button_text()
        
        if selected_filter in ["All", "Black/Gold/Mag"]:
            # Make panels visible if they aren't already
            if not mag_glass_visible:
                self.mag_glass_container.setVisible(True)
            if not right_visible:
                self.right_panel_container.setVisible(True)
            # Show left panel only for All with a year selected
            if selected_filter == "All" and selected_year != "Select Year" and not left_visible:
                self.left_panel_container.setVisible(True)
            self.update_button_text()
        
        if selected_filter == "Prequals" and not left_visible:
            self.left_panel_container.setVisible(True)
            self.update_button_text()
        
        if selected_filter in ["Blacklist", "Goldlist", "Gold and Black"] and not right_visible:
            self.right_panel_container.setVisible(True)
            self.update_button_text()

        # Perform the search to update the view
        self.perform_search()
    
    def update_button_text(self):
        """Update the hide/show button text based on panel visibility"""
        # Update Prequals button
        if self.left_panel_container.isVisible():
            self.left_hide_show_button.setText("Hide Prequals")
        else:
            self.left_hide_show_button.setText("Show Prequals")
            
        # Update Gold and Black button
        if self.right_panel_container.isVisible():
            self.right_hide_show_button.setText("Hide Gold and Black")
        else:
            self.right_hide_show_button.setText("Show Gold and Black")
            
        # Update Mag Glass button
        if self.mag_glass_container.isVisible():
            self.mag_glass_hide_show_button.setText("Hide Mag Glass")
        else:
            self.mag_glass_hide_show_button.setText("Show Mag Glass")
            
        # Update CarSys button
        if self.carsys_container.isVisible():
            self.carsys_hide_show_button.setText("Hide CarSys")
        else:
            self.carsys_hide_show_button.setText("Show CarSys")

    def add_hide_show_buttons(self, layout):
        # Remove the negative spacing that was moving buttons up too much
        # layout.addSpacing(-10)  # This was causing the buttons to overlap with the transparency slider
        
        button_layout = QHBoxLayout()
        # Set initial button text to "Show" since panels are hidden by default
        self.left_hide_show_button = QPushButton("Show Prequals")
        self.left_hide_show_button.setFixedSize(120, 30)
        self.left_hide_show_button.clicked.connect(self.toggle_left_panel)
        button_layout.addWidget(self.left_hide_show_button)

        self.right_hide_show_button = QPushButton("Show Gold and Black")
        self.right_hide_show_button.setFixedSize(160, 30)
        self.right_hide_show_button.clicked.connect(self.toggle_right_panel)
        button_layout.addWidget(self.right_hide_show_button)

        self.mag_glass_hide_show_button = QPushButton("Show Mag Glass")
        self.mag_glass_hide_show_button.setFixedSize(120, 30)
        self.mag_glass_hide_show_button.clicked.connect(self.toggle_mag_glass_panel)
        button_layout.addWidget(self.mag_glass_hide_show_button)

        # Add CarSys hide/show button
        self.carsys_hide_show_button = QPushButton("Show CarSys")
        self.carsys_hide_show_button.setFixedSize(120, 30)
        self.carsys_hide_show_button.clicked.connect(self.toggle_carsys_panel)
        button_layout.addWidget(self.carsys_hide_show_button)

        layout.addLayout(button_layout)
        
        # Add some spacing after the buttons to ensure they don't overlap with transparency controls
        layout.addSpacing(10)
        
        # Ensure all panels are hidden at startup
        self.left_panel_container.setVisible(False)
        self.right_panel_container.setVisible(False)
        self.mag_glass_container.setVisible(False)
        self.carsys_container.setVisible(False)

    def toggle_left_panel(self):
        if self.left_panel_container.isVisible():
            self.left_panel_container.hide()
        else:
            self.left_panel_container.show()
        self.update_button_text()

    def toggle_right_panel(self):
        if self.right_panel_container.isVisible():
            self.right_panel_container.hide()
        else:
            self.right_panel_container.show()
        self.update_button_text()

    def toggle_mag_glass_panel(self):
        if self.mag_glass_container.isVisible():
            self.mag_glass_container.hide()
        else:
            self.mag_glass_container.show()
        self.update_button_text()

    def toggle_carsys_panel(self):
        if self.carsys_container.isVisible():
            self.carsys_container.hide()
        else:
            self.carsys_container.show()
        self.update_button_text()

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

    def access_admin_panel(self):
        # Skip the PIN check and directly access admin options
        self.log_action(self.current_user, "Accessed Admin options via keyboard shortcut (Ctrl+Shift+F1)")
        self.open_admin_options()

    def add_refresh_button(self):
        refresh_button = QPushButton("Refresh Lists")
        refresh_button.setObjectName("refresh_button")
        refresh_button.setToolTip("Reload all data from configured Excel files")
        refresh_button.clicked.connect(self.refresh_lists_thread)
        return refresh_button

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
        clear_filters_button = QPushButton("Clear Filters")
        clear_filters_button.setObjectName("clear_filters_button")
        clear_filters_button.setToolTip("Reset all dropdown filters and search input")
        clear_filters_button.clicked.connect(self.clear_filters)
        return clear_filters_button

    def update_black_and_gold_display(self):
        selected_make = self.make_dropdown.currentText().strip()
        selected_filter = self.filter_dropdown.currentText().strip()
        
        if selected_make and selected_make != "Select Make":
            # Pass the current filter or use Gold and Black as default
            if selected_filter in ["Goldlist", "Blacklist"]:
                self.display_gold_and_black(selected_make, selected_filter)
            else:
                self.display_gold_and_black(selected_make, "Gold and Black")
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
        self.pin_on_top_button = QPushButton("Pin on Top")
        self.pin_on_top_button.setObjectName("pin_on_top_button")
        self.pin_on_top_button.setToolTip("Keep this window always on top of other windows")
        self.pin_on_top_button.setCheckable(True)
        self.pin_on_top_button.setChecked(False)
        self.pin_on_top_button.toggled.connect(self.toggle_always_on_top)
        return self.pin_on_top_button

    def toggle_always_on_top(self, checked):
        flags = self.windowFlags()
        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
            
            # Get current theme
            theme = self.theme_dropdown.currentText()
            
            # Define theme colors mapping
            theme_colors = {
                "Red": ("#DC143C", "#A52A2A", "white"),
                "Blue": ("#4169E1", "#1E3A5F", "white"),
                "Dark": ("#2b2b2b", "#333333", "#ddd"),
                "Light": ("#f0f0f0", "#d0d0d0", "black"),
                "Green": ("#006400", "#003a00", "white"),
                "Yellow": ("#ffd700", "#b29500", "black"),
                "Pink": ("#ff69b4", "#b2477d", "black"),
                "Purple": ("#800080", "#4b004b", "white"),
                "Teal": ("#008080", "#004b4b", "white"),
                "Cyan": ("#00ffff", "#00b2b2", "black"),
                "Orange": ("#ff8c00", "#ff4500", "black")
            }
            
            # If in dark theme, use light theme colors for the activated button state
            if theme == "Dark":
                # Use Light theme colors for the activated button in dark theme
                color_start, color_end, text_color = theme_colors["Light"]
            else:
                # For other themes, use their own colors
                if theme in theme_colors:
                    color_start, color_end, text_color = theme_colors[theme]
                else:
                    # Default to Dark theme if theme not found
                    color_start, color_end, text_color = theme_colors["Dark"]
                
            # Create theme-colored button style
            button_style = f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
                border: 2px solid #555;
                border-radius: 10px;
                padding: 5px;
                min-height: 30px;
                margin: 2px;
                color: {text_color};
                font-weight: bold;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_end}, stop:1 {color_start});
                border: 2px solid #777;
            }}
            QPushButton:pressed {{
                background-color: #666;
            }}
            """
            
            self.pin_on_top_button.setStyleSheet(button_style)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
            
            # Reset to default button style
            self.pin_on_top_button.setStyleSheet("")
            # Apply the standard button style from apply_button_styles
            self.apply_button_styles()
            
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
        self.update_theme_dropdown_style()  # Update the theme dropdown style
        self.apply_theme_to_expand_buttons()  # Apply theme to expand buttons
        
        # Update "Pin on Top" button if it's toggled on
        if hasattr(self, 'pin_on_top_button') and self.pin_on_top_button.isChecked():
            # Simulate toggling it to update its styling
            self.toggle_always_on_top(True)

    def apply_dropdown_styles(self):
        # Use a consistent style based on the Clear Filters button style
        dropdown_style = """
        QComboBox {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            min-height: 30px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QComboBox:hover {
            background-color: #555;
            border: 2px solid #777;
        }
        QComboBox:on {
            background-color: #666;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
            border-left: 1px solid #555;
            border-top-right-radius: 8px;
            border-bottom-right-radius: 8px;
        }
        QComboBox::down-arrow {
            width: 14px;
            height: 14px;
        }
        QComboBox QAbstractItemView {
            border: 2px solid #555;
            background-color: #333;
            selection-background-color: #555;
            color: white;
        }
        """

        # Apply to all combo boxes in the application
        for combo in self.findChildren(QComboBox):
            combo.setStyleSheet(dropdown_style)
            
        # Also apply to specific dropdowns we know about
        if hasattr(self, 'theme_dropdown'):
            self.theme_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'year_dropdown'):
            self.year_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'make_dropdown'):
            self.make_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'model_dropdown'):
            self.model_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'filter_dropdown'):
            self.filter_dropdown.setStyleSheet(dropdown_style)

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
        # Get the currently selected theme
        theme = self.theme_dropdown.currentText()
        
        # Define theme colors mapping
        theme_colors = {
            "Red": ("#DC143C", "#A52A2A", "white"),
            "Blue": ("#4169E1", "#1E3A5F", "white"),
            "Dark": ("#2b2b2b", "#333333", "#ddd"),
            "Light": ("#f0f0f0", "#d0d0d0", "black"),
            "Green": ("#006400", "#003a00", "white"),
            "Yellow": ("#ffd700", "#b29500", "black"),
            "Pink": ("#ff69b4", "#b2477d", "black"),
            "Purple": ("#800080", "#4b004b", "white"),
            "Teal": ("#008080", "#004b4b", "white"),
            "Cyan": ("#00ffff", "#00b2b2", "black"),
            "Orange": ("#ff8c00", "#ff4500", "black")
        }
        
        # Get colors for the selected theme
        if theme in theme_colors:
            color_start, color_end, text_color = theme_colors[theme]
        else:
            # Default to Dark theme if theme not found
            color_start, color_end, text_color = theme_colors["Dark"]
            
        # Create a style that uses the theme colors but keeps the structure from our consistent style
        dropdown_style = f"""
        QComboBox {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            min-height: 30px;
            margin: 2px;
            color: {text_color};
            font-weight: bold;
        }}
        QComboBox:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_end}, stop:1 {color_start});
            border: 2px solid #777;
        }}
        QComboBox:on {{
            background-color: #666;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
            border-left: 1px solid #555;
            border-top-right-radius: 8px;
            border-bottom-right-radius: 8px;
        }}
        QComboBox::down-arrow {{
            width: 14px;
            height: 14px;
        }}
        QComboBox QAbstractItemView {{
            border: 2px solid #555;
            background-color: #333;
            selection-background-color: {color_start};
            color: white;
        }}
        """
        
        self.theme_dropdown.setStyleSheet(dropdown_style)

    def apply_color_theme(self, color_start, color_end, text_color="#e0e0e0"):
        # Convert hex colors to RGB for gradient calculation
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        start_rgb = hex_to_rgb(color_start)
        end_rgb = hex_to_rgb(color_end)
        
        # Calculate a darker shade for backgrounds
        dark_start = tuple(max(0, c - 40) for c in start_rgb)
        dark_end = tuple(max(0, c - 40) for c in end_rgb)
        
        # Convert RGB back to hex
        dark_start_hex = '#{:02x}{:02x}{:02x}'.format(*dark_start)
        dark_end_hex = '#{:02x}{:02x}{:02x}'.format(*dark_end)
        
        style = f"""
        QWidget {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
            font-family: 'Segoe UI', Arial, sans-serif;
        }}
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {{
            font: 11px 'Segoe UI';
            color: {text_color};
        }}
        QPushButton {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 8px 16px;
            min-height: 32px;
            margin: 2px;
            transition: all 0.2s;
        }}
        QPushButton:hover {{
            background-color: {color_start};
            border-color: {color_end};
        }}
        QPushButton:pressed {{
            background-color: {color_end};
            border-style: inset;
        }}
        QPushButton#search_button, QPushButton#refresh_button, 
        QPushButton#admin_button, QPushButton#export_button, 
        QPushButton#pin_on_top_button, QPushButton#leader_button {{
            background-color: {dark_start_hex};
            border-color: {color_start};
            color: {text_color};
        }}
        QPushButton#search_button:hover, QPushButton#refresh_button:hover,
        QPushButton#admin_button:hover, QPushButton#export_button:hover,
        QPushButton#pin_on_top_button:hover, QPushButton#leader_button:hover {{
            background-color: {color_start};
            border-color: {color_end};
        }}
        QTextBrowser {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 8px;
            color: {text_color};
            font-family: 'Consolas', monospace;
            font-size: 11px;
        }}
        QComboBox {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 6px 12px;
            min-width: 120px;
            min-height: 32px;
        }}
        QComboBox:hover {{
            border-color: {color_end};
        }}
        QComboBox::drop-down {{
            border: none;
            width: 20px;
        }}
        QComboBox::down-arrow {{
            image: url(down_arrow.png);
            width: 12px;
            height: 12px;
        }}
        QComboBox QAbstractItemView {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            selection-background-color: {color_start};
            selection-color: {text_color};
        }}
        QLineEdit {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 6px 12px;
            min-height: 32px;
        }}
        QLineEdit:hover {{
            border-color: {color_end};
        }}
        QLineEdit:focus {{
            border-color: {color_end};
        }}
        QListWidget {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 4px;
        }}
        QListWidget::item {{
            padding: 6px;
            border-radius: 2px;
        }}
        QListWidget::item:selected {{
            background-color: {color_start};
            color: {text_color};
        }}
        QStatusBar {{
            background-color: {dark_start_hex};
            border-top: 1px solid {color_start};
            padding: 4px;
        }}
        QToolBar {{
            background-color: {dark_start_hex};
            border-bottom: 1px solid {color_start};
            spacing: 4px;
            padding: 4px;
        }}
        QProgressBar {{
            border: 1px solid {color_start};
            border-radius: 4px;
            text-align: center;
            background-color: {dark_start_hex};
        }}
        QProgressBar::chunk {{
            background-color: {color_start};
            border-radius: 3px;
        }}
        QSlider::groove:horizontal {{
            border: 1px solid {color_start};
            height: 8px;
            background: {dark_start_hex};
            border-radius: 4px;
        }}
        QSlider::handle:horizontal {{
            background: {color_start};
            border: 1px solid {color_end};
            width: 18px;
            margin: -2px 0;
            border-radius: 9px;
        }}
        QSlider::handle:horizontal:hover {{
            background: {color_end};
            border-color: {color_start};
        }}
        QLabel {{
            color: {text_color};
            font-weight: bold;
        }}
        QFrame {{
            border: 1px solid {color_start};
        }}
        QSplitter::handle {{
            background-color: {color_start};
        }}
        """
        self.apply_stylesheet(style)

    def apply_dark_theme(self):
        style = """
        QWidget {
            background-color: #1a1a1a;
            color: #e0e0e0;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {
            font: 11px 'Segoe UI';
            color: #e0e0e0;
        }
        QPushButton {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 8px 16px;
            min-height: 32px;
            margin: 2px;
            transition: all 0.2s;
        }
        QPushButton:hover {
            background-color: #3d3d3d;
            border-color: #4d4d4d;
        }
        QPushButton:pressed {
            background-color: #4d4d4d;
            border-style: inset;
        }
        QPushButton#search_button, QPushButton#refresh_button, 
        QPushButton#admin_button, QPushButton#export_button, 
        QPushButton#pin_on_top_button, QPushButton#leader_button {
            background-color: #2d5a88;
            border-color: #3d6a98;
            color: white;
        }
        QPushButton#search_button:hover, QPushButton#refresh_button:hover,
        QPushButton#admin_button:hover, QPushButton#export_button:hover,
        QPushButton#pin_on_top_button:hover, QPushButton#leader_button:hover {
            background-color: #3d6a98;
            border-color: #4d7aa8;
        }
        QTextBrowser {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 8px;
            color: #e0e0e0;
            font-family: 'Consolas', monospace;
            font-size: 11px;
        }
        QComboBox {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 6px 12px;
            min-width: 120px;
            min-height: 32px;
        }
        QComboBox:hover {
            border-color: #4d4d4d;
        }
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        QComboBox::down-arrow {
            image: url(down_arrow.png);
            width: 12px;
            height: 12px;
        }
        QComboBox QAbstractItemView {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            selection-background-color: #2d5a88;
            selection-color: white;
        }
        QLineEdit {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 6px 12px;
            min-height: 32px;
        }
        QLineEdit:hover {
            border-color: #4d4d4d;
        }
        QLineEdit:focus {
            border-color: #2d5a88;
        }
        QListWidget {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 4px;
        }
        QListWidget::item {
            padding: 6px;
            border-radius: 2px;
        }
        QListWidget::item:selected {
            background-color: #2d5a88;
            color: white;
        }
        QStatusBar {
            background-color: #2d2d2d;
            border-top: 1px solid #3d3d3d;
            padding: 4px;
        }
        QToolBar {
            background-color: #2d2d2d;
            border-bottom: 1px solid #3d3d3d;
            spacing: 4px;
            padding: 4px;
        }
        QProgressBar {
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            text-align: center;
            background-color: #2d2d2d;
        }
        QProgressBar::chunk {
            background-color: #2d5a88;
            border-radius: 3px;
        }
        QSlider::groove:horizontal {
            border: 1px solid #3d3d3d;
            height: 8px;
            background: #2d2d2d;
            border-radius: 4px;
        }
        QSlider::handle:horizontal {
            background: #2d5a88;
            border: 1px solid #3d6a98;
            width: 18px;
            margin: -2px 0;
            border-radius: 9px;
        }
        QSlider::handle:horizontal:hover {
            background: #3d6a98;
            border-color: #4d7aa8;
        }
        QLabel {
            color: #e0e0e0;
            font-weight: bold;
        }
        QFrame {
            border: 1px solid #3d3d3d;
        }
        QSplitter::handle {
            background-color: #3d3d3d;
        }
        """
        self.apply_stylesheet(style)

    def apply_button_styles(self):
        # Define a consistent style for all buttons based on the Clear Filters button style
        button_style = """
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 30px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        """
        
        # Apply this style to all buttons in the application
        for button in self.findChildren(QPushButton):
            button.setStyleSheet(button_style)
        
        # Also apply to specific buttons we know about
        if hasattr(self, 'clear_filters_button'):
            self.clear_filters_button.setStyleSheet(button_style)
        if hasattr(self, 'leader_button'):
            self.leader_button.setStyleSheet(button_style)
        if hasattr(self, 'pin_on_top_button'):
            self.pin_on_top_button.setStyleSheet(button_style)
        if hasattr(self, 'refresh_button'):
            self.refresh_button.setStyleSheet(button_style)
        
        # Ensure we apply the style to button widgets added to layouts
        if hasattr(self, 'toolbar'):
            for i in range(self.toolbar.layout().count()):
                widget = self.toolbar.layout().itemAt(i).widget()
                if isinstance(widget, QPushButton):
                    widget.setStyleSheet(button_style)

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
        selected_year = self.year_dropdown.currentText().strip()
        selected_make = self.make_dropdown.currentText().strip()
        
        # Use the new simpler method to populate models
        self.populate_models(selected_year, selected_make)
        
        # Log that we're updating models
        logging.debug(f"Updated model dropdown for Year: {selected_year}, Make: {selected_make}")

    def handle_model_change(self, index):
        selected_model = self.model_dropdown.currentText().strip()
        logging.debug(f"Model selected: {selected_model}")
        
        # The year is already selected at this point, so we just need to perform the search
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
        self.update_theme_dropdown_style()  # Update the theme dropdown style
        self.apply_theme_to_expand_buttons()  # Apply theme to expand buttons
        
        # Update "Pin on Top" button if it's toggled on
        if hasattr(self, 'pin_on_top_button') and self.pin_on_top_button.isChecked():
            # Simulate toggling it to update its styling
            self.toggle_always_on_top(True)

    def apply_dropdown_styles(self):
        # Use a consistent style based on the Clear Filters button style
        dropdown_style = """
        QComboBox {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            min-height: 30px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QComboBox:hover {
            background-color: #555;
            border: 2px solid #777;
        }
        QComboBox:on {
            background-color: #666;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
            border-left: 1px solid #555;
            border-top-right-radius: 8px;
            border-bottom-right-radius: 8px;
        }
        QComboBox::down-arrow {
            width: 14px;
            height: 14px;
        }
        QComboBox QAbstractItemView {
            border: 2px solid #555;
            background-color: #333;
            selection-background-color: #555;
            color: white;
        }
        """

        # Apply to all combo boxes in the application
        for combo in self.findChildren(QComboBox):
            combo.setStyleSheet(dropdown_style)
            
        # Also apply to specific dropdowns we know about
        if hasattr(self, 'theme_dropdown'):
            self.theme_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'year_dropdown'):
            self.year_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'make_dropdown'):
            self.make_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'model_dropdown'):
            self.model_dropdown.setStyleSheet(dropdown_style)
        if hasattr(self, 'filter_dropdown'):
            self.filter_dropdown.setStyleSheet(dropdown_style)

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
        # Get the currently selected theme
        theme = self.theme_dropdown.currentText()
        
        # Define theme colors mapping
        theme_colors = {
            "Red": ("#DC143C", "#A52A2A", "white"),
            "Blue": ("#4169E1", "#1E3A5F", "white"),
            "Dark": ("#2b2b2b", "#333333", "#ddd"),
            "Light": ("#f0f0f0", "#d0d0d0", "black"),
            "Green": ("#006400", "#003a00", "white"),
            "Yellow": ("#ffd700", "#b29500", "black"),
            "Pink": ("#ff69b4", "#b2477d", "black"),
            "Purple": ("#800080", "#4b004b", "white"),
            "Teal": ("#008080", "#004b4b", "white"),
            "Cyan": ("#00ffff", "#00b2b2", "black"),
            "Orange": ("#ff8c00", "#ff4500", "black")
        }
        
        # Get colors for the selected theme
        if theme in theme_colors:
            color_start, color_end, text_color = theme_colors[theme]
        else:
            # Default to Dark theme if theme not found
            color_start, color_end, text_color = theme_colors["Dark"]
            
        # Create a style that uses the theme colors but keeps the structure from our consistent style
        dropdown_style = f"""
        QComboBox {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            min-height: 30px;
            margin: 2px;
            color: {text_color};
            font-weight: bold;
        }}
        QComboBox:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_end}, stop:1 {color_start});
            border: 2px solid #777;
        }}
        QComboBox:on {{
            background-color: #666;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
            border-left: 1px solid #555;
            border-top-right-radius: 8px;
            border-bottom-right-radius: 8px;
        }}
        QComboBox::down-arrow {{
            width: 14px;
            height: 14px;
        }}
        QComboBox QAbstractItemView {{
            border: 2px solid #555;
            background-color: #333;
            selection-background-color: {color_start};
            color: white;
        }}
        """
        
        self.theme_dropdown.setStyleSheet(dropdown_style)

    def apply_color_theme(self, color_start, color_end, text_color="#e0e0e0"):
        # Convert hex colors to RGB for gradient calculation
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        start_rgb = hex_to_rgb(color_start)
        end_rgb = hex_to_rgb(color_end)
        
        # Calculate a darker shade for backgrounds
        dark_start = tuple(max(0, c - 40) for c in start_rgb)
        dark_end = tuple(max(0, c - 40) for c in end_rgb)
        
        # Convert RGB back to hex
        dark_start_hex = '#{:02x}{:02x}{:02x}'.format(*dark_start)
        dark_end_hex = '#{:02x}{:02x}{:02x}'.format(*dark_end)
        
        style = f"""
        QWidget {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {color_start}, stop:1 {color_end});
            color: {text_color};
            font-family: 'Segoe UI', Arial, sans-serif;
        }}
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {{
            font: 11px 'Segoe UI';
            color: {text_color};
        }}
        QPushButton {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 8px 16px;
            min-height: 32px;
            margin: 2px;
            transition: all 0.2s;
        }}
        QPushButton:hover {{
            background-color: {color_start};
            border-color: {color_end};
        }}
        QPushButton:pressed {{
            background-color: {color_end};
            border-style: inset;
        }}
        QPushButton#search_button, QPushButton#refresh_button, 
        QPushButton#admin_button, QPushButton#export_button, 
        QPushButton#pin_on_top_button, QPushButton#leader_button {{
            background-color: {dark_start_hex};
            border-color: {color_start};
            color: {text_color};
        }}
        QPushButton#search_button:hover, QPushButton#refresh_button:hover,
        QPushButton#admin_button:hover, QPushButton#export_button:hover,
        QPushButton#pin_on_top_button:hover, QPushButton#leader_button:hover {{
            background-color: {color_start};
            border-color: {color_end};
        }}
        QTextBrowser {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 8px;
            color: {text_color};
            font-family: 'Consolas', monospace;
            font-size: 11px;
        }}
        QComboBox {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 6px 12px;
            min-width: 120px;
            min-height: 32px;
        }}
        QComboBox:hover {{
            border-color: {color_end};
        }}
        QComboBox::drop-down {{
            border: none;
            width: 20px;
        }}
        QComboBox::down-arrow {{
            image: url(down_arrow.png);
            width: 12px;
            height: 12px;
        }}
        QComboBox QAbstractItemView {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            selection-background-color: {color_start};
            selection-color: {text_color};
        }}
        QLineEdit {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 6px 12px;
            min-height: 32px;
        }}
        QLineEdit:hover {{
            border-color: {color_end};
        }}
        QLineEdit:focus {{
            border-color: {color_end};
        }}
        QListWidget {{
            background-color: {dark_start_hex};
            border: 1px solid {color_start};
            border-radius: 4px;
            padding: 4px;
        }}
        QListWidget::item {{
            padding: 6px;
            border-radius: 2px;
        }}
        QListWidget::item:selected {{
            background-color: {color_start};
            color: {text_color};
        }}
        QStatusBar {{
            background-color: {dark_start_hex};
            border-top: 1px solid {color_start};
            padding: 4px;
        }}
        QToolBar {{
            background-color: {dark_start_hex};
            border-bottom: 1px solid {color_start};
            spacing: 4px;
            padding: 4px;
        }}
        QProgressBar {{
            border: 1px solid {color_start};
            border-radius: 4px;
            text-align: center;
            background-color: {dark_start_hex};
        }}
        QProgressBar::chunk {{
            background-color: {color_start};
            border-radius: 3px;
        }}
        QSlider::groove:horizontal {{
            border: 1px solid {color_start};
            height: 8px;
            background: {dark_start_hex};
            border-radius: 4px;
        }}
        QSlider::handle:horizontal {{
            background: {color_start};
            border: 1px solid {color_end};
            width: 18px;
            margin: -2px 0;
            border-radius: 9px;
        }}
        QSlider::handle:horizontal:hover {{
            background: {color_end};
            border-color: {color_start};
        }}
        QLabel {{
            color: {text_color};
            font-weight: bold;
        }}
        QFrame {{
            border: 1px solid {color_start};
        }}
        QSplitter::handle {{
            background-color: {color_start};
        }}
        """
        self.apply_stylesheet(style)

    def apply_dark_theme(self):
        style = """
        QWidget {
            background-color: #1a1a1a;
            color: #e0e0e0;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        QPushButton, QComboBox, QLineEdit, QCheckBox, QSlider, QToolBar, QTextBrowser, QStatusBar, QListWidget {
            font: 11px 'Segoe UI';
            color: #e0e0e0;
        }
        QPushButton {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 8px 16px;
            min-height: 32px;
            margin: 2px;
            transition: all 0.2s;
        }
        QPushButton:hover {
            background-color: #3d3d3d;
            border-color: #4d4d4d;
        }
        QPushButton:pressed {
            background-color: #4d4d4d;
            border-style: inset;
        }
        QPushButton#search_button, QPushButton#refresh_button, 
        QPushButton#admin_button, QPushButton#export_button, 
        QPushButton#pin_on_top_button, QPushButton#leader_button {
            background-color: #2d5a88;
            border-color: #3d6a98;
            color: white;
        }
        QPushButton#search_button:hover, QPushButton#refresh_button:hover,
        QPushButton#admin_button:hover, QPushButton#export_button:hover,
        QPushButton#pin_on_top_button:hover, QPushButton#leader_button:hover {
            background-color: #3d6a98;
            border-color: #4d7aa8;
        }
        QTextBrowser {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 8px;
            color: #e0e0e0;
            font-family: 'Consolas', monospace;
            font-size: 11px;
        }
        QComboBox {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 6px 12px;
            min-width: 120px;
            min-height: 32px;
        }
        QComboBox:hover {
            border-color: #4d4d4d;
        }
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        QComboBox::down-arrow {
            image: url(down_arrow.png);
            width: 12px;
            height: 12px;
        }
        QComboBox QAbstractItemView {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            selection-background-color: #2d5a88;
            selection-color: white;
        }
        QLineEdit {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 6px 12px;
            min-height: 32px;
        }
        QLineEdit:hover {
            border-color: #4d4d4d;
        }
        QLineEdit:focus {
            border-color: #2d5a88;
        }
        QListWidget {
            background-color: #2d2d2d;
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            padding: 4px;
        }
        QListWidget::item {
            padding: 6px;
            border-radius: 2px;
        }
        QListWidget::item:selected {
            background-color: #2d5a88;
            color: white;
        }
        QStatusBar {
            background-color: #2d2d2d;
            border-top: 1px solid #3d3d3d;
            padding: 4px;
        }
        QToolBar {
            background-color: #2d2d2d;
            border-bottom: 1px solid #3d3d3d;
            spacing: 4px;
            padding: 4px;
        }
        QProgressBar {
            border: 1px solid #3d3d3d;
            border-radius: 4px;
            text-align: center;
            background-color: #2d2d2d;
        }
        QProgressBar::chunk {
            background-color: #2d5a88;
            border-radius: 3px;
        }
        QSlider::groove:horizontal {
            border: 1px solid #3d3d3d;
            height: 8px;
            background: #2d2d2d;
            border-radius: 4px;
        }
        QSlider::handle:horizontal {
            background: #2d5a88;
            border: 1px solid #3d6a98;
            width: 18px;
            margin: -2px 0;
            border-radius: 9px;
        }
        QSlider::handle:horizontal:hover {
            background: #3d6a98;
            border-color: #4d7aa8;
        }
        QLabel {
            color: #e0e0e0;
            font-weight: bold;
        }
        QFrame {
            border: 1px solid #3d3d3d;
        }
        QSplitter::handle {
            background-color: #3d3d3d;
        }
        """
        self.apply_stylesheet(style)

    def apply_button_styles(self):
        # Define a consistent style for all buttons based on the Clear Filters button style
        button_style = """
        QPushButton {
            background-color: #333;
            border: 2px solid #555;
            border-radius: 10px;
            padding: 5px;
            height: 30px;
            margin: 2px;
            color: white;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #555;
        }
        QPushButton:pressed {
            background-color: #666;
            border-style: inset;
        }
        """
        
        # Apply this style to all buttons in the application
        for button in self.findChildren(QPushButton):
            button.setStyleSheet(button_style)
        
        # Also apply to specific buttons we know about
        if hasattr(self, 'clear_filters_button'):
            self.clear_filters_button.setStyleSheet(button_style)
        if hasattr(self, 'leader_button'):
            self.leader_button.setStyleSheet(button_style)
        if hasattr(self, 'pin_on_top_button'):
            self.pin_on_top_button.setStyleSheet(button_style)
        if hasattr(self, 'refresh_button'):
            self.refresh_button.setStyleSheet(button_style)
        
        # Ensure we apply the style to button widgets added to layouts
        if hasattr(self, 'toolbar'):
            for i in range(self.toolbar.layout().count()):
                widget = self.toolbar.layout().itemAt(i).widget()
                if isinstance(widget, QPushButton):
                    widget.setStyleSheet(button_style)

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

    def open_admin(self):
        self.log_action(self.current_user, "Clicked Manage Lists button")
        # Skip authentication and directly show the paths management dialog
        self.manage_paths()

    def show_admin_options(self):
        choice, ok = QInputDialog.getItem(self, "Admin Actions", "Select action:", ["Update Paths", "Clear Data"], 0, False)
        if ok and choice == "Clear Data":
            config_type, ok = QInputDialog.getItem(self, "Clear Data", "Select configuration to clear or 'All' to reset database:", 
                ["Blacklist", "Goldlist", "Prequals (Longsheets)", "MagGlass (Goldlist)", "All"], 0, False)  # Removed CarSys option
            if ok:
                # Map the display names back to internal config types
                config_type_map = {
                    "Blacklist": "blacklist",
                    "Goldlist": "goldlist",
                    "Prequals (Longsheets)": "prequal",
                    "MagGlass (Goldlist)": "mag_glass",
                    # CarSys mapping is kept in the code but hidden from UI
                    # "CarSYS (Goldlist)": "CarSys",
                    "All": "All"
                }
                internal_config_type = config_type_map.get(config_type, config_type)
                
                if internal_config_type == "All":
                    self.clear_data()
                else:
                    self.clear_data(internal_config_type)
        elif ok:
            self.manage_paths()

    def update_year_dropdown(self):
        # This method is no longer needed since we select year first
        # but keeping it to avoid breaking existing code that might call it
        pass

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
        # Create a dialog to manage paths
        self.path_dialog = QDialog(self)
        self.path_dialog.setWindowTitle("Manage Lists")
        self.path_dialog.setGeometry(100, 100, 500, 350)  # Reduced height since we have fewer items
        
        # Apply theme based on current application theme
        if hasattr(self, 'current_theme'):
            if self.current_theme == "Dark":
                self.path_dialog.setStyleSheet(self.get_dialog_dark_style())
            else:
                colors = self.get_colors_for_theme(self.current_theme)
                self.path_dialog.setStyleSheet(self.get_dialog_color_style(*colors))
        
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
        
        # Create a mapping from display names back to internal names
        config_type_map = {
            "Blacklist": "blacklist",
            "Goldlist": "goldlist",
            "Prequals (Longsheets)": "prequal"
        }
        
        # Create a function to handle load button clicks
        def on_load_button_clicked(config_type):
            folder_path = QFileDialog.getExistingDirectory(self, f"Select {config_display_names[config_type]} Directory")
            if folder_path:
                self.browse_buttons[config_type].setText(folder_path)
        
        # Create a function to handle clear button clicks
        def on_clear_button_clicked(config_type):
            # Create confirmation dialog with styled message box
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
                    # Show success message
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
            load_button.setObjectName("action_button")  # Add class for styling
            # Connect the load button to the handler with the correct config_type
            load_button.clicked.connect(lambda checked, config=config_name: on_load_button_clicked(config))
            self.load_buttons[config_name] = load_button
            buttons_layout.addWidget(load_button)
            
            clear_button = QPushButton("Clear Data")
            clear_button.setObjectName("danger_button")  # Add class for styling
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
        main_layout.addSpacing(10)  # Add some space after the title
        
        # Add a container widget for the form
        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        main_layout.addWidget(form_widget)
        
        # Add some space before buttons
        main_layout.addSpacing(10)
        main_layout.addLayout(button_layout)
        
        self.path_dialog.setLayout(main_layout)
        
        # Apply additional button styling based on theme
        if hasattr(self, 'current_theme'):
            if self.current_theme == "Dark":
                self.clear_all_button.setStyleSheet(
                    "background-color: #d9534f; color: white; font-weight: bold; "
                    "border: 2px solid #c9302c; border-radius: 10px; padding: 5px;"
                )
                
                # Determine text color based on background brightness
                # Use the text color from theme colors mapping to ensure proper contrast
                theme_text_color = self.get_colors_for_theme(self.current_theme)[2]  # Get the text color for this theme
                
                self.save_button.setStyleSheet(
                    f"background-color: #5cb85c; color: {theme_text_color}; font-weight: bold; "
                    "border: 2px solid #4cae4c; border-radius: 10px; padding: 5px;"
                )
            else:
                primary_color = self.get_colors_for_theme(self.current_theme)[0]
                self.clear_all_button.setStyleSheet(
                    "background-color: #d9534f; color: white; font-weight: bold; "
                    "border: 2px solid #c9302c; border-radius: 10px; padding: 5px;"
                )
                
                # Determine text color based on background brightness
                # Use the text color from theme colors mapping to ensure proper contrast
                theme_text_color = self.get_colors_for_theme(self.current_theme)[2]  # Get the text color for this theme
                
                self.save_button.setStyleSheet(
                    f"background-color: {primary_color}; color: {theme_text_color}; font-weight: bold; "
                    "border: 2px solid #4cae4c; border-radius: 10px; padding: 5px;"
                )
        
        
        # Pre-populate path fields with existing paths
        for config_name in config_display_names.keys():
            existing_path = load_path_from_db(config_name, self.db_path)
            if existing_path:
                self.browse_buttons[config_name].setText(existing_path)
        
        self.path_dialog.exec_()
        
    def save_paths(self):
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
        
        self.progress_bar.setVisible(False)
        self.load_configurations()  # Reload all configurations
        self.populate_dropdowns()  # Repopulate dropdowns
        self.check_data_loaded()  # Check if data is loaded
        
        # Show a single success message for the entire operation
        msg = self.create_styled_messagebox("Success", "Paths saved and data loaded successfully!", QMessageBox.Information)
        msg.exec_()
        self.path_dialog.accept()
        
    def refresh_lists(self):
        self.log_action(self.current_user, "Clicked Refresh Lists button")
        
        # Track whether any data was loaded successfully
        any_data_loaded = False
        last_processed_path = ""
        
        # Include 'mag_glass' and 'CarSys' in the list for backend processing
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'CarSys']:
            folder_path = load_path_from_db(config_type, self.db_path)
            
            # If no path set for mag_glass or CarSys, use goldlist path
            if not folder_path and config_type in ['mag_glass', 'CarSys']:
                folder_path = load_path_from_db('goldlist', self.db_path)
                if folder_path:
                    # Save the goldlist path for these types
                    save_path_to_db(config_type, folder_path, self.db_path)
            
            if folder_path:
                last_processed_path = folder_path
                # First, clear existing data for the config_type
                self.clear_data(config_type)
                logging.info(f"Cleared existing data for {config_type}")

                # Then get the valid files from the folder path
                files = self.get_valid_excel_files(folder_path)
                if not files:
                    if config_type != 'CarSys':  # Don't show warnings for CarSys to hide it from users
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
                            result = load_mag_glass_data(filepath, config_type, db_path=self.db_path)
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
                        if config_type != 'CarSys':  # Don't show errors for CarSys to hide it from users
                            QMessageBox.critical(self, "Load Error", f"Failed to load {filename}: {str(e)}")

                    self.progress_bar.setValue(i + 1)  # Update progress bar

                self.progress_bar.setVisible(False)

                if data_loaded:
                    any_data_loaded = True
                    self.load_configurations()
                    self.populate_dropdowns()  # Repopulate dropdowns
                    self.check_data_loaded()  # Ensure the dropdowns are enabled if data is loaded
                    
                    # Update status bar but don't show individual success messages
                    self.status_bar.showMessage(f"Data refreshed from: {folder_path}")
            else:
                logging.warning(f"No saved path found for {config_type}")
        
        # Show a single success message after all lists are refreshed
        if any_data_loaded:
            msg = self.create_styled_messagebox("Success", "All data refreshed successfully!", QMessageBox.Information)
            msg.exec_()

    def load_configurations(self):
        logging.debug("Loading configurations...")
        for config_type in ['blacklist', 'goldlist', 'prequal', 'mag_glass', 'carsys']:
            data = load_configuration(config_type, self.db_path)
            self.data[config_type] = data if data else []
            logging.debug(f"Loaded {len(data)} items for {config_type}")
        if 'prequal' in self.data:
            self.populate_dropdowns()

    def populate_dropdowns(self):
        logging.debug("Populating dropdowns...")

        # First populate the Year dropdown
        # Filter prequal data to only include entries with valid prequalification information
        valid_prequal_data = [item for item in self.data['prequal'] if self.has_valid_prequal(item)]

        # Ensure the year values are whole numbers (integers) and remove any non-integer values
        years = []
        for item in valid_prequal_data:
            try:
                if 'Year' in item and pd.notna(item['Year']):  # Check for NaN values
                    year = int(float(item['Year']))  # Convert the year to an integer
                    years.append(year)
            except (ValueError, TypeError):
                continue

        # Remove duplicates and sort the years in descending order (newest first)
        unique_years = sorted(set(years), reverse=True)
        
        self.year_dropdown.clear()
        self.year_dropdown.addItem("Select Year")
        self.year_dropdown.addItems([str(year) for year in unique_years])
        
        # Reset make and model dropdowns
        self.make_dropdown.clear()
        self.make_dropdown.addItem("Select Make")
        self.make_dropdown.addItem("All")
        
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")

        # Gather make values with proper error handling
        makes = set()
        for item in valid_prequal_data:
            try:
                if 'Make' in item and item['Make'] and isinstance(item['Make'], str):
                    make = item['Make'].strip()
                    if make and make.lower() != 'unknown':
                        makes.add(make)
            except (AttributeError, KeyError):
                continue

        # Remove empty string if present
        if '' in makes:
            makes.remove('')

        self.make_dropdown.addItems(sorted(makes))
        logging.debug(f"Makes populated: {sorted(makes)}")

    def has_valid_prequal(self, item):
        # Add your logic here to determine if the item has valid prequalification information
        return item.get('Calibration Pre-Requisites') not in [None, 'N/A']

    def debug_data(self):
        # Print the first few items from each data source for debugging
        print("\nDebugging Data:")
        print("\nBlacklist Data:")
        for item in self.data['blacklist'][:3]:
            print(item)
        print("\nGoldlist Data:")
        for item in self.data['goldlist'][:3]:
            print(item)
        print("\nPrequal Data:")
        for item in self.data['prequal'][:3]:
            print(item)
        print("\nMag Glass Data:")
        for item in self.data['mag_glass'][:3]:
            print(item)

    def perform_search_thread(self):
        threading.Thread(target=self.perform_search).start()

    def perform_search(self):
        if not self.current_user:
            return

        dtc_code = self.search_bar.text().strip().upper()
        selected_filter = self.filter_dropdown.currentText()
        selected_year = self.year_dropdown.currentText()
        selected_make = self.make_dropdown.currentText()
        selected_model = self.model_dropdown.currentText()

        # Get current panel visibility states
        left_visible = self.left_panel_container.isVisible()
        right_visible = self.right_panel_container.isVisible()
        mag_glass_visible = self.mag_glass_container.isVisible()
        carsys_visible = self.carsys_container.isVisible()

        # Reset model if make is set to "All"
        if selected_make == "All":
            self.model_dropdown.setCurrentIndex(0)
            selected_model = "Select Model"

        # Clear panels only if filter is "Select List"
        if selected_filter == "Select List":
            # Don't hide panels, just clear their content
            self.clear_display_panels()
            return

        # Only require year selection for Prequals
        if selected_filter == "Prequals" and selected_year == "Select Year" and left_visible:
            self.left_panel.setPlainText("Please select a Year for Prequals search.")
            # Still process other panels

        # Handle CarSys filter directly when selected
        if selected_filter == "CarSys" or carsys_visible:
            self.search_carsys_dtc(dtc_code)
            if selected_filter == "CarSys":
                self.display_carsys_data(selected_make)

        # Log the search
        self.log_action(self.current_user, f"Performed search with DTC: {dtc_code}, Filter: {selected_filter}, Year: {selected_year}, Make: {selected_make}, Model: {selected_model}")
        
        # Handle prequal search based on year, make, and model
        if (selected_filter == "Prequals" or (selected_filter == "All" and selected_year != "Select Year")) and left_visible:
            if selected_year != "Select Year" and selected_make != "Select Make":
                try:
                    selected_year_int = int(selected_year)  # Convert selected year to integer
                    
                    # Filter prequal data based on the selected year, make, and model with proper error handling for NaN values
                    filtered_prequals = []
                    
                    if selected_model != "Select Model":
                        for item in self.data['prequal']:
                            try:
                                if ('Year' in item and 'Make' in item and 'Model' in item and 
                                    pd.notna(item['Year']) and  # Check for NaN values
                                    item['Make'] == selected_make and 
                                    str(item['Model']) == str(selected_model)):
                                    # Safe conversion
                                    year_value = int(float(item['Year']))
                                    if year_value == selected_year_int:
                                        filtered_prequals.append(item)
                            except (ValueError, TypeError, KeyError):
                                # Skip items with invalid year values
                                continue
                    else:
                        # If model is not selected, show all models for the selected year and make
                        for item in self.data['prequal']:
                            try:
                                if ('Year' in item and 'Make' in item and 
                                    pd.notna(item['Year']) and  # Check for NaN values
                                    item['Make'] == selected_make):
                                    # Safe conversion
                                    year_value = int(float(item['Year']))
                                    if year_value == selected_year_int:
                                        filtered_prequals.append(item)
                            except (ValueError, TypeError, KeyError):
                                # Skip items with invalid year values
                                continue

                    # Display the filtered prequal data
                    if filtered_prequals:
                        self.display_results(filtered_prequals, context='prequal')
                    else:
                        self.left_panel.setPlainText("No prequal data found for the selected criteria.")
                except ValueError:
                    self.left_panel.setPlainText("Invalid year format. Please select a valid year.")

        # Update panels based on visibility
        if right_visible:
            # Show Gold/Black list content
            if dtc_code:
                self.search_dtc_codes(dtc_code, selected_filter if selected_filter in ["Blacklist", "Goldlist"] else "Gold and Black", selected_make)
            else:
                self.display_gold_and_black(selected_make, selected_filter)
                
        if mag_glass_visible:
            # Show Mag Glass content
            self.display_mag_glass(selected_make)
            
        # Make sure button text is up to date
        self.update_button_text()

    def clear_display_panels(self):
        """Clear the content in all panels without changing their visibility"""
        self.left_panel.clear()
        self.right_panel.clear()
        self.mag_glass_panel.clear()
        self.carsys_panel.clear()

    def update_panel_visibility(self, selected_filter, selected_make):
        # For Prequals, the panel should be visible only when not viewing "All" makes
        prequal_visible = selected_filter in ["Prequals", "All"] and selected_make != "All"
        
        # Gold/Black lists are always visible when their filter is selected
        gb_visible = selected_filter in ["Gold and Black", "Blacklist", "Goldlist", "Black/Gold/Mag", "All"]
        
        # Mag Glass is visible when its filter is selected
        mag_visible = selected_filter in ["Mag Glass", "Black/Gold/Mag", "All"]
        
        # Update panel visibility
        self.splitter.widget(0).setVisible(prequal_visible)  # Prequals panel
        self.splitter.widget(1).setVisible(gb_visible)       # Gold/Black panel
        self.splitter.widget(2).setVisible(mag_visible)      # Mag Glass panel

    def update_displays_based_on_filter(self, selected_filter, dtc_code, selected_make, selected_model, selected_year):
        if selected_filter == "All":
            # Display all categories including prequals, gold and black lists, and mag glass
            if selected_year != "Select Year":
                self.handle_prequal_search(selected_make, selected_model, selected_year)
            if dtc_code:
                self.search_dtc_codes(dtc_code, "Gold and Black", selected_make)  # This will depend on DTC input
            else:
                self.display_gold_and_black(selected_make, "Gold and Black")  # Display gold and black lists based on selected make
            self.display_mag_glass(selected_make)  # Display mag glass data

        elif selected_filter == "Prequals":
            self.handle_prequal_search(selected_make, selected_model, selected_year)

        elif selected_filter == "Black/Gold/Mag":
            self.display_gold_and_black(selected_make, "Gold and Black")  # Display gold and black lists based on selected make
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

    def display_gold_and_black(self, selected_make, filter_type="Gold and Black"):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Determine which lists to include based on filter_type
            if filter_type == "Gold and Black" or filter_type == "All" or filter_type == "Black/Gold/Mag":
                # Show both lists
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
            elif filter_type == "Blacklist":
                # Show only blacklist
                if selected_make == "All":
                    query = """
                    SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist
                    """
                else:
                    query = f"""
                    SELECT 'blacklist' as Source, dtcCode, genericSystemName, dtcDescription, dtcSys, carMake, comments FROM blacklist WHERE carMake = '{selected_make}'
                    """
            elif filter_type == "Goldlist":
                # Show only goldlist
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
                self.right_panel.setHtml(df.to_html(index=False, escape=False))
            else:
                self.right_panel.setPlainText("No DTC code results found.")
        except Exception as e:
            logging.error(f"Failed to execute query: {query}\nError: {e}")
            self.right_panel.setPlainText("An error occurred while fetching the data.")
        finally:
            if conn:
                conn.close()

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

    def populate_models(self, year_text, make_text):
        """
        Populate the model dropdown based on selected year and make.
        This is a simpler method to ensure model dropdown is correctly updated.
        """
        self.model_dropdown.clear()
        self.model_dropdown.addItem("Select Model")
        
        # If either year or make is not selected, return early
        if year_text == "Select Year" or make_text == "Select Make" or make_text == "All":
            return
            
        try:
            # Convert year to integer
            year_int = int(year_text)
            
            # Find matching prequal records with proper NaN handling
            matching_models = set()
            for item in self.data['prequal']:
                try:
                    if (self.has_valid_prequal(item) and 
                        'Year' in item and 'Make' in item and 'Model' in item and
                        pd.notna(item['Year']) and  # Explicitly check for NaN values
                        item['Make'] == make_text):
                        # Safe conversion
                        item_year = int(float(item['Year']))
                        if item_year == year_int:
                            matching_models.add(str(item['Model']))
                except (ValueError, TypeError, KeyError):
                    continue  # Skip problematic items
                    
            # Add models to dropdown
            if matching_models:
                self.model_dropdown.addItems(sorted(matching_models))
                logging.info(f"Added {len(matching_models)} models for Year: {year_int}, Make: {make_text}")
            else:
                logging.warning(f"No models found for Year: {year_int}, Make: {make_text}")
                
        except (ValueError, TypeError) as e:
            logging.error(f"Error in populate_models: {e}")

    def update_path_label(self):
        for config_name, label in self.path_labels.items():
            label.setText(f"{config_name} Path: {self.browse_buttons[config_name].text()}")

    def clear_all_data(self):
        # Show confirmation dialog using styled message box
        msg_box = self.create_styled_messagebox(
            "Clear All Data", 
            "Are you sure you want to clear ALL data?\nThis action cannot be undone.",
            QMessageBox.Question
        )
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.No)  # Default to No to prevent accidental clicks
        
        confirmation = msg_box.exec_()
        
        if confirmation == QMessageBox.Yes:
            # Clear all data types
            self.clear_data()  # This will call the clear_data method without a specific config_type
            
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

    def keyPressEvent(self, event):
        # Check for Ctrl+Shift+F1 shortcut
        if (event.modifiers() & Qt.ControlModifier) and (event.modifiers() & Qt.ShiftModifier) and event.key() == Qt.Key_F1:
            self.access_admin_panel()
        else:
            super().keyPressEvent(event)

    def apply_theme_to_expand_buttons(self):
        """Apply the current theme's color to expand buttons with a lighter shade."""
        theme = self.theme_dropdown.currentText()
        
        # Define theme colors mapping
        theme_colors = {
            "Red": ("#DC143C", "#A52A2A", "white"),
            "Blue": ("#4169E1", "#1E3A5F", "white"),
            "Dark": ("#2b2b2b", "#333333", "#ddd"),
            "Light": ("#f0f0f0", "#d0d0d0", "black"),
            "Green": ("#006400", "#003a00", "white"),
            "Yellow": ("#ffd700", "#b29500", "black"),
            "Pink": ("#ff69b4", "#b2477d", "black"),
            "Purple": ("#800080", "#4b004b", "white"),
            "Teal": ("#008080", "#004b4b", "white"),
            "Cyan": ("#00ffff", "#00b2b2", "black"),
            "Orange": ("#ff8c00", "#ff4500", "black")
        }
        
        # Get colors for the selected theme
        if theme in theme_colors:
            color_start, color_end, text_color = theme_colors[theme]
        else:
            # Default to Dark theme if theme not found
            color_start, color_end, text_color = theme_colors["Dark"]
            
        # Create lighter versions of the colors for a better look
        def lighten_color(hex_color, factor=0.3):
            """Lighten a color by the given factor (0-1)"""
            # Remove '#' if present
            hex_color = hex_color.lstrip('#')
            
            # Convert to RGB
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            
            # Lighten
            r = min(255, int(r + (255 - r) * factor))
            g = min(255, int(g + (255 - g) * factor))
            b = min(255, int(b + (255 - b) * factor))
            
            # Convert back to hex
            return f"#{r:02x}{g:02x}{b:02x}"
        
        # Create lighter versions of the theme colors
        light_color_start = lighten_color(color_start)
        light_color_end = lighten_color(color_end)
                
        # Create the button style with lighter colors
        button_style = f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {light_color_start}, stop:1 {light_color_end});
                color: {text_color};
                border-radius: 5px;
                font-size: 11px;
                font-weight: bold;
                padding: 2px 5px;
                max-height: 20px;
                margin-top: 0px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {light_color_end}, stop:1 {light_color_start});
            }}
        """
        
        # Apply the style to all expand buttons
        if hasattr(self, 'left_panel_pop_out_button'):
            self.left_panel_pop_out_button.setStyleSheet(button_style)
        
        if hasattr(self, 'right_panel_pop_out_button'):
            self.right_panel_pop_out_button.setStyleSheet(button_style)
            
        if hasattr(self, 'mag_glass_pop_out_button'):
            self.mag_glass_pop_out_button.setStyleSheet(button_style)
            
        if hasattr(self, 'carsys_pop_out_button'):
            self.carsys_pop_out_button.setStyleSheet(button_style)

class SignupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create Account")
        self.setGeometry(100, 100, 350, 250)  # Increased height for additional fields
        self.setup_ui()
        
        # Apply theme if parent has current_theme
        if parent and hasattr(parent, 'current_theme'):
            if parent.current_theme == "Dark":
                self.setStyleSheet(parent.get_dialog_dark_style())
            else:
                colors = parent.get_colors_for_theme(parent.current_theme)
                self.setStyleSheet(parent.get_dialog_color_style(*colors))

    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Add title
        title_label = QLabel("Create a New Account")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)
        
        # First Name input
        first_name_layout = QHBoxLayout()
        first_name_label = QLabel("First Name:")
        first_name_layout.addWidget(first_name_label)
        
        self.first_name_input = QLineEdit()
        self.first_name_input.setPlaceholderText("Enter first name")
        self.first_name_input.textChanged.connect(self.update_username_preview)
        first_name_layout.addWidget(self.first_name_input)
        layout.addLayout(first_name_layout)
        
        # Last Name input
        last_name_layout = QHBoxLayout()
        last_name_label = QLabel("Last Name:")
        last_name_layout.addWidget(last_name_label)
        
        self.last_name_input = QLineEdit()
        self.last_name_input.setPlaceholderText("Enter last name")
        self.last_name_input.textChanged.connect(self.update_username_preview)
        last_name_layout.addWidget(self.last_name_input)
        layout.addLayout(last_name_layout)
        
        # Username preview
        username_layout = QHBoxLayout()
        username_label = QLabel("Your username will be:")
        username_layout.addWidget(username_label)
        
        self.username_preview = QLabel("")
        self.username_preview.setStyleSheet("font-weight: bold; color: #3498db;")
        username_layout.addWidget(self.username_preview)
        layout.addLayout(username_layout)
        
        # Pin input
        pin_layout = QHBoxLayout()
        pin_label = QLabel("Enter PIN:")
        pin_layout.addWidget(pin_label)
        
        self.pin_input = QLineEdit()
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        self.pin_input.setPlaceholderText("Enter 4-digit PIN")
        pin_layout.addWidget(self.pin_input)
        layout.addLayout(pin_layout)
        
        # Confirm pin input
        confirm_pin_layout = QHBoxLayout()
        confirm_pin_label = QLabel("Confirm PIN:")
        confirm_pin_layout.addWidget(confirm_pin_label)
        
        self.confirm_pin_input = QLineEdit()
        self.confirm_pin_input.setEchoMode(QLineEdit.Password)
        self.confirm_pin_input.setMaxLength(4)
        self.confirm_pin_input.setPlaceholderText("Confirm PIN")
        confirm_pin_layout.addWidget(self.confirm_pin_input)
        layout.addLayout(confirm_pin_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.create_button = QPushButton("Create Account")
        self.create_button.clicked.connect(self.create_user)
        button_layout.addWidget(self.create_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def update_username_preview(self):
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        username = ""
        
        if first_name and last_name:
            username = first_name[0] + last_name
            self.username_preview.setText(username)
        else:
            self.username_preview.setText("(Example: JDoe)")
            
    def create_user(self):
        # Get input values
        pin = self.pin_input.text().strip()
        confirm_pin = self.confirm_pin_input.text().strip()
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        
        # Validate inputs
        if not pin:
            msg = self.parent().create_styled_messagebox("Error", "Please enter a PIN.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if not first_name or not last_name:
            msg = self.parent().create_styled_messagebox("Error", "Please enter both your first and last name.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if pin != confirm_pin:
            msg = self.parent().create_styled_messagebox("Error", "PINs do not match.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if len(pin) != 4 or not pin.isdigit():
            msg = self.parent().create_styled_messagebox("Error", "PIN must be 4 digits.", QMessageBox.Warning)
            msg.exec_()
            return
        
        # Generate username (first initial + last name)
        username = first_name[0] + last_name
        
        # Check if PIN already exists
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT name FROM leader_log WHERE pin = ?', (pin,))
        existing_user = cursor.fetchone()
        
        if existing_user:
            msg = self.parent().create_styled_messagebox("Error", "This PIN is already in use. Please choose another one.", QMessageBox.Warning)
            msg.exec_()
            conn.close()
            return
            
        # Check if username already exists
        cursor.execute('SELECT pin FROM leader_log WHERE name = ?', (username,))
        existing_username = cursor.fetchone()
        
        if existing_username:
            msg = self.parent().create_styled_messagebox("Error", f"An account with username '{username}' already exists. Please use a different name.", QMessageBox.Warning)
            msg.exec_()
            conn.close()
            return
        
        # Create the user with the generated username
        try:
            cursor.execute('INSERT INTO leader_log (pin, name) VALUES (?, ?)', (pin, username))
            conn.commit()
            self.parent().log_action("System", f"New user created: {username} (for {first_name} {last_name})")
            
            # Simple success message
            msg = self.parent().create_styled_messagebox("Success", "User created successfully!", QMessageBox.Information)
            msg.exec_()
            self.accept()
        except Exception as e:
            logging.error(f"Failed to create user: {e}")
            msg = self.parent().create_styled_messagebox("Error", f"Failed to create user: {str(e)}", QMessageBox.Critical)
            msg.exec_()
        finally:
            conn.close()

class ResetPinDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reset PIN")
        self.setGeometry(100, 100, 350, 250)  # Increased height for additional fields
        self.setup_ui()
        
        # Apply theme if parent has current_theme
        if parent and hasattr(parent, 'current_theme'):
            if parent.current_theme == "Dark":
                self.setStyleSheet(parent.get_dialog_dark_style())
            else:
                colors = parent.get_colors_for_theme(parent.current_theme)
                self.setStyleSheet(parent.get_dialog_color_style(*colors))

    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Add title
        title_label = QLabel("Reset Your PIN")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        font.setBold(True)
        title_label.setFont(font)
        layout.addWidget(title_label)
        
        # First Name input
        first_name_layout = QHBoxLayout()
        first_name_label = QLabel("First Name:")
        first_name_layout.addWidget(first_name_label)
        
        self.first_name_input = QLineEdit()
        self.first_name_input.setPlaceholderText("Enter first name")
        self.first_name_input.textChanged.connect(self.update_username_preview)
        first_name_layout.addWidget(self.first_name_input)
        layout.addLayout(first_name_layout)
        
        # Last Name input
        last_name_layout = QHBoxLayout()
        last_name_label = QLabel("Last Name:")
        last_name_layout.addWidget(last_name_label)
        
        self.last_name_input = QLineEdit()
        self.last_name_input.setPlaceholderText("Enter last name")
        self.last_name_input.textChanged.connect(self.update_username_preview)
        last_name_layout.addWidget(self.last_name_input)
        layout.addLayout(last_name_layout)
        
        # Username preview
        username_layout = QHBoxLayout()
        username_label = QLabel("Your username is:")
        username_layout.addWidget(username_label)
        
        self.username_preview = QLabel("")
        self.username_preview.setStyleSheet("font-weight: bold; color: #3498db;")
        username_layout.addWidget(self.username_preview)
        layout.addLayout(username_layout)
        
        # New Pin input
        pin_layout = QHBoxLayout()
        pin_label = QLabel("New PIN:")
        pin_layout.addWidget(pin_label)
        
        self.pin_input = QLineEdit()
        self.pin_input.setEchoMode(QLineEdit.Password)
        self.pin_input.setMaxLength(4)
        self.pin_input.setPlaceholderText("Enter new 4-digit PIN")
        pin_layout.addWidget(self.pin_input)
        layout.addLayout(pin_layout)
        
        # Confirm pin input
        confirm_pin_layout = QHBoxLayout()
        confirm_pin_label = QLabel("Confirm PIN:")
        confirm_pin_layout.addWidget(confirm_pin_label)
        
        self.confirm_pin_input = QLineEdit()
        self.confirm_pin_input.setEchoMode(QLineEdit.Password)
        self.confirm_pin_input.setMaxLength(4)
        self.confirm_pin_input.setPlaceholderText("Confirm new PIN")
        confirm_pin_layout.addWidget(self.confirm_pin_input)
        layout.addLayout(confirm_pin_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        
        self.reset_button = QPushButton("Reset PIN")
        self.reset_button.clicked.connect(self.reset_pin)
        button_layout.addWidget(self.reset_button)
        
        layout.addLayout(button_layout)
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
        # Get input values
        first_name = self.first_name_input.text().strip()
        last_name = self.last_name_input.text().strip()
        pin = self.pin_input.text().strip()
        confirm_pin = self.confirm_pin_input.text().strip()
        
        # Validate inputs
        if not first_name or not last_name:
            msg = self.parent().create_styled_messagebox("Error", "Please enter both your first and last name.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if not pin:
            msg = self.parent().create_styled_messagebox("Error", "Please enter a new PIN.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if pin != confirm_pin:
            msg = self.parent().create_styled_messagebox("Error", "PINs do not match.", QMessageBox.Warning)
            msg.exec_()
            return
        
        if len(pin) != 4 or not pin.isdigit():
            msg = self.parent().create_styled_messagebox("Error", "PIN must be 4 digits.", QMessageBox.Warning)
            msg.exec_()
            return
        
        # Generate expected username format
        expected_username = first_name[0] + last_name
        
        # Check if user exists
        conn = self.parent().get_db_connection()
        cursor = conn.cursor()
        
        # For debugging, log all available users
        debug_cursor = conn.cursor()
        debug_cursor.execute('SELECT pin, name FROM leader_log')
        all_users = debug_cursor.fetchall()
        logging.debug(f"All users in database for PIN reset: {all_users}")
        logging.debug(f"Looking for username: {expected_username}")
        
        cursor.execute('SELECT pin FROM leader_log WHERE name = ?', (expected_username,))
        existing_user = cursor.fetchone()
        
        if not existing_user:
            # Check if there might be a case sensitivity issue with the username
            cursor.execute('SELECT name FROM leader_log')
            all_usernames = cursor.fetchall()
            lowercase_usernames = [name[0].lower() for name in all_usernames]
            
            if expected_username.lower() in lowercase_usernames:
                # Found a match with different case
                for name in all_usernames:
                    if name[0].lower() == expected_username.lower():
                        actual_username = name[0]
                        msg = self.parent().create_styled_messagebox(
                            "Case Mismatch", 
                            f"Found your account with username '{actual_username}'. Please use that exact capitalization.",
                            QMessageBox.Information
                        )
                        msg.exec_()
                        conn.close()
                        return
            
            # Provide more helpful error message
            msg = self.parent().create_styled_messagebox(
                "Account Not Found", 
                f"No user found with the username '{expected_username}'.\n\n"
                "Please check that you've entered your name exactly as you did during registration, "
                "or create a new account if you haven't registered yet.",
                QMessageBox.Warning
            )
            msg.exec_()
            conn.close()
            return
        
        # Check if new PIN is already used by another user
        old_pin = existing_user[0]
        if pin != old_pin:  # Only check if it's a different PIN
            cursor.execute('SELECT name FROM leader_log WHERE pin = ? AND pin != ?', (pin, old_pin))
            pin_user = cursor.fetchone()
            if pin_user:
                msg = self.parent().create_styled_messagebox("Error", "This PIN is already in use by another user. Please choose a different one.", QMessageBox.Warning)
                msg.exec_()
                conn.close()
                return
        
        # Update the PIN
        try:
            cursor.execute('UPDATE leader_log SET pin = ? WHERE name = ?', (pin, expected_username))
            conn.commit()
            self.parent().log_action("System", f"PIN reset for user: {expected_username}")
            
            # Success message with clean simple format
            msg = self.parent().create_styled_messagebox("Success", "PIN reset successfully!", QMessageBox.Information)
            msg.exec_()
            self.accept()
        except Exception as e:
            logging.error(f"Failed to reset PIN: {e}")
            msg = self.parent().create_styled_messagebox("Error", f"Failed to reset PIN: {str(e)}", QMessageBox.Critical)
            msg.exec_()
        finally:
            conn.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())
