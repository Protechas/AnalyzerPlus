from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QWidget, QPushButton, QComboBox, QTextBrowser, QDialog,
    QLineEdit, QProgressBar, QSlider, QTabWidget, QSplitter,
    QStatusBar, QToolBar
)

class ModernDialog(QDialog):
    """Base class for modern dialog styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QDialog {
                background: #ffffff;
            }
        """)

class ModernButton(QPushButton):
    """A modern button with hover effects and rounded corners"""
    def __init__(self, text="", parent=None, style="primary"):
        super().__init__(text, parent)
        self.style = style
        self.setup_style()

    def setup_style(self):
        colors = {
            'primary': ('#1976d2', '#1565c0'),
            'success': ('#4caf50', '#43a047'),
            'danger': ('#dc3545', '#c82333'),
            'warning': ('#ffc107', '#e0a800'),
            'info': ('#17a2b8', '#138496')
        }
        base_color, hover_color = colors.get(self.style, colors['primary'])
        
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {base_color};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
            QPushButton:pressed {{
                background-color: {base_color};
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
            }}
        """)

class ModernComboBox(QComboBox):
    """A modern combobox with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QComboBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px 10px;
                background: white;
                min-width: 6em;
            }
            QComboBox:hover {
                border-color: #80bdff;
            }
            QComboBox:focus {
                border-color: #80bdff;
                outline: 0;
                box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25);
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(down_arrow.png);
            }
        """)

class ModernTextBrowser(QTextBrowser):
    """A modern text browser with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QTextBrowser {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 10px;
                background: white;
                selection-background-color: #cce5ff;
            }
        """)

class ModernLineEdit(QLineEdit):
    """A modern line edit with custom styling"""
    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px 10px;
                background: white;
            }
            QLineEdit:hover {
                border-color: #80bdff;
            }
            QLineEdit:focus {
                border-color: #80bdff;
                outline: 0;
                box-shadow: 0 0 0 0.2rem rgba(0,123,255,.25);
            }
        """)

class ModernProgressBar(QProgressBar):
    """A modern progress bar with gradient styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                background: #e9ecef;
                height: 8px;
                text-align: center;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #007bff, stop:1 #6610f2);
                border-radius: 4px;
            }
        """)

class ModernSlider(QSlider):
    """A modern slider with custom styling"""
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QSlider::groove:horizontal {
                border: none;
                height: 4px;
                background: #e9ecef;
                margin: 2px 0;
            }
            QSlider::handle:horizontal {
                background: #007bff;
                border: none;
                width: 16px;
                margin: -6px 0;
                border-radius: 8px;
            }
            QSlider::handle:horizontal:hover {
                background: #0056b3;
            }
        """)

class ModernTabWidget(QTabWidget):
    """A modern tab widget with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                border-radius: 4px;
                background: white;
                top: -1px;
            }
            QTabBar::tab {
                background: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom-color: white;
            }
            QTabBar::tab:hover:!selected {
                background: #e9ecef;
            }
        """)

class ModernSplitter(QSplitter):
    """A modern splitter with custom styling"""
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QSplitter::handle {
                background: #dee2e6;
            }
            QSplitter::handle:horizontal {
                width: 4px;
            }
            QSplitter::handle:vertical {
                height: 4px;
            }
        """)

class ModernStatusBar(QStatusBar):
    """A modern status bar with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QStatusBar {
                background: #f8f9fa;
                border-top: 1px solid #dee2e6;
            }
            QStatusBar::item {
                border: none;
            }
        """)

class ModernToolBar(QToolBar):
    """A modern toolbar with custom styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()

    def setup_style(self):
        self.setStyleSheet("""
            QToolBar {
                background: #f8f9fa;
                border: none;
                spacing: 5px;
                padding: 5px;
            }
            QToolBar::separator {
                background: #dee2e6;
                width: 1px;
                margin: 5px;
            }
        """)
