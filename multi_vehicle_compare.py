from PyQt5.QtCore import Qt, pyqtSignal, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox,
    QFormLayout, QScrollArea, QFrame
)
from modern_components import ModernDialog, ModernComboBox, ModernButton, ModernTextBrowser
from database_utils import get_prequal_data, get_unique_years, get_unique_makes, get_unique_models
import logging

class VehicleSelector(QWidget):
    """A reusable vehicle selection widget"""
    vehicle_changed = pyqtSignal(int, str)  # index, field

    def __init__(self, parent=None, index=1):
        super().__init__(parent)
        self.parent = parent
        self.index = index
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Header with remove button
        header_layout = QHBoxLayout()
        header_label = QLabel(f"Vehicle {self.index}")
        header_label.setStyleSheet("font-weight: bold; color: #333;")
        header_layout.addWidget(header_label)
        
        if self.index > 2:  # Allow removing additional vehicles
            remove_btn = ModernButton("✕", style="danger")
            remove_btn.setFixedSize(24, 24)
            remove_btn.clicked.connect(self.remove_vehicle)
            header_layout.addWidget(remove_btn)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # Vehicle selection form
        form_layout = QFormLayout()
        form_layout.setSpacing(8)
        
        # Year selection
        self.year = ModernComboBox()
        self.year.addItem("Select Year")
        self.year.currentIndexChanged.connect(lambda: self.on_field_changed('year'))
        form_layout.addRow("Year:", self.year)
        
        # Make selection
        self.make = ModernComboBox()
        self.make.addItem("Select Make")
        self.make.currentIndexChanged.connect(lambda: self.on_field_changed('make'))
        form_layout.addRow("Make:", self.make)
        
        # Model selection
        self.model = ModernComboBox()
        self.model.addItem("Select Model")
        self.model.currentIndexChanged.connect(lambda: self.on_field_changed('model'))
        form_layout.addRow("Model:", self.model)
        
        layout.addLayout(form_layout)
        
        # Apply modern styling
        self.setStyleSheet("""
            QWidget {
                background: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
            }
        """)
        
    def get_selection(self):
        """Get the current vehicle selection"""
        return {
            'year': self.year.currentText(),
            'make': self.make.currentText(),
            'model': self.model.currentText()
        }
        
    def is_selection_complete(self):
        """Check if all fields are selected"""
        selection = self.get_selection()
        return all(selection[field] != f"Select {field.title()}" for field in selection)
        
    def on_field_changed(self, field):
        """Handle field change events"""
        self.vehicle_changed.emit(self.index, field)
            
    def update_models(self, year, make):
        """Update the models dropdown based on year and make"""
        self.model.clear()
        self.model.addItem("Select Model")
        
        if year == "Select Year" or make == "Select Make":
            return
            
        try:
            data = get_prequal_data()
            if data:
                # Get unique models for selected year and make
                models = get_unique_models(data, year, make)
                logging.debug(f"Found {len(models)} models for {year} {make}")
                logging.debug(f"Models: {models}")
                
                for model in models:
                    self.model.addItem(model)
        except Exception as e:
            logging.error(f"Error updating models: {e}")
            
    def remove_vehicle(self):
        """Remove this vehicle selector"""
        if self.parent and hasattr(self.parent, 'remove_vehicle_selector'):
            self.parent.remove_vehicle_selector(self.index)

class MultiVehicleCompareDialog(ModernDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Multi-Vehicle Comparison")
        self.setMinimumSize(1200, 800)
        self.vehicle_selectors = []
        self.setup_ui()
        self.add_initial_vehicles()
        
    def setup_ui(self):
        """Setup the user interface"""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header_layout = QHBoxLayout()
        title_label = QLabel("Multi-Vehicle Comparison")
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: 700;
            color: #333;
        """)
        
        subtitle_label = QLabel("Compare prequalification data between multiple vehicles")
        subtitle_label.setStyleSheet("""
            font-size: 12px;
            color: #666;
            font-style: italic;
        """)
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(subtitle_label)
        layout.addLayout(header_layout)
        
        # ADAS System filter
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filter by OEM ADAS System:")
        filter_label.setStyleSheet("""
            font-weight: bold;
            color: #333;
        """)
        self.adas_filter = ModernComboBox()
        self.adas_filter.addItem("All")
        self.adas_filter.currentTextChanged.connect(self.on_filter_changed)
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.adas_filter)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)
        
        # Vehicle selection area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background: transparent;
            }
        """)
        
        self.scroll_content = QWidget()
        self.vehicles_layout = QVBoxLayout(self.scroll_content)
        self.vehicles_layout.setSpacing(10)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area)
        
        # Add vehicle button
        add_vehicle_btn = ModernButton("+ Add Vehicle", style="success")
        add_vehicle_btn.clicked.connect(self.add_vehicle_selector)
        layout.addWidget(add_vehicle_btn, alignment=Qt.AlignCenter)
        
        # Compare button
        compare_btn = ModernButton("Compare Vehicles", style="primary")
        compare_btn.clicked.connect(self.compare_vehicles)
        layout.addWidget(compare_btn, alignment=Qt.AlignCenter)
        
        # Results display
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
        self.results_display.setMinimumHeight(400)
        self.results_display.setOpenExternalLinks(True)  # Allow opening links in external browser
        self.results_display.anchorClicked.connect(self.handle_link_click)
        layout.addWidget(self.results_display, 1)  # Give results more space
        
    def add_initial_vehicles(self):
        """Add the initial two vehicle selectors"""
        self.add_vehicle_selector()
        self.add_vehicle_selector()
        
    def add_vehicle_selector(self):
        """Add a new vehicle selector"""
        index = len(self.vehicle_selectors) + 1
        selector = VehicleSelector(self, index)
        selector.vehicle_changed.connect(self.on_vehicle_changed)
        
        # Add separator if not first vehicle
        if self.vehicle_selectors:
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setStyleSheet("background: #dee2e6; margin: 10px 0;")
            self.vehicles_layout.addWidget(separator)
        
        self.vehicles_layout.addWidget(selector)
        self.vehicle_selectors.append(selector)
        self.populate_dropdowns(selector)
        
    def remove_vehicle_selector(self, index):
        """Remove a vehicle selector"""
        if len(self.vehicle_selectors) <= 2:
            return  # Keep at least 2 vehicles
            
        # Find and remove the selector and its separator
        for i, selector in enumerate(self.vehicle_selectors):
            if selector.index == index:
                # Remove separator
                if i > 0:
                    separator = self.vehicles_layout.itemAt(i * 2 - 1).widget()
                    self.vehicles_layout.removeWidget(separator)
                    separator.deleteLater()
                
                # Remove selector
                self.vehicles_layout.removeWidget(selector)
                selector.deleteLater()
                self.vehicle_selectors.pop(i)
                break
                
        # Renumber remaining selectors
        for i, selector in enumerate(self.vehicle_selectors, 1):
            selector.index = i
            
    def populate_dropdowns(self, selector):
        """Populate the dropdowns for a vehicle selector"""
        try:
            data = get_prequal_data()
            if data:
                # Extract unique years and makes
                years = get_unique_years(data)
                makes = get_unique_makes(data)
                
                logging.debug(f"Found {len(years)} years and {len(makes)} makes")
                logging.debug(f"Years: {years}")
                logging.debug(f"Makes: {makes}")
                
                # Populate dropdowns
                selector.year.clear()
                selector.year.addItem("Select Year")
                for year in years:
                    selector.year.addItem(year)
                    
                selector.make.clear()
                selector.make.addItem("Select Make")
                for make in makes:
                    selector.make.addItem(make)
                    
        except Exception as e:
            logging.error(f"Error populating dropdowns: {e}")
            logging.exception("Full traceback:")
                
    def on_vehicle_changed(self, index, field):
        """Handle vehicle field changes"""
        selector = next((s for s in self.vehicle_selectors if s.index == index), None)
        if not selector:
            return
            
        if field in ['year', 'make']:
            year = selector.year.currentText()
            make = selector.make.currentText()
            if year != "Select Year" and make != "Select Make":
                selector.update_models(year, make)
                
        # Update ADAS systems filter whenever any vehicle selection changes
        if field in ['year', 'make', 'model']:
            self.update_adas_systems()
                
    def update_adas_systems(self):
        """Update the ADAS systems filter dropdown based on selected vehicles"""
        try:
            data = get_prequal_data()
            if not data:
                return
                
            # Get all complete vehicle selections
            vehicles = []
            for selector in self.vehicle_selectors:
                if selector.is_selection_complete():
                    vehicles.append(selector.get_selection())
                    
            if not vehicles:
                return
                
            # Get unique ADAS systems for all selected vehicles
            adas_systems = set()
            for vehicle in vehicles:
                for item in data:
                    try:
                        if (str(int(float(item['Year']))) == vehicle['year'] and
                            item['Make'].strip() == vehicle['make'].strip() and
                            item['Model'].strip() == vehicle['model'].strip() and
                            'Parent Component' in item and
                            item['Parent Component'] and
                            str(item['Parent Component']).lower() not in ['nan', 'none', 'null']):
                            adas_systems.add(str(item['Parent Component']).strip())
                    except (ValueError, TypeError):
                        continue
                        
            # Update dropdown
            current_text = self.adas_filter.currentText()
            self.adas_filter.clear()
            self.adas_filter.addItem("All")
            for system in sorted(adas_systems):
                self.adas_filter.addItem(system)
                
            # Restore previous selection if it still exists
            index = self.adas_filter.findText(current_text)
            if index >= 0:
                self.adas_filter.setCurrentIndex(index)
            else:
                self.adas_filter.setCurrentIndex(0)  # Default to "All"
                
        except Exception as e:
            logging.error(f"Error updating ADAS systems: {e}")
                
    def on_filter_changed(self):
        """Handle ADAS system filter changes"""
        self.compare_vehicles()
                
    def compare_vehicles(self):
        """Compare all selected vehicles"""
        # Validate selections
        vehicles = []
        for selector in self.vehicle_selectors:
            if not selector.is_selection_complete():
                self.results_display.setPlainText("Please complete vehicle information for all vehicles.")
                return
            vehicles.append(selector.get_selection())
            
        # Get data for all vehicles
        vehicle_data = []
        data = get_prequal_data()
        adas_filter = self.adas_filter.currentText()
        
        if data:
            for vehicle in vehicles:
                # Filter data for selected vehicle
                vehicle_systems = {}
                for item in data:
                    try:
                        if (str(int(float(item['Year']))) == vehicle['year'] and
                            item['Make'].strip() == vehicle['make'].strip() and
                            item['Model'].strip() == vehicle['model'].strip() and
                            'Parent Component' in item and
                            item['Parent Component'] and
                            str(item['Parent Component']).lower() not in ['nan', 'none', 'null']):
                            
                            system = str(item['Parent Component']).strip()
                            if adas_filter == "All" or system == adas_filter:
                                vehicle_systems[system] = item
                    except (ValueError, TypeError):
                        continue
                
                vehicle_data.append((vehicle, vehicle_systems))
                
        if not vehicle_data:
            self.results_display.setPlainText("No data found for the selected vehicles.")
            return
            
        # Generate and display comparison
        comparison = self.generate_comparison(vehicle_data, adas_filter)
        self.results_display.setHtml(comparison)
            
    def handle_link_click(self, url):
        """Handle clicking on links in the results display"""
        QDesktopServices.openUrl(url)

    def generate_comparison(self, vehicle_data, adas_filter):
        """Generate HTML comparison of multiple vehicles"""
        html = """
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 12px;
                text-align: left;
            }
            th {
                background-color: #f5f5f5;
                font-weight: bold;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            tr:hover {
                background-color: #f5f5f5;
            }
            .vehicle-header {
                background-color: #e9ecef;
                font-weight: bold;
                text-align: center;
            }
            .system-header {
                background-color: #e3f2fd;
                font-weight: bold;
                font-size: 16px;
                text-align: center;
                padding: 15px;
            }
            .feature-cell {
                font-weight: bold;
                background-color: #f8f9fa;
                width: 200px;
            }
            .value-cell {
                width: calc((100% - 200px) / NUM_VEHICLES);
            }
            .different {
                background-color: #fff3e0;
            }
            .not-available {
                color: #999;
                font-style: italic;
            }
            .hyperlink {
                color: #007bff;
                text-decoration: none;
            }
            .hyperlink:hover {
                text-decoration: underline;
            }
            .table-container {
                margin: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                overflow: hidden;
            }
            .table-wrapper {
                overflow-x: auto;
            }
            .table-title {
                font-size: 20px;
                font-weight: bold;
                color: #333;
                margin-bottom: 10px;
                padding: 0 20px;
            }
        </style>
        <div class="table-container">
        <div class="table-wrapper">
        """
        
        # Get all unique ADAS systems
        all_systems = set()
        for _, systems in vehicle_data:
            all_systems.update(systems.keys())
            
        # Filter systems if needed
        if adas_filter != "All":
            all_systems = {s for s in all_systems if s == adas_filter}
            
        if not all_systems:
            return "No ADAS systems found for the selected vehicles and filter."
            
        # Define the fields we want to display, in order
        display_fields = [
            'Year',
            'Make',
            'Model',
            'Parent Component',
            'Protech Generic System Name.1',  # Second instance of Protech Generic System Name
            'Calibration Type',
            'Calibration Pre-Requisites',
            'Calibration Pre-Requisites (Short Hand)',
            'Service Information Hyperlink',
            'Point of Impact #',
            'Component Generic Acronyms'
        ]
            
        # Vehicle headers
        html += "<table class='comparison-table'>"
        
        # For each ADAS system
        for system in sorted(all_systems):
            # System header
            html += f"<tr><td colspan='{len(vehicle_data) + 1}' class='system-header'>{system}</td></tr>"
            
            # Vehicle headers for this system
            html += "<tr><th>Feature</th>"
            for vehicle, _ in vehicle_data:
                html += f"<th class='vehicle-header'>{vehicle['year']} {vehicle['make']} {vehicle['model']}</th>"
            html += "</tr>"
            
            # Compare each field in our display list
            for field in display_fields:
                values = []
                for _, systems in vehicle_data:
                    if system in systems:
                        value = systems[system].get(field, 'N/A')
                        # Handle empty or NaN values
                        if value is None or str(value).lower() in ['nan', 'none', 'null', '']:
                            value = 'N/A'
                        values.append(str(value))
                    else:
                        values.append("Not available")
                        
                # Skip if all values are N/A
                if all(v == 'N/A' for v in values):
                    continue
                    
                # Check if values are different
                valid_values = [v for v in values if v not in ["Not available", "N/A"]]
                all_same = len(set(valid_values)) == 1 if valid_values else True
                cell_class = '' if all_same else 'different'
                
                # Format the field name for display
                display_field = field
                if field == 'Protech Generic System Name.1':
                    display_field = 'Protech Generic System Name'
                
                html += f"<tr><td class='feature-cell'>{display_field}</td>"
                for value in values:
                    if value == "Not available":
                        html += f"<td class='value-cell not-available'>System not available</td>"
                    elif value == "N/A":
                        html += f"<td class='value-cell not-available'>N/A</td>"
                    else:
                        # Make hyperlinks clickable and open in new tab
                        if field == 'Service Information Hyperlink' and value.lower() not in ['n/a', 'not available']:
                            value = f"<a href='{value}' class='hyperlink'>{value}</a>"
                        html += f"<td class='value-cell {cell_class}'>{value}</td>"
                html += "</tr>"
                
            # Add spacing between systems
            html += f"<tr><td colspan='{len(vehicle_data) + 1}' style='height: 20px;'></td></tr>"
                
        html += "</table></div></div>"
        return html
