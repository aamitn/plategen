import sys
import os
import subprocess
import requests
from datetime import datetime
from docxtpl import DocxTemplate
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLineEdit, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QGridLayout, QComboBox, QTabWidget, QListWidget,
    QMessageBox, QInputDialog, QSpinBox, QDoubleSpinBox, QFileDialog,
    QMenuBar, QMenu, QMainWindow, QDialog, QCheckBox
)
from PyQt6.QtCore import Qt, QSettings, QThread, pyqtSignal, QCoreApplication
from PyQt6.QtGui import QAction, QIcon


# --- Global Application Metadata ---
try:
    with open("appver.txt", "r") as f:
        APP_VERSION = f.read().strip()
except Exception:
        APP_VERSION = "v0.0.0"
APP_NAME = "Manual Generator"
COMPANY_NAME = "Bitmutex Technolgies"


GITHUB_REPO = "aamitn/plategen" # usename/reponame
GITHUB_URL_BASE = f"https://github.com/{GITHUB_REPO}/releases" # helper URL
# -----------------------------------

# --- 1. Document Generation Core Function ---
def _generate_docx_file(input_params, list_data, output_filepath):
    """Generates the DOCX file using the provided parameters."""
    
    # -----------------------------------------
    # Safely retrieve all list data using .get() to prevent KeyError
    # -----------------------------------------
    
    # Simple comma-separated lists
    metering_input_list = ", ".join(list_data.get("metering_input", []))
    metering_battery_list = ", ".join(list_data.get("metering_battery", [])) # FIXED: Safely access list data
    metering_output_list = ", ".join(list_data.get("metering_output", []))
    
    # Semicolon-separated indications
    indications_list = "; ".join(list_data.get("all_indications", []))
    
    # Newline-separated and numbered audio alarms
    audio_alarms = list_data.get("audio_alarms", [])
    audio_alarm_list = "\n".join([f"{i+1}. {alarm}" for i, alarm in enumerate(audio_alarms)])
    
    # Protection lists 
    rectifier_protections_list = ", ".join(list_data.get("rectifier_protections", []))
    inverter_protections_list = ", ".join(list_data.get("inverter_protections", []))

    # Potential Free Contacts
    signals = list_data.get("pot_free_contacts_signals", [])
    num_contacts = len(signals)
    signals_list = "\n".join([f"    {i+1}. {signal}" for i, signal in enumerate(signals)])
    pot_free_contacts_list = f"{num_contacts} Nos. provided for remote signal\n{signals_list}"

    
    # -----------------------------------------
    # UPS Config and Rating calculations
    # -----------------------------------------
    ups_config = input_params["ups_config"]
    rating = 0
    try:
        parts = ups_config.lower().split("x")
        rating = int(parts[1])
    except (ValueError, IndexError):
        rating = 8 # Default fallback

    ipf = input_params.get("ipf", 0.9)
    rating_real = round(rating * ipf, 2)
    ipf_rounded = round(ipf, 1)

    # -----------------------------------------
    # Battery voltage calculation
    # -----------------------------------------
    batno = input_params.get("batno", 1)
    batvpcell = input_params.get("batvpcell", 12)
    batv = batno * batvpcell
    
    # -----------------------------------------
    # Spec Number and Title Extraction/Construction
    # -----------------------------------------
    job_no = input_params.get("job_no", 0)
    op_no = input_params.get("op_no", 0)
    current_year_suffix = datetime.now().year % 100
    
    new_spec_number = f"TEC SPEC-{job_no}-OP{op_no}-{current_year_suffix:02d}UPS3"
    
    full_spec_title = (
        f"TECHNICAL SPECIFICATIONS OF ON-LINE STANDALONE {ups_config} KVA UPS & BYPASS PANEL SYSTEM "
        f"({input_params['phaseconfigfrom']}PH – {input_params['phaseconfigto']}PH)\n"
        f"SPECIFICATION NUMBER: {new_spec_number}"
    )

    # -----------------------------------------
    # Bypass Line Equipment Selection
    # -----------------------------------------
    ble_options_map = {
        "stabilizer_iso": "Isolation transformer with servo Stabilizer.",
        "iso_only": "Isolation Transformer Only.",
        "bypass_breaker": "External Maintenance Bypass Breaker Panel.",
        "none": "No additional bypass line equipment provided.",
        "static_switch_only": "Internal Static Switch for fast transfer without additional line conditioning.",
        "harmonic_filter": "Passive Harmonic Filter to meet required THDi limits.",
        "surge_protector": "High-capacity Surge Protection Device (SPD) integrated into the bypass line.",
        "integrated_pdu": "Integrated Power Distribution Unit (PDU) with output breakers and manual bypass switch."
    }
    
    ble_option_key = input_params.get("ble_option_key", "stabilizer_iso")
    bypass_line_equipment = ble_options_map.get(ble_option_key, ble_options_map["stabilizer_iso"])

    # -----------------------------------------
    # Build context for template
    # -----------------------------------------
    context = {
        "full_spec_title": full_spec_title,
        "new_spec_number": new_spec_number,
        "rev_no": input_params.get("rev_no", "00"),
        "doc_date": datetime.now().strftime("%d.%m.%y"),
        "ups_config": ups_config,
        "rating": rating,
        "rating_real": rating_real,
        "ipf_rounded": ipf_rounded,
        "bypass_line_equipment": bypass_line_equipment,
        
        # Input/Output
        "phaseconfigfrom": input_params.get("phaseconfigfrom", 3), "phaseconfigto": input_params.get("phaseconfigto", 1),
        "ivoltage": input_params.get("ivoltage", 400), "phase": input_params.get("phase", 3), "wire": input_params.get("wire", 4),
        "ivariationv": input_params.get("ivariationv", 10), "ifrequency": input_params.get("ifrequency", 50), "ivariationf": input_params.get("ivariationf", 5),
        "ipf": ipf, "ovoltage": input_params.get("ovoltage", 230), "ophase": input_params.get("ophase", 1), "owire": input_params.get("owire", 2),
        "ovariationv": input_params.get("ovariationv", 1.0), "ofrequency": input_params.get("ofrequency", 50), "ovariationf": input_params.get("ovariationf", 0.1),
        "ovariation_balanced": input_params.get("ovariation_balanced", 1.0), "ovariation_unbalanced": input_params.get("ovariation_unbalanced", 2.5),
        
        # Battery
        "batno": batno, "batcap": input_params.get("batcap", 0), "batvpcell": batvpcell, "batv": batv,
        
        # Lists
        "rectifier_protections_list": rectifier_protections_list,
        "inverter_protections_list": inverter_protections_list,
        "metering_input_list": metering_input_list,
        "metering_battery_list": metering_battery_list,
        "metering_output_list": metering_output_list,
        "indications_list": indications_list,
        "audio_alarm_list": audio_alarm_list,
        "pot_free_contacts_list": pot_free_contacts_list,
        
        # Environment
        "ip_rating": input_params.get("ip_rating", "IP 42"),
        "communication_protocol": input_params.get("communication_protocol", "RS 485 MODBUS RTU"),
        "optemplow": input_params.get("optemplow", 0), "optemphigh": input_params.get("optemphigh", 40),
        "altitude_factor": input_params.get("altitude_factor", 1000), "rh_percent": input_params.get("rh_percent", 95),
        "cooling_process": input_params.get("cooling_process", "Forced"), "cooling_type": input_params.get("cooling_type", "air"),
        "noise_power": input_params.get("noise_power", 65), "noise_distance_m": input_params.get("noise_distance_m", 1),
        "paint_process": input_params.get("paint_process", "7 Tank Process"),
    }

    # Render and save
    try:
        if not os.path.exists("template-mgen-ups.docx"):
            return False, "ERROR: The required 'template-mgen-ups.docx' file was not found in the application directory."
            
        doc = DocxTemplate("template-mgen-ups.docx")
        doc.render(context)
        doc.save(output_filepath)
        return True, f"DOCX Specification generated successfully to: {output_filepath}"
    except Exception as e:
        return False, f"An error occurred during DOCX rendering: {e}"

class GithubVersionWorker(QThread):
    """Worker thread to fetch the latest GitHub release version."""
    version_fetched = pyqtSignal(str) # Signal to send the result back to the GUI

    def __init__(self, repo_path):
        super().__init__()
        self.repo_path = repo_path
        
    def run(self):
        latest_version = "Failed to fetch latest version"
        api_url = f"https://api.github.com/repos/{self.repo_path}/releases/latest"
        
        try:
            response = requests.get(api_url, timeout=5) # 5 second timeout
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            data = response.json()
            
            # The 'tag_name' is usually the version number
            latest_version = data.get('tag_name', 'N/A')
            
        except requests.exceptions.RequestException as e:
            # Handle connection errors, timeouts, and bad status codes
            print(f"Error fetching GitHub release: {e}")
            latest_version = f"Error: {e.__class__.__name__}"
            
        self.version_fetched.emit(latest_version) # Emit the result when done
        
        
# --- 2. PyQt GUI Application (Main Window) ---
class UPSConfiguratorApp(QMainWindow):
    def __init__(self, default_params, default_list_data):
        super().__init__()
        self.default_params = default_params
        self.default_list_data = default_list_data
        self.settings = QSettings("MyCompany", "UPSConfigurator")
        self.last_docx_filepath = "" # To store the path for PDF conversion

        self.setWindowTitle("UPS Specification Configurator")
        self.setGeometry(100, 100, 1000, 700)
        
        # ---  WINDOW ICON  ---
        try:
            self.setWindowIcon(QIcon.fromTheme("document-new"))
        except Exception as e:
            # Handle case where icon file might be missing
            print(f"Warning: Could not load window icon: {e}") 
        # ----------------------------
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.widgets = {} # For input fields
        self.list_widgets = {} # For list editors
        self.ble_options = { # Map display text to internal keys
            "Isolation transformer with servo Stabilizer.": "stabilizer_iso",
            "Isolation Transformer Only.": "iso_only",
            "External Maintenance Bypass Breaker Panel.": "bypass_breaker",
            "No additional bypass line equipment provided.": "none",
            "Internal Static Switch for fast transfer without additional line conditioning.": "static_switch_only",
            "Passive Harmonic Filter to meet required THDi limits.": "harmonic_filter",
            "High-capacity Surge Protection Device (SPD) integrated into the bypass line.": "surge_protector",
            "Integrated Power Distribution Unit (PDU) with output breakers and manual bypass switch.": "integrated_pdu"
        }
        
        self.init_ui()
        self.init_menubar()
        
    def init_menubar(self):
        menubar = self.menuBar()

        # --- Settings Menu ---
        settings_menu = menubar.addMenu("&Settings")
        
        self.auto_open_action = QAction("Auto-Open Generated File", self)
        self.auto_open_action.setCheckable(True)
        # Load saved setting or default to checked
        default_auto_open = self.settings.value("auto_open_file", True, type=bool) 
        self.auto_open_action.setChecked(default_auto_open)
        self.auto_open_action.triggered.connect(self.save_settings)
        settings_menu.addAction(self.auto_open_action)

        # --- About Menu ---
        about_menu = menubar.addMenu("&Help")
        
        help_action = QAction("&About", self)
        help_action.triggered.connect(self.show_about_dialog)
        about_menu.addAction(help_action)

    def save_settings(self):
        """Saves the current state of the Auto-Open checkbox."""
        self.settings.setValue("auto_open_file", self.auto_open_action.isChecked())

    def show_about_dialog(self):
        """Displays a professional dialog with application info and initiates version check."""
        
        # 1. HTML Formatted Initial Message
        initial_message = f"""
        <html>
        <head>
            <style>
                h3 {{ color: #2e8b57; margin-bottom: 5px; }}
                p {{ margin: 3px 0; }}
                .section-header {{ font-weight: bold; margin-top: 10px; color: #3498db; }}
                .app-info {{ font-weight: bold; }}
            </style>
        </head>
        <body>
            <h3>{APP_NAME}</h3>
            <p class="app-info">Version: {APP_VERSION}</p>
            <p class="app-info">Company: {COMPANY_NAME}</p>
            
            <hr>

            <div class="section-header">Required Dependencies</div>
            <ul>
                <li><code>docxtpl</code> (for DOCX generation)</li>
                <li><code>docx2pdf</code> (for PDF generation)</li>
            </ul>
            <p>Installation: <code>pip install docxtpl docx2pdf</code></p>
            
            <hr>

            <div class="section-header">Template Requirement</div>
            <p>Requires 'template.docx' in the application directory.</p>
            
            <hr>
            
            <p id="github-version"><b>Latest GitHub Version:</b> <i>Checking...</i></p>
        </body>
        </html>
        """

        # We must use QMessageBox.information to get a reference to the box
        self.about_box = QMessageBox(
            self,
            windowTitle=f"About {APP_NAME}",
            text=initial_message,
            standardButtons=QMessageBox.StandardButton.Ok
        )
        
        # Crucial step: Set the text format to RichText (HTML)
        self.about_box.setTextFormat(Qt.TextFormat.RichText)
        
        # Set text as selectable to allow copying dependency info
        self.about_box.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse | 
            Qt.TextInteractionFlag.TextSelectableByKeyboard |
            Qt.TextInteractionFlag.LinksAccessibleByMouse # Allow links to be clicked
        )
        
        self.about_box.show() # Show the dialog immediately

        # 2. Start the GitHub version check in a thread
        self.github_worker = GithubVersionWorker(GITHUB_REPO)
        self.github_worker.version_fetched.connect(self.update_about_dialog)
        self.github_worker.start()

    def update_about_dialog(self, latest_version):
        """Updates the QMessageBox with the fetched GitHub version and makes it a link."""
        
        # Get the current HTML text
        current_html = self.about_box.text()
        
        placeholder = '<p id="github-version"><b>Latest GitHub Version:</b> <i>Checking...</i></p>'
        
        # 1. Construct the link/version text
        if latest_version.startswith("Error"):
            # If error, display the error text
            version_html = f'<p id="github-version"><b>Latest GitHub Version:</b> <span style="color: red;">{latest_version}</span></p>'
        else:
            # Construct the clickable hyperlink
            link_url = f"{GITHUB_URL_BASE}/tag/{latest_version}"
            version_html = f"""
            <p id="github-version">
                <b>Latest GitHub Version:</b> 
                <a href="{link_url}"><u>{latest_version}</u></a>
            </p>
            """
            
        # 2. Update the QMessageBox content
        updated_html = current_html.replace(placeholder, version_html)
        
        self.about_box.setText(updated_html)
    def init_ui(self):
        main_layout = QVBoxLayout(self.central_widget)

        # Title
        title = QLabel("UPS Specification Generator")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 10px;")
        main_layout.addWidget(title)

        # Tab Widget
        tabs = QTabWidget()
        tabs.addTab(self.create_general_tab(), "General & Ratings")
        tabs.addTab(self.create_io_battery_tab(), "I/O & Battery")
        tabs.addTab(self.create_lists_tab(), "Protections & Lists")
        tabs.addTab(self.create_environment_tab(), "Environment")
        main_layout.addWidget(tabs)

        # Action Buttons Layout
        action_layout = QHBoxLayout()
        
        # Generate DOCX Button
        generate_docx_btn = QPushButton("Generate DOCX Specification")
        generate_docx_btn.setStyleSheet("background-color: #2e8b57; color: white; padding: 10px; font-size: 16px; border-radius: 8px;")
        generate_docx_btn.clicked.connect(self.generate_docx)
        action_layout.addWidget(generate_docx_btn)
        
        # Convert to PDF Button
        generate_pdf_btn = QPushButton("Convert Last Generated DOCX to PDF")
        generate_pdf_btn.setStyleSheet("background-color: #3498db; color: white; padding: 10px; font-size: 16px; border-radius: 8px;")
        generate_pdf_btn.clicked.connect(self.convert_to_pdf)
        action_layout.addWidget(generate_pdf_btn)
        
        main_layout.addLayout(action_layout)
        
        # Instructions/Warnings
        warning_label = QLabel("Note: Requires 'template.docx' in the same directory. PDF generation requires 'docx2pdf'.")
        warning_label.setStyleSheet("color: #d35400; font-style: italic;")
        main_layout.addWidget(warning_label)


    # --- Tab 1: General & Ratings ---
    def create_general_tab(self):
        tab = QWidget()
        layout = QGridLayout(tab)

        fields = [
            ("UPS Config (e.g., 1x8):", "ups_config", QLineEdit, str, None),
            ("Input Power Factor (IPF):", "ipf", QDoubleSpinBox, float, (0.01, 1.0, 0.01)),
            ("Job Number:", "job_no", QSpinBox, int, (1, 9999, 1)),
            ("OP Number:", "op_no", QSpinBox, int, (1, 9999, 1)),
            ("Revision No:", "rev_no", QLineEdit, str, None),
        ]

        for i, (label_text, key, widget_class, data_type, spin_range) in enumerate(fields):
            label = QLabel(label_text)
            
            if widget_class == QDoubleSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setSingleStep(spin_range[2])
                widget.setValue(self.default_params.get(key, 0.0))
            elif widget_class == QSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setValue(self.default_params.get(key, 0))
                widget.setMaximum(9999)
            else:
                widget = widget_class()
                widget.setText(str(self.default_params.get(key, "")))
                
            self.widgets[key] = widget
            layout.addWidget(label, i, 0)
            layout.addWidget(widget, i, 1)
        
        # Bypass Line Equipment (BLE) ComboBox
        ble_label = QLabel("Bypass Line Equipment:")
        ble_combo = QComboBox()
        ble_combo.addItems(list(self.ble_options.keys()))
        
        default_ble_key = self.default_params.get("ble_option_key", "harmonic_filter")
        # Find the display text corresponding to the default key
        default_ble_text = next((k for k, v in self.ble_options.items() if v == default_ble_key), list(self.ble_options.keys())[0])
        ble_combo.setCurrentText(default_ble_text)
        
        self.widgets["ble_option_key"] = ble_combo
        layout.addWidget(ble_label, len(fields), 0)
        layout.addWidget(ble_combo, len(fields), 1)
        
        layout.setRowStretch(len(fields) + 1, 1)
        return tab

    # --- Tab 2: I/O & Battery (Helper method for group layouts) ---
    def _create_group_layout(self, title, fields):
        v_layout = QVBoxLayout()
        v_layout.addWidget(QLabel(f"<span style='font-weight: bold; font-size: 14px;'>{title}</span>"))
        
        g_layout = QGridLayout()
        
        # Note: All fields in this section are QSpinBox/QDoubleSpinBox
        for i, (label_text, key, widget_class, data_type, spin_range) in enumerate(fields):
            label = QLabel(label_text)
            
            if widget_class == QSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setValue(self.default_params.get(key, 0))
                widget.setMaximum(9999)
            elif widget_class == QDoubleSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setSingleStep(spin_range[2])
                widget.setValue(self.default_params.get(key, 0.0))
            
            self.widgets[key] = widget
            g_layout.addWidget(label, i, 0)
            g_layout.addWidget(widget, i, 1)

        v_layout.addLayout(g_layout)
        v_layout.addStretch(1)
        return v_layout

    def create_io_battery_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)

        input_group = self._create_group_layout("Input Parameters", [
            ("Phase Config From (PH):", "phaseconfigfrom", QSpinBox, int, (1, 3, 1)),
            ("Phase Config To (PH):", "phaseconfigto", QSpinBox, int, (1, 3, 1)),
            ("Voltage (V):", "ivoltage", QSpinBox, int, (100, 600, 1)),
            ("Phase Count:", "phase", QSpinBox, int, (1, 3, 1)),
            ("Wire Count:", "wire", QSpinBox, int, (1, 5, 1)),
            ("Voltage Variation (%):", "ivariationv", QSpinBox, int, (1, 20, 1)),
            ("Frequency (Hz):", "ifrequency", QSpinBox, int, (40, 60, 1)),
            ("Frequency Variation (%):", "ivariationf", QSpinBox, int, (1, 10, 1)),
        ])

        output_group = self._create_group_layout("Output Parameters", [
            ("Voltage (V):", "ovoltage", QSpinBox, int, (100, 400, 1)),
            ("Phase Count:", "ophase", QSpinBox, int, (1, 3, 1)),
            ("Wire Count:", "owire", QSpinBox, int, (1, 5, 1)),
            ("Voltage Variation (%):", "ovariationv", QDoubleSpinBox, float, (0.1, 5.0, 0.1)),
            ("Frequency (Hz):", "ofrequency", QSpinBox, int, (40, 60, 1)),
            ("Frequency Variation (%):", "ovariationf", QDoubleSpinBox, float, (0.01, 1.0, 0.01)),
            ("Regulation Balanced (%):", "ovariation_balanced", QDoubleSpinBox, float, (0.1, 5.0, 0.1)),
            ("Regulation Unbalanced (%):", "ovariation_unbalanced", QDoubleSpinBox, float, (0.1, 5.0, 0.1)),
        ])

        battery_group = self._create_group_layout("Battery Parameters", [
            ("Number of Batteries (Nos.):", "batno", QSpinBox, int, (1, 100, 1)),
            ("Capacity (Ah):", "batcap", QSpinBox, int, (1, 500, 1)),
            ("Voltage per Cell (V/Cell):", "batvpcell", QSpinBox, int, (2, 12, 1)),
        ])
        
        layout.addLayout(input_group)
        layout.addLayout(output_group)
        layout.addLayout(battery_group)
        return tab


    # --- Tab 3: Protections & Lists ---
    def create_lists_tab(self):
        tab = QWidget()
        layout = QGridLayout(tab)

        list_keys = [
            ("Rectifier Protections", "rectifier_protections", 0, 0),
            ("Inverter Protections", "inverter_protections", 0, 1),
            ("Metering (Input)", "metering_input", 1, 0),
            ("Metering (Output)", "metering_output", 1, 1),
            ("Indications (MIMIC)", "all_indications", 2, 0),
            ("Audio Alarms", "audio_alarms", 2, 1),
            ("Potential Free Contacts", "pot_free_contacts_signals", 3, 0)
        ]

        for title, key, row, col in list_keys:
            # Safely fetch default data for list initialization
            default_data = self.default_list_data.get(key, [])
            list_layout = self._create_list_editor(title, key, default_data)
            layout.addLayout(list_layout, row, col)

        layout.setRowStretch(4, 1)
        layout.setColumnStretch(0, 1)
        layout.setColumnStretch(1, 1)
        return tab

    # --- Tab 4: Environment ---
    def create_environment_tab(self):
        tab = QWidget()
        layout = QGridLayout(tab)
        
        fields = [
            ("Enclosure (IP Rating):", "ip_rating", QLineEdit, str, None),
            ("Communication Protocol:", "communication_protocol", QLineEdit, str, None),
            ("Operating Temp Low (°C):", "optemplow", QSpinBox, int, (-50, 0, 1)),
            ("Operating Temp High (°C):", "optemphigh", QSpinBox, int, (0, 70, 1)),
            ("Altitude Factor (M):", "altitude_factor", QSpinBox, int, (100, 5000, 1)),
            ("Humidity (% RH):", "rh_percent", QSpinBox, int, (10, 100, 1)),
            ("Cooling Process:", "cooling_process", QLineEdit, str, None),
            ("Cooling Type:", "cooling_type", QLineEdit, str, None),
            ("Noise Power (dbA):", "noise_power", QSpinBox, int, (30, 80, 1)),
            ("Noise Distance (M):", "noise_distance_m", QSpinBox, int, (1, 10, 1)),
            ("Painting Procedure:", "paint_process", QLineEdit, str, None),
        ]
        
        for i, (label_text, key, widget_class, data_type, spin_range) in enumerate(fields):
            row, col = divmod(i, 2)
            
            label = QLabel(label_text)
            
            if widget_class == QSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setValue(self.default_params.get(key, 0))
            elif widget_class == QDoubleSpinBox:
                widget = widget_class()
                widget.setRange(spin_range[0], spin_range[1])
                widget.setSingleStep(spin_range[2])
                widget.setValue(self.default_params.get(key, 0.0))
            else:
                widget = widget_class()
                widget.setText(str(self.default_params.get(key, "")))
                
            self.widgets[key] = widget
            layout.addWidget(label, row, col * 2)
            layout.addWidget(widget, row, col * 2 + 1)
        
        layout.setRowStretch(len(fields) // 2 + 1, 1)
        return tab


    # --- List Editor Widget Creation and Logic ---
    def _create_list_editor(self, title, key, default_items):
        v_layout = QVBoxLayout()
        v_layout.addWidget(QLabel(f"<span style='font-weight: bold;'>{title}</span>"))

        list_widget = QListWidget()
        list_widget.addItems(default_items)
        list_widget.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_widgets[key] = list_widget
        v_layout.addWidget(list_widget)

        h_layout = QHBoxLayout()
        add_btn = QPushButton("Add")
        remove_btn = QPushButton("Remove")
        edit_btn = QPushButton("Edit")

        add_btn.clicked.connect(lambda: self.add_list_item(key))
        remove_btn.clicked.connect(lambda: self.remove_list_item(key))
        edit_btn.clicked.connect(lambda: self.edit_list_item(key))
        
        h_layout.addWidget(add_btn)
        h_layout.addWidget(edit_btn)
        h_layout.addWidget(remove_btn)
        v_layout.addLayout(h_layout)

        return v_layout

    def add_list_item(self, key):
        list_widget = self.list_widgets[key]
        item_text, ok = QInputDialog.getText(self, f"Add {key.replace('_', ' ').title()} Item", "New Item Description:")
        if ok and item_text:
            list_widget.addItem(item_text)

    def remove_list_item(self, key):
        list_widget = self.list_widgets[key]
        for item in list_widget.selectedItems():
            list_widget.takeItem(list_widget.row(item))

    def edit_list_item(self, key):
        list_widget = self.list_widgets[key]
        selected_items = list_widget.selectedItems()
        if not selected_items:
            return

        current_item = selected_items[0]
        current_text = current_item.text()
        
        new_text, ok = QInputDialog.getText(self, f"Edit {key.replace('_', ' ').title()} Item", "Edit Item Description:", text=current_text)
        
        if ok and new_text:
            current_item.setText(new_text)

    # --- Data Collection ---
    def collect_data(self):
        # Collect single-value parameters
        input_params = {}
        for key, widget in self.widgets.items():
            if key == "ble_option_key":
                # Get the internal key for the selected description
                input_params[key] = self.ble_options[widget.currentText()]
            elif isinstance(widget, QSpinBox):
                input_params[key] = widget.value()
            elif isinstance(widget, QDoubleSpinBox):
                input_params[key] = widget.value()
            else: # QLineEdit
                input_params[key] = widget.text() # Keep as string for now
        
        # Collect list data
        list_data = {}
        for key, list_widget in self.list_widgets.items():
            list_data[key] = [list_widget.item(i).text() for i in range(list_widget.count())]
            
        return input_params, list_data

    # --- Report Generation Handlers ---
    def generate_docx(self):
        input_params, list_data = self.collect_data()
        
        # Open file dialog to choose save location
        default_filename = "UPS_SPEC_OUTPUT.docx"
        filepath, _ = QFileDialog.getSaveFileName(self, "Save Specification", default_filename, "Word Document (*.docx)")

        if not filepath:
            return # User cancelled

        success, message = _generate_docx_file(input_params, list_data, filepath)
        
        if success:
            self.last_docx_filepath = filepath # Store path for PDF conversion
            QMessageBox.information(self, "Success", message)
            
            if self.auto_open_action.isChecked():
                self.open_file(filepath)
        else:
            QMessageBox.critical(self, "Error", message)

    def convert_to_pdf(self):
        if not self.last_docx_filepath or not os.path.exists(self.last_docx_filepath):
            QMessageBox.warning(self, "Warning", "Please generate and save the DOCX file first.")
            return

        try:
            from docx2pdf import convert
        except ImportError:
            QMessageBox.critical(
                self, 
                "Dependency Error", 
                "The 'docx2pdf' library is not installed.\n"
                "Please run: `pip install docx2pdf`"
            )
            return

        pdf_filepath = self.last_docx_filepath.replace(".docx", ".pdf")
        
        try:
            # Open file dialog to choose save location for PDF
            pdf_filepath, _ = QFileDialog.getSaveFileName(
                self, 
                "Save PDF Specification", 
                pdf_filepath, 
                "PDF Document (*.pdf)"
            )
            
            if not pdf_filepath:
                return # User cancelled
                
            convert(self.last_docx_filepath, pdf_filepath)
            
            QMessageBox.information(self, "Success", f"PDF successfully generated to: {pdf_filepath}")
            
            if self.auto_open_action.isChecked():
                self.open_file(pdf_filepath)

        except Exception as e:
            QMessageBox.critical(self, "PDF Conversion Error", f"An error occurred during PDF conversion: {e}\n\n"
                                                                "Ensure Microsoft Word is installed and closed.")

    def open_file(self, filepath):
        """Opens a file using the default system application."""
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin": # macOS
            subprocess.call(("open", filepath))
        else: # linux
            subprocess.call(("xdg-open", filepath))


# -----------------------------------------
# 3. Default Data Setup
# -----------------------------------------

DEFAULT_PARAMS = {
    "ups_config": "1x8", "ipf": 0.92,
    "job_no": 967, "op_no": 1972, "rev_no": "06",
    "ble_option_key": "harmonic_filter", 
    
    # Input
    "phaseconfigfrom": 3, "phaseconfigto": 1, "ivoltage": 400, "phase": 3, "wire": 4,
    "ivariationv": 10, "ifrequency": 50, "ivariationf": 5,
    
    # Output
    "ovoltage": 230, "ophase": 1, "owire": 2, "ovariationv": 1.0, "ofrequency": 50,
    "ovariationf": 0.1, "ovariation_balanced": 1.0, "ovariation_unbalanced": 2.5,
    
    # Battery
    "batno": 14, "batcap": 65, "batvpcell": 12,
    
    # Environment/Details
    "ip_rating": "IP 42", "communication_protocol": "RS 485 MODBUS RTU",
    "optemplow": -15, "optemphigh": 50, "altitude_factor": 1302, "rh_percent": 95,
    "cooling_process": "Forced", "cooling_type": "air", "noise_power": 55,
    "noise_distance_m": 1, "paint_process": "7 Tank Process",
}

DEFAULT_LIST_DATA = {
    "rectifier_protections": [
        "Input Over/Under Voltage Protection", "Input Surge Protection (MOVs on Input Lines)",
        "Input Current Limit/Soft Start Feature", "Input Overcurrent Protection (Internal Fuses/Breaker)",
        "Rectifier Over Temperature Protection"
    ],
    "inverter_protections": [
        "Output Overload Protection (e.g., 150% for 1 minute)", "Output Short Circuit Protection (Instantaneous trip)",
        "Output Over/Under Voltage Protection", "Inverter DC Bus Over/Under Voltage Protection",
        "Heat Sink Over Temperature Shutdown", "Inverter PWM Overcurrent Protection"
    ],
    "metering_input": ["Voltage", "Current", "Frequency"],
    "metering_battery": ["Voltage", "Charge/Dis-charge Current"],
    "metering_output": ["Voltage", "Current", "Frequency", "O/P PF", "kVA", "kW", "Load%", "Temp"],
    "all_indications": [
        "Mains: ON/Fail, Low, High, Single Phasing", "DC: On, High, Low, Trip",
        "Battery: Low Alarm, High Alarm, Trip", "Inverter: On, Trip, Low, High, Overload",
        "Bypass: On, Off"
    ],
    "audio_alarms": [
        "Mains Fail", "Battery Low Pre-alarm", "Battery Low Trip", "Inverter Trip for Output Voltage Low",
        "Inverter Trip for Output Voltage High", "Inverter Trip for Output Over Current",
        "Inverter Trip for Output Over Temp", "Earth Fault Alarm", "Battery Charging Over Current",
        "Over Voltage / Under Voltage (Output)", "Input Over Voltage / Under Voltage",
        "Under / Over Frequency", "Output Short Circuit"
    ],
    "pot_free_contacts_signals": [
        "Inverter Trip", "Inverter ON", "Battery Low", "Mains Fail", "Mains ON"
    ]
}

# -----------------------------------------
# 4. Main Execution
# -----------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = UPSConfiguratorApp(DEFAULT_PARAMS, DEFAULT_LIST_DATA)
    window.show()
    sys.exit(app.exec())