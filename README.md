<img src="./plategen_icon.ico" align="left">

# UPS Spec & Manual Generator

A powerful, **PyQt6** based desktop application that automates the generation of detailed **Technical Specifications** for Uninterruptible Power Supply (UPS) and Bypass Panel Systems. Leveraging the `docxtpl` engine, this utility rapidly compiles complex configuration, rating, I/O, battery, and environmental data into a professional, ready-to-use `.docx` specification document based on a customizable template.

Built for teams at **Liveline Electronics**.

[![Latest Release](https://img.shields.io/github/v/release/aamitn/manualgen?style=for-the-badge&logo=github&label=Latest%20Version)](https://github.com/aamitn/manualgen/releases)
[![License](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)

pyinstaller --noconfirm --onefile --windowed --icon=plategen_icon.ico --name=plategen app.py --collect-all requests

---

## ✨ Key Features

| Icon | Feature | Description |
| :---: | :--- | :--- |
| 📄 | **Template-Driven Generation** | Creates complex technical documents (`.docx`) from a single `template.docx` file using the `docxtpl` library. |
| 🧮 | **Dynamic Calculations** | Automatically computes key parameters like **Real Rating (kVA)** and **Battery Bus Voltage** based on input data. |
| 📊 | **Comprehensive Data Input** | Organized into **four thematic tabs** (General, I/O & Battery, Lists, Environment) to cover every required specification detail. |
| ➕➖ | **List Editor** | Intuitive list management for complex sections like **Protections**, **Metering Points**, **Audio Alarms**, and **Potential-Free Contacts**. |
| 💾 | **PDF Conversion** | One-click conversion of the last generated `.docx` file into a `.pdf` (requires the `docx2pdf` dependency). |
| 🔄 | **Version Check** | Built-in "About" dialog fetches and displays the latest available GitHub release version in real-time. |
| 📁 | **Auto-Open** | Optional setting to automatically open the generated `.docx` or `.pdf` file upon creation. |

---

## ⚙️ Dependencies & Installation (Development)

To set up the development environment and run the application:

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/aamitn/manualgen.git](https://github.com/aamitn/manualgen.git)
    cd manualgen # Replace with your actual project directory name if different
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv .venv
    .\.venv\Scripts\activate.bat 
    # OR (Linux/Mac)
    source ./.venv/bin/activate
    ```

3.  **Install Python Libraries:**
    ```bash
    pip install docxtpl PyQt6 requests
    # OPTIONAL: Install the PDF conversion dependency
    pip install docx2pdf
    ```

4.  **Template Requirement:** Ensure a file named **`template.docx`** is present in the application's root directory. This template must contain Jinja2 placeholders (e.g., `{{ full_spec_title }}`) corresponding to the context variables defined in `_generate_docx_file`.

5.  **Run the application:**
    ```bash
    python app.py
    ```

---

## 🚀 Build Executable (PyInstaller)

To create a standalone executable for deployment:

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Run the build command with custom naming:** Use the `--name` flag to specify the output filename, as discussed earlier.
    ```bash
    pyinstaller --noconfirm --onefile --windowed --icon=icon.ico --name=UPSManualGen app.py
    ```
    The built executable, **`UPSManualGen.exe`**, will be located in the **`dist/`** folder.

---

## 💻 Technical Highlights

### 1. Specification Numbering

The program automatically constructs a formal specification number based on the Job Number (`job_no`), OP Number (`op_no`), and the current year, ensuring compliance with internal document standards:

$$
\text{SPEC-No} = \text{TEC SPEC-}\{\text{job\_no}\}-\text{OP}\{\text{op\_no}\}-\{\text{YY}\}\text{UPS}3
$$

*(Where YY is the last two digits of the current year).*

### 2. Bypass Line Equipment Mapping

A robust dictionary handles the conversion between user-friendly equipment descriptions displayed in the ComboBox and the internal keys used in the template rendering logic. 

```python
ble_options_map = {
    "Isolation transformer with servo Stabilizer.": "stabilizer_iso",
    # ... more options
    "Integrated Power Distribution Unit (PDU)...": "integrated_pdu"
}
```

### 3. Asynchronous Version Checking
The GitHub API call to fetch the latest release tag is executed in a separate QThread (GithubVersionWorker) to prevent the main GUI from freezing, maintaining a smooth user experience.