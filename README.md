<img src="./installer/icons/plategen_icon.ico" align="left" width="75">

# Plategen : Industrial Plate Generation Suite

Plategen is a **modular desktop application** designed for automated generation of technical rating plates and nameplates for heavy industrial electrical equipment, including UPS, BCH, AC/DC DB panels. It provides **direct integration with AutoCAD via COM automation**, ensuring high precision and eliminating manual drafting errors in critical manufacturing documentation.


Built for teams at **Liveline Electronics**.

[![Download](https://img.shields.io/github/v/release/aamitn/plategen?style=for-the-badge&logo=github&label=Latest%20Version)](https://github.com/aamitn/plategen/releases/latest/download/PlateGenSetup.exe)
[![License](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)

---

## Features Overview

| Feature                            | Description                                                                                                              | Key Technology                          |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------ | --------------------------------------- |
| Modular Multi-App Architecture     | Separate executables for UPS, BCH, DB plates, Nameplate management, Stickers, and Technical Manual generation.           | PyInstaller, PyQt6                      |
| AutoCAD COM Automation             | Direct control of AutoCAD’s Object Model for fully automated industrial plate drawings.                                  | `win32com.client` (Python COM), AutoCAD |
| Structured Data Entry & Validation | Guided electrical and mechanical input forms with domain-level validation (PF, kVA/kW, circuit load, feeder parameters). | PyQt6, internal rules engine            |
| Export & Documentation             | Generate Excel, DOCX, and PDF outputs for shop-floor execution, handover formats, and QC workflows.                      | `openpyxl`, `reportlab`, `docxtpl`      |
| Nameplate Database Layer           | SQLite-backed storage enabling reusable plate definitions, reusable data groups, and repeat handling.                    | SQLite + DAO layer                      |
| Launcher Hub                       | Central controller for spawning sub-apps, detecting AutoCAD runtime, and maintaining clean lifecycle.                    | PyQt6, `psutil`, `subprocess`           |
| Sticker & Label Generator          | Create production-ready equipment stickers, tags, and safety labels with template-based sizing.                          | PyQt6 + drawing output                  |
| Technical Manual Generators        | Automatically compile DOCX-based technical handbooks for UPS and BCH using collected configuration and templates.        | `docxtpl`, templated reporting          |

---

##  Getting Started

### System Prerequisites
- **OS:** Windows 10/11 (Required for AutoCAD COM Interoperability)  
- **Python:** 3.8+  
- **CAD Software:** AutoCAD installation (Required for plate drawing)  

### Dependencies
Managed via `requirements.txt`:
- `PyQt6`
- `pywin32` (COM integration)
- `openpyxl`
- `reportlab`
- `psutil` (AutoCAD process management)

### Installation & Execution
```bash
# Clone the repository
git clone https://github.com/your-repo/plategen.git
cd plategen

# Install dependencies
pip install -r requirements.txt

# Run the Launcher
python app.py
```

---

##  Application Architecture

Plategen follows a **Micro-Application Architecture** with a central launcher orchestrating multiple specialized generator modules.

### System Architecture & External Interfaces
```mermaid
graph TD
    subgraph Core_Applications
        A[app.py - Launcher]
        B[app_bch.py - BCH Plate]
        C[app_ups.py - UPS Plate]
        D[app_db.py - DB Plate]
        E[app_np.py - Nameplate List]
    end

    subgraph External_Services
        F["AutoCAD COM Object Model (IUnknown)"]
        G[nameplates.db - SQLite]
        H[app_np_db_schema.py - DB Creator/Initializer]
    end

    A --> B
    A --> C
    A --> D
    A --> E

    B --> F
    C --> F
    D --> F

    E --> G
    H --> G

    style A fill:#f9f,stroke:#333
    style F fill:#add8e6,stroke:#333
    style G fill:#ccffcc,stroke:#333
```

### Module Responsibilities

| Module                | Responsibility                       | Technical Implementation                                                                           |
| --------------------- | ------------------------------------ | -------------------------------------------------------------------------------------------------- |
| `app.py`              | Lifecycle & State Management         | QApplication init, AutoCAD process checks (`psutil`), launches sub-apps via `subprocess`.          |
| `app_bch.py`          | BCH Rating Plate Generation          | Converts typed GUI inputs into config dict and draws block geometry in AutoCAD via COM.            |
| `app_ups.py`          | UPS Rating Plate Generation          | Performs electrical calculations (kW = kVA × PF), grid tiling, and AutoCAD model placement logic.  |
| `app_db.py`           | DB Rating Plate Generation           | Manages structured drawing for DB feeders, breaker layout, cable data, and coordinate positioning. |
| `app_np.py`           | Nameplate List I/O                   | SQLite DAL, grouping logic, repeats handling, Excel/PDF export via `openpyxl` / `reportlab`.       |
| `app_np_db_schema.py` | Database Schema Initialization       | Creates SQLite schema with `plate_types`, `ch_groups`, `nameplates` tables and seeds defaults.     |
| `app_sticker.py`      | Industrial Sticker & Label Generator | PyQt6 GUI for dimension-driven sticker generation and AutoCAD/print-ready output.                  |
| `app_mgen_ups.py`     | UPS Technical Manual Generator       | Produces DOCX technical manuals (`docxtpl`), processing plate spec + operational data.             |
| `app_mgen_bch.py`     | BCH Manual Design Tool (WIP)         | Currently displays an “Under Construction” Qt dialog; manual structuring and export planned.       |


##  Key Workflows

### 1. Rating Plate Drawing: COM Automation
```mermaid
sequenceDiagram
    participant U as User
    participant G as Plate Generator GUI
    participant P as pythoncom
    participant C as win32com.client
    participant A as AutoCAD Application

    U->>G: Enter configuration
    G->>P: Call CoInitialize()
    G->>C: `acad = Dispatch('AutoCAD.Application')`
    C-->>A: Access active document
    G->>G: Prepare Config Dict
    G->>G: Call draw_plates_grid(doc, Config Dict)
    G->>C: Invoke AutoCAD ModelSpace commands
    A-->>A: Execute drawing commands
    C-->>G: Return execution status
    G->>P: Call CoUninitialize()
    G->>U: Display status
```

### 2. Nameplate List Generation & Export
```mermaid
sequenceDiagram
    participant U as User
    participant N as Nameplate List GUI
    participant D as SQLite DAL
    participant I as In-Memory Logic
    participant E as Export Handler

    U->>N: Select Group ID
    N->>D: SELECT * FROM nameplates WHERE ch_group_id = ?
    D-->>N: Return raw record set
    N->>I: Process records (group, handle qty/repeater)
    I-->>N: Structured data
    U->>N: Click Export
    N->>E: Send data to openpyxl/reportlab
    E->>E: Generate styled Excel/PDF
    E-->>U: Save/Open generated file
```

### 3. Nameplate Database Schema
```mermaid
erDiagram
    plate_types ||--o{ nameplates : has
    ch_groups ||--o{ nameplates : contains

    plate_types {
        int id PK
        string type_name
        string default_size
    }

    ch_groups {
        int id PK
        string group_name
    }

    nameplates {
        int id PK
        int sl_no
        int type_id FK "FK references plate_types(id)"
        int ch_group_id FK "FK references ch_groups(id)"
        string name
        int qty
        int repeater "0=no repeat; N=number of repeated plates"
    }
```

---

## ️ Development Notes

### AutoCAD COM Interface
```python
import pythoncom
import win32com.client

try:
    pythoncom.CoInitialize()
    acad = win32com.client.Dispatch('AutoCAD.Application')
    doc = acad.ActiveDocument
    # Drawing logic
except Exception as e:
    pass
finally:
    pythoncom.CoUninitialize()
```

### Drawing Primitives
- **Lines:** `doc.ModelSpace.AddLine(StartPoint, EndPoint)`
- **Text:** `doc.ModelSpace.AddMText(InsertionPoint, Width, TextString)`  
Text must use predefined AutoCAD styles (`STYLE_REG`, `STYLE_BOLD`) in the drawing template.

### Database Logic (`app_np.py`)
- Ensures `nameplates.db` exists and is structured on startup.
- Handles repeater logic: `0` = one-off plate, `>0` = multiple sequential plates.

---

##  License

MIT License © [Bitmutex Technologies / Liveline Eletronics]

