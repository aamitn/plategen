import sys
import os
import subprocess
import threading
import urllib.request
import json
import webbrowser
import time

# Default embedded PNG icon (tiny fallback). Place `plategen_icon.png` next
DEFAULT_ICON_B64 = ("iVBORw0KG")

ICON_FILENAME = "plategen_icon.png"

def ensure_app_icon():
    import sys, os, base64

    # Determine application runtime folder
    if hasattr(sys, '_MEIPASS'):
        # Running from PyInstaller bundle
        run_dir = sys._MEIPASS
    else:
        # Running from source
        run_dir = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(run_dir, ICON_FILENAME)

    # If running in PyInstaller mode, just return it (no write allowed)
    if hasattr(sys, '_MEIPASS'):
        return icon_path if os.path.exists(icon_path) else None

    # Running from source → ensure icon exists, generate if missing
    if not os.path.exists(icon_path):
        try:
            with open(icon_path, 'wb') as f:
                f.write(base64.b64decode(DEFAULT_ICON_B64))
        except Exception:
            return None

    return icon_path


# TASKBAR ICON TWEAK FOR WINDOWS
def set_windows_app_id():
    try:
        from ctypes import windll
        APP_ID = "com.bitmutex.plategen"
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


# --- Third-party dependencies needed for AutoCAD control ---
# psutil is needed for reliable process termination
try:
    import psutil
except ImportError:
    psutil = None
    print("Warning: psutil not found. Process termination will not be fully supported.")
    
# pywin32 COM libraries for Windows (AutoCAD interaction)
try:
    import win32com.client
    import pythoncom
    COM_AVAILABLE = True
except Exception:
    win32com = None
    pythoncom = None
    COM_AVAILABLE = False
    print("Warning: pywin32 not found. COM interaction is disabled.")


from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QGroupBox, QGridLayout, QMessageBox, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal, QSettings, QTimer
from PyQt6.QtGui import QIcon, QAction

APPVER_FILE = os.path.join(os.path.dirname(__file__), 'appver.txt')
DEFAULT_GITHUB_REPO = 'aamitn/winhider'


def read_local_version():
    try:
        with open(APPVER_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return 'v0.0.0'


def fetch_latest_github_release(repo, timeout=8):
    """Return (tag, html_url, err) for latest release of repo 'owner/repo'."""
    if not repo or '/' not in repo:
        return None, None, 'No repo configured'
    url = f'https://api.github.com/repos/{repo}/releases/latest'
    req = urllib.request.Request(url, headers={'User-Agent': 'plategen-launcher'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status != 200:
                return None, None, f'HTTP {resp.status}'
            data = resp.read().decode('utf-8')
            j = json.loads(data)
            tag = j.get('tag_name') or j.get('name')
            html = j.get('html_url')
            return tag, html, None
    except Exception as e:
        return None, None, str(e)

def get_app_dir():
    if getattr(sys, 'frozen', False):
        # running from PyInstaller EXE
        return os.path.dirname(sys.executable)
    else:
        # running from source
        return os.path.dirname(os.path.abspath(__file__))

class LauncherWindow(QMainWindow):
    release_check_finished = pyqtSignal(object, object, object)
    
    def __init__(self, apps=None, repo=DEFAULT_GITHUB_REPO):
        super().__init__()
        self.setWindowTitle('Plategen Launcher')
        self.setMinimumSize(480, 260)
        self.repo = repo
        
        # Store COM object reference here if found active, otherwise None
        self.acad_com_ref = None 

        # apps: list of tuples (label, filename_without_path)
        here = get_app_dir()
        if apps is None:
            apps = [
                ('Battery Charger (BCH)', 'app_bch.py'),
                ('DB Rating Plate (DB)', 'app_db.py'),
                ('UPS Rating Plate (UPS)', 'app_ups.py'),
            ]
        self.apps = [(label, os.path.join(here, fname)) for label, fname in apps]

        # Persistent settings
        self.settings = QSettings('plategen', 'launcher')

        self._init_menu()
        self._init_ui()

        # statusbar: show version
        ver = read_local_version()
        self.statusBar().showMessage(f'Version: {ver}')

        self.release_check_finished.connect(self._on_release_check_finished)

    def _init_menu(self):
        menubar = self.menuBar()
        settings_menu = menubar.addMenu('Settings')

        self.auto_update_action = QAction('Auto-check updates on start', self, checkable=True)
        val = self.settings.value('auto_update', False, type=bool)
        self.auto_update_action.setChecked(val)
        self.auto_update_action.toggled.connect(self._on_auto_update_toggled)
        settings_menu.addAction(self.auto_update_action)

        # Auto-open AutoCAD on launch (persisted)
        self.auto_open_acad_action = QAction('Auto-open AutoCAD before launching apps', self, checkable=True)
        aopen = self.settings.value('auto_open_acad', False, type=bool)
        self.auto_open_acad_action.setChecked(aopen)
        self.auto_open_acad_action.toggled.connect(self._on_auto_open_acad_toggled)
        settings_menu.addAction(self.auto_open_acad_action)
        
        # Add a menu option to manually kill AutoCAD
        kill_acad_action = QAction('Kill AutoCAD Process(es)', self)
        kill_acad_action.triggered.connect(self.kill_autocad_process)
        settings_menu.addAction(kill_acad_action)

        help_menu = menubar.addMenu('Help')
        chk_action = QAction('Check for update (manual)', self)
        chk_action.triggered.connect(self.check_for_update)
        help_menu.addAction(chk_action)

        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def _init_ui(self):
        w = QWidget()
        self.setCentralWidget(w)
        layout = QVBoxLayout(w)
        layout.setSpacing(12)

        info = QLabel('Select an app and click Launch')
        info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info)

        box = QGroupBox('Available Apps')
        gl = QGridLayout()
        box.setLayout(gl)

        self.launch_buttons = []
        for i, (label, path) in enumerate(self.apps):
            btn = QPushButton(label)
            btn.setMinimumHeight(64)
            btn.setMinimumWidth(180)
            btn.setStyleSheet('font-size:14px; text-align:left; padding:8px;')
            # attempt to load an icon next to the app file: same basename .png or .ico
            icon = self._find_icon_for(path)
            if icon:
                btn.setIcon(icon)
                try:
                    from PyQt6.QtCore import QSize
                    btn.setIconSize(QSize(48, 48))
                except Exception:
                    pass
            btn.clicked.connect(self._make_launcher(path))
            gl.addWidget(btn, i // 2, i % 2)
            self.launch_buttons.append(btn)

        layout.addWidget(box)

        # --- Bottom layout with Launch/Kill/Update buttons ---
        bottom = QHBoxLayout()
        
        # Launch AutoCAD button 
        self.launch_acad_btn = QPushButton('Launch AutoCAD')
        self.launch_acad_btn.setToolTip('Start a new instance of AutoCAD.')
        # Connect to the threaded wrapper to prevent GUI freeze
        self.launch_acad_btn.clicked.connect(self._launch_autocad_threaded)
        self.launch_acad_btn.setEnabled(False) # Disabled initially
        
        bottom.addWidget(self.launch_acad_btn)
        bottom.addSpacing(5)


        # Kill AutoCAD button
        self.kill_acad_btn = QPushButton('Kill AutoCAD')
        self.kill_acad_btn.setToolTip('Forcefully close all running AutoCAD instances.')
        self.kill_acad_btn.clicked.connect(self.kill_autocad_process)
        self.kill_acad_btn.setEnabled(False) # Disabled initially
        
        bottom.addWidget(self.kill_acad_btn)
        
        bottom.addStretch() # Push Launch/Kill buttons to the left

        # Update button
        self.update_btn = QPushButton('Check Update')
        self.update_btn.clicked.connect(self.check_for_update)
        bottom.addWidget(self.update_btn)

        # AutoCAD status label in status bar
        self.acad_status_label = QLabel('AutoCAD: Unknown')
        self.statusBar().addPermanentWidget(self.acad_status_label)
        
        layout.addLayout(bottom)

        # if auto-check enabled on start, trigger background check
        if self.auto_update_action.isChecked():
            QTimerThread(self._start_release_check).start()

        # Start periodic AutoCAD status checks (every 3s)
        self.acad_timer = QTimer(self)
        self.acad_timer.timeout.connect(self.update_autocad_status)
        self.acad_timer.start(3000)

        # initial status
        self.update_autocad_status()

    def _make_launcher(self, path):
        def _launch():
            try:
                # Check for .exe (PyInstaller mode)
                exe_path = os.path.splitext(path)[0] + '.exe'
                if os.path.exists(exe_path):
                    subprocess.Popen([exe_path], cwd=os.path.dirname(exe_path))
                    return

                # In development mode we can launch .py
                if os.path.exists(path) and path.lower().endswith('.py'):
                    subprocess.Popen([sys.executable, path], cwd=os.path.dirname(path))
                    return

                QMessageBox.warning(
                    self,
                    'Launch failed',
                    f'Could not find application: {os.path.basename(path)}'
                )

            except Exception as e:
                QMessageBox.critical(
                    self, 'Error',
                    f'Failed to launch: {e}'
                )
        return _launch


    def _on_auto_update_toggled(self, checked):
        self.settings.setValue('auto_update', bool(checked))

    def _on_auto_open_acad_toggled(self, checked):
        self.settings.setValue('auto_open_acad', bool(checked))

    def _find_icon_for(self, path):
        # look for same basename with .png or .ico in same dir or icons/ subdir
        try:
            base = os.path.splitext(os.path.basename(path))[0]
            d = os.path.dirname(path)
            candidates = [os.path.join(d, base + ext) for ext in ('.png', '.ico')]
            candidates += [os.path.join(d, 'icons', base + ext) for ext in ('.png', '.ico')]
            for c in candidates:
                if os.path.exists(c):
                    return QIcon(c)
        except Exception:
            pass
        return None
    
    # --- Threaded AutoCAD Launch Handler ---

    def _launch_autocad_threaded(self):
        """Starts the launch and wait process in a separate thread to prevent GUI freeze."""
        # Temporarily show status and disable the button while launching
        self.statusBar().showMessage('Launching AutoCAD (do not close launcher)...')
        self.launch_acad_btn.setEnabled(False)
        
        # Run the launch logic in the background
        threading.Thread(target=self._run_autocad_launch_logic, daemon=True).start()

    def _run_autocad_launch_logic(self):
        """Worker function running in a separate thread."""
        try:
            # This call contains the blocking loop if wait=True
            self.launch_autocad(wait=True)
        finally:
            # Ensure UI status and button states are updated on the main thread after completion
            QTimer.singleShot(0, self.update_autocad_status)


    # --- AutoCAD Status and Control Methods ---

    def _get_autocad_pids(self):
        """Checks the Windows process list and returns a list of (name, PID) tuples."""
        target_processes = ('acad.exe', 'accoreconsole.exe', 'acadlt.exe', 'accore.exe')
        running_pids = []
        
        try:
            # Use tasklist to check for common AutoCAD executable names
            out = subprocess.check_output(
                ['tasklist', '/NH', '/FO', 'CSV'], 
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            text = out.decode('utf-8', errors='ignore')
            
            for line in text.splitlines():
                parts = [p.strip(' "') for p in line.split(',')]
                if len(parts) >= 2 and parts[0].lower() in target_processes:
                    try:
                        pid = int(parts[1])
                        running_pids.append((parts[0], pid))
                    except ValueError:
                        continue 
            
        except Exception:
            # Tasklist failed or not available
            pass
            
        return running_pids

    def check_autocad_running(self):
        """
        Checks if AutoCAD is running via COM (preferred) or via OS process list (fallback).
        If a COM AutoCAD instance is found, stores the COM object in self.acad_com_ref.
        Returns True if AutoCAD is running (COM or process detected), otherwise False.
        """
        # 1) If we already have a stored COM reference, verify it's still usable.
        if getattr(self, 'acad_com_ref', None) is not None:
            try:
                # Access a harmless property to confirm proxy validity
                _ = getattr(self.acad_com_ref, 'Version', None)
                return True
            except Exception:
                # Stale reference — drop it and continue detection
                try:
                    del self.acad_com_ref
                except Exception:
                    pass
                self.acad_com_ref = None

        # 2) Try COM GetActiveObject for multiple known ProgIDs (Electrical, Civil, LT, versioned)
        if COM_AVAILABLE:
            # List of ProgIDs to try (common ones + versioned variants)
            progids = [
                'AutoCAD.Application',
                'AutoCAD.Application.26',  # 2026
                'AutoCAD.Application.25',
                'AutoCAD.Application.24',
                'AutoCAD.Application.23',
                'AutoCAD.Application.22',
                'AutoCAD.Application.21',
                'AutoCADLT.Application',
                'AutoCADElectrical.Application',
                'AutoCAD.Electrical.Application',
                'AutoCADCivil3D.Application'
            ]
            try:
                # Initialize COM on this thread before calling GetActiveObject
                pythoncom.CoInitialize()
                for pid in progids:
                    try:
                        # GetActiveObject will raise if that ProgID is not active
                        obj = win32com.client.GetActiveObject(pid)
                        if obj:
                            # store the COM reference for later use
                            self.acad_com_ref = obj
                            # touch a property to ensure proxy is alive
                            try:
                                _ = getattr(self.acad_com_ref, 'Version', None)
                            except Exception:
                                # if proxy fails, unset and continue searching
                                try:
                                    del self.acad_com_ref
                                except Exception:
                                    pass
                                self.acad_com_ref = None
                                continue
                            return True
                    except Exception:
                        # not found — continue trying other ProgIDs
                        continue
            finally:
                # Only uninitialize if we didn't retain a COM reference.
                # If we did store acad_com_ref, leaving the COM apartment active here is OK —
                # caller code should be careful when calling CoInitialize/CoUninitialize in threads.
                if getattr(self, 'acad_com_ref', None) is None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        # 3) Fallback: check OS process list for known AutoCAD executables
        pids = self._get_autocad_pids()
        if pids:
            return True

        return False



    def update_autocad_status(self):
        running = self.check_autocad_running()
        txt = 'AutoCAD: Running' if running else 'AutoCAD: Not running'
        style = "color: green;" if running else "color: red;"
        try:
            # Update status bar text and style
            self.acad_status_label.setText(txt)
            self.acad_status_label.setStyleSheet(style)
            
            # Control button enabled state
            self.kill_acad_btn.setEnabled(running)
            
            # Only enable launch button if not running 
            self.launch_acad_btn.setEnabled(not running)

            # Reset status message if it was a temporary "waiting" message
            if not self.statusBar().currentMessage().startswith('Version:'):
                 self.statusBar().showMessage(f'Version: {read_local_version()}')

        except Exception:
            pass
            
    def kill_autocad_process(self):
        """
        Attempts to gracefully quit AutoCAD via COM, otherwise force-terminates 
        the processes detected by PID.
        """
        if not self.check_autocad_running():
            QMessageBox.information(self, 'AutoCAD Status', 'AutoCAD is not currently running.')
            return

        # Confirmation dialog before killing
        reply = QMessageBox.question(self, 'Confirm Kill',
            "Are you sure you want to forcibly close all AutoCAD instances? Any unsaved work will be lost.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
        if reply == QMessageBox.StandardButton.No:
            return

        # 1. Attempt graceful close via COM (.Quit())
        if self.acad_com_ref is not None and COM_AVAILABLE:
            try:
                # Re-initialize COM context for the call
                pythoncom.CoInitialize() 
                self.statusBar().showMessage("Attempting graceful AutoCAD close (COM.Quit())...")
                self.acad_com_ref.Quit()
                
                # Release the COM reference immediately
                del self.acad_com_ref
                self.acad_com_ref = None
                pythoncom.CoUninitialize()
                
                # Give it a moment to shut down
                time.sleep(1) 

                if not self.check_autocad_running():
                    QMessageBox.information(self, 'Success', 'AutoCAD closed successfully via COM.')
                    self.statusBar().showMessage("AutoCAD closed successfully.")
                    return # Successfully closed, exit

            except Exception as e:
                self.statusBar().showMessage(f"COM Quit failed. Falling back to process kill: {e}")
            finally:
                 # Cleanup attempt if COM object didn't quit cleanly
                 if self.acad_com_ref is not None:
                     try:
                        del self.acad_com_ref
                        self.acad_com_ref = None
                        if COM_AVAILABLE:
                            pythoncom.CoUninitialize()
                     except Exception:
                        pass


        # 2. Forceful termination via Process ID (PID)
        pids = self._get_autocad_pids()
        if pids:
            if psutil is None:
                QMessageBox.critical(self, 'Error', 'psutil library is required for forceful termination but not found.')
                self.statusBar().showMessage("Termination failed (psutil missing).")
                return

            self.statusBar().showMessage(f"Forcefully terminating {len(pids)} AutoCAD process(es)...")
            success_count = 0
            
            for name, pid in pids:
                try:
                    process = psutil.Process(pid)
                    process.terminate() # Send SIGTERM/terminate signal
                    process.wait(timeout=5)
                    
                    if process.is_running():
                        process.kill() # Forceful kill
                        
                    success_count += 1
                except psutil.NoSuchProcess:
                    success_count += 1 # Process died between check and kill attempt
                    continue
                except Exception as e:
                    print(f"Could not terminate PID {pid}: {e}")
            
            self.update_autocad_status()
            if success_count > 0:
                QMessageBox.information(self, 'Success', f'Successfully terminated {success_count} AutoCAD process(es).')
                self.statusBar().showMessage("AutoCAD closed.")
            else:
                QMessageBox.warning(self, 'Failure', 'Failed to terminate any AutoCAD processes.')
                self.statusBar().showMessage("Termination failed.")
        else:
            self.statusBar().showMessage("AutoCAD not found after kill attempt.")
            QMessageBox.information(self, 'Status', 'AutoCAD process was not found.')


    def launch_autocad(self, wait=False, timeout=20.0):
        """
        Start AutoCAD by best-effort strategies:
        - try exe names in PATH
        - try Dispatch()ing known ProgIDs (which can start COM server)
        - search common Program Files\Autodesk folders for executables and launch
        If wait==True, poll until COM becomes available (or timeout). When COM appears,
        force Visible=True and maximize the window (WindowState=3).
        Returns True on successful start/detection, False otherwise.
        """
        # If AutoCAD already detected, we're done
        try:
            if self.check_autocad_running():
                QTimer.singleShot(0, lambda: self.statusBar().showMessage('AutoCAD is already running.'))
                return True
        except Exception:
            pass

        started = False

        # 1) Try launching well-known executable names (acad.exe, accoreconsole.exe, acadlt.exe)
        exe_names = ['acad.exe', 'accoreconsole.exe', 'acadlt.exe', 'accore.exe']
        for exe in exe_names:
            try:
                subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
                started = True
                break
            except Exception:
                started = False

        # 2) If not started yet, try Dispatch of known ProgIDs (this can start AutoCAD COM server)
        if not started and COM_AVAILABLE:
            progids = [
                'AutoCAD.Application.26',
                'AutoCAD.Application.25',
                'AutoCAD.Application.24',
                'AutoCAD.Application.23',
                'AutoCAD.Application.22',
                'AutoCAD.Application.21',
                'AutoCADElectrical.Application',
                'AutoCAD.Electrical.Application',
                'AutoCADCivil3D.Application',
                'AutoCADLT.Application',
                'AutoCAD.Application'
            ]
            for pid in progids:
                try:
                    # Dispatch may start the COM server. It sometimes returns an object even if GUI not visible.
                    obj = win32com.client.Dispatch(pid)
                    if obj:
                        # Try to make GUI visible and maximize ASAP
                        try:
                            obj.Visible = True
                        except Exception:
                            pass
                        try:
                            obj.WindowState = 3  # SW_SHOWMAXIMIZED
                        except Exception:
                            pass
                        # store reference for later actions (e.g. Quit())
                        self.acad_com_ref = obj
                        started = True
                        break
                except Exception:
                    continue

        # 3) If still not started, search Program Files\Autodesk for executables and launch one
        if not started:
            search_roots = []
            pf = os.environ.get('ProgramFiles')
            pfx = os.environ.get('ProgramFiles(x86)')
            if pf:
                search_roots.append(pf)
            if pfx and pfx != pf:
                search_roots.append(pfx)
            # also include explicit common candidates
            search_roots += [r'C:\Program Files', r'C:\Program Files (x86)']

            found_path = None
            for root in search_roots:
                try:
                    # Look for top-level folders that mention Autodesk to focus the search
                    for folder in os.listdir(root):
                        if 'autodesk' in folder.lower() or 'autocad' in folder.lower():
                            d = os.path.join(root, folder)
                            if os.path.isdir(d):
                                for dirpath, dirnames, filenames in os.walk(d):
                                    for name in ('acad.exe', 'accoreconsole.exe', 'acadlt.exe', 'accore.exe'):
                                        if name in filenames:
                                            found_path = os.path.join(dirpath, name)
                                            break
                                    if found_path:
                                        break
                        if found_path:
                            break
                    if found_path:
                        break
                except Exception:
                    continue

            if found_path:
                try:
                    subprocess.Popen([found_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
                    started = True
                except Exception:
                    started = False

        if not started:
            QTimer.singleShot(0, lambda: QMessageBox.warning(self, 'Launch AutoCAD', 'Could not start or find AutoCAD. Ensure AutoCAD is installed.'))
            return False

        # If caller asked to wait, poll for COM availability up to timeout
        if wait and COM_AVAILABLE:
            QTimer.singleShot(0, lambda: self.statusBar().showMessage('Waiting for AutoCAD to become available...'))
            end = time.time() + timeout
            while time.time() < end:
                # check_autocad_running will try to find and populate self.acad_com_ref
                if self.check_autocad_running():
                    # If we have a COM reference, force UI visible and maximized
                    if getattr(self, 'acad_com_ref', None) is not None:
                        try:
                            self.acad_com_ref.Visible = True
                        except Exception:
                            pass
                        try:
                            self.acad_com_ref.WindowState = 3
                        except Exception:
                            pass

                    # Update UI on main thread and return success
                    QTimer.singleShot(0, lambda: self.statusBar().showMessage('AutoCAD started'))
                    QTimer.singleShot(0, self.update_autocad_status)
                    return True

                time.sleep(0.6)

            # timed out
            QTimer.singleShot(0, lambda: QMessageBox.warning(self, 'Launch AutoCAD', 'AutoCAD did not become available within timeout.'))
            return False

        # If not waiting for COM, we consider launch successful (executable started or Dispatch succeeded)
        QTimer.singleShot(0, lambda: self.statusBar().showMessage('AutoCAD start attempted'))
        QTimer.singleShot(0, self.update_autocad_status)
        return True

    def check_for_update(self):
        self.statusBar().showMessage('Checking latest release...')
        self.update_btn.setEnabled(False)
        threading.Thread(target=self._start_release_check, daemon=True).start()

    def _start_release_check(self):
        tag, url, err = fetch_latest_github_release(self.repo)
        self.release_check_finished.emit(tag, url, err)

    def _on_release_check_finished(self, tag, url, err):
        self.update_btn.setEnabled(True)
        if err:
            self.statusBar().showMessage(f'Update check failed: {err}')
            QMessageBox.critical(self, 'Update check failed', f'Error: {err}')
            return
        cur = read_local_version()
        self.statusBar().showMessage(f'Latest: {tag} — Current: {cur}')
        if tag and cur and tag.strip() != cur.strip():
            rv = QMessageBox(self)
            rv.setWindowTitle('Update Available')
            rv.setText(f'Latest release: {tag}\nYou have: {cur}')
            open_btn = rv.addButton('Open Release', QMessageBox.ButtonRole.AcceptRole)
            rv.addButton('Dismiss', QMessageBox.ButtonRole.RejectRole)
            rv.exec()
            if rv.clickedButton() == open_btn and url:
                webbrowser.open(url)
        else:
            QMessageBox.information(self, 'Up-to-date', f'You are up-to-date. Latest: {tag} (You: {cur})')

    def show_about(self):
        cur = read_local_version()
        # This part should probably be moved to a separate thread like check_for_update,
        # but maintaining the user's original immediate fetch for 'About'.
        tag, url, err = fetch_latest_github_release(self.repo, timeout=6) 
        body = f"Plategen Launcher\nVersion: {cur}\nRepo: {self.repo}\n"
        if tag:
            body += f"Latest release: {tag}\n{url}\n"
        if err:
            body += f"(Latest lookup failed: {err})\n"
        QMessageBox.information(self, 'About Plategen', body)


class QTimerThread(threading.Thread):
    """Utility thread for delayed UI-safe start (very small helper)."""
    def __init__(self, fn, delay=0.1):
        super().__init__(daemon=True)
        self.fn = fn
        self.delay = delay

    def run(self):
        time.sleep(self.delay)
        try:
            self.fn()
        except Exception:
            pass


def main():
    set_windows_app_id()
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Ensure app icon exists and set it for the application and main window
    try:
        icon_path = ensure_app_icon()
        if icon_path:
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass

    # Try to read a repository override from env var (optional)
    repo = os.environ.get('PLATGEN_GITHUB_REPO', DEFAULT_GITHUB_REPO)

    w = LauncherWindow(repo=repo)
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    # This requires psutil, pywin32 (for AutoCAD), and PyQt6 to run.
    main()