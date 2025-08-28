"""
This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
Fixed version addressing common WPS Office add-in loading issues and PyInstaller compatibility.
This is different to addin_client2 because it has been added come enhanced loggings, but not with persistent log file.
"""
import datetime
import traceback
import os
import sys
import threading
import requests
import win32com.client
import win32api
import logging
import winreg
import pythoncom
from tkinter import simpledialog, Tk

# Enhanced logging setup with error handling and permissions bypass
def setup_logging():
    """Setup robust logging with fallback mechanisms"""
    try:
        # Base directory (adjust path if needed)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(base_dir, "logs")
        
        # Create logs directory with full permissions
        try:
            os.makedirs(log_dir, exist_ok=True)
            # Try to set permissions (Windows)
            os.chmod(log_dir, 0o777)
        except PermissionError:
            print(f"Warning: Could not set full permissions on log directory: {log_dir}")
        except Exception as e:
            print(f"Warning: Error setting up log directory: {e}")
        
        # Log file path with timestamp to avoid conflicts
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"addin_debug_{timestamp}.log"
        log_file = os.path.join(log_dir, log_filename)
        
        # Test write access before setting up logging
        try:
            with open(log_file, 'a') as test_file:
                test_file.write("# Log file created\n")
        except PermissionError:
            # If permission denied, try alternative locations
            fallback_locations = [
                os.path.join(os.getenv('TEMP', ''), 'WPS_AI_Logs'),
                os.path.join(os.getenv('USERPROFILE', ''), 'Desktop', 'WPS_AI_Logs'),
                os.path.join(os.getcwd(), 'temp_logs')
            ]
            
            for fallback_dir in fallback_locations:
                try:
                    os.makedirs(fallback_dir, exist_ok=True)
                    test_log_file = os.path.join(fallback_dir, log_filename)
                    with open(test_log_file, 'a') as test_file:
                        test_file.write("# Fallback log file created\n")
                    log_file = test_log_file
                    print(f"Using fallback log location: {log_file}")
                    break
                except Exception:
                    continue
            else:
                # If all fallback locations fail, use console only
                print("WARNING: Could not create log file. Using console logging only.")
                log_file = None
        
        # Configure logging with multiple handlers
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)  # Capture all levels
        
        # Clear any existing handlers
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # Create formatter
        formatter = logging.Formatter(
            '[%(asctime)s] - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Add file handler if we have a valid log file
        if log_file:
            try:
                file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
                print(f"File logging enabled: {log_file}")
            except Exception as e:
                print(f"Could not create file handler: {e}")
        
        # Always add console handler as backup
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        return log_file
        
    except Exception as e:
        print(f"CRITICAL: Logging setup completely failed: {e}")
        # Create minimal console-only logging
        logging.basicConfig(
            level=logging.INFO,
            format='[%(asctime)s] - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        return None

# Initialize logging
current_log_file = setup_logging()

# Enhanced logging function with error handling
def log_message(message, level="info"):
    """Enhanced logging function with fallback mechanisms"""
    try:
        logger = logging.getLogger()
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Log at specified level
        if level.lower() == "debug":
            logger.debug(message)
        elif level.lower() == "warning":
            logger.warning(message)
        elif level.lower() == "error":
            logger.error(message)
        elif level.lower() == "critical":
            logger.critical(message)
        else:
            logger.info(message)
        
        # Force flush to ensure message is written
        for handler in logger.handlers:
            if hasattr(handler, 'flush'):
                handler.flush()
                
    except Exception as e:
        # Last resort: direct print with timestamp
        print(f"[{timestamp}] LOGGING ERROR: {e}")
        print(f"[{timestamp}] ORIGINAL MESSAGE: {message}")
        
        # Try to write directly to log file if it exists
        if current_log_file and os.path.exists(current_log_file):
            try:
                with open(current_log_file, 'a', encoding='utf-8') as f:
                    f.write(f"[{timestamp}] - {level.upper()} - DIRECT_WRITE - {message}\n")
                    f.flush()
            except Exception:
                pass

# Initial log messages
log_message("=== SCRIPT EXECUTION STARTED ===")
log_message(f"Python Version: {sys.version}")
log_message(f"Executable Path: {sys.executable}")
log_message(f"Command Line Arguments: {sys.argv}")
log_message(f"Current Working Directory: {os.getcwd()}")
log_message(f"Log file location: {current_log_file}")

# Configuration - BACKEND IP address
BACKEND_URL = "http://127.0.0.1:8000"

# Consistent naming
WPS_ADDIN_ENTRY_NAME = "WPSAIAddin.Connect"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller bundling """
    try:
        base_path = sys._MEIPASS
        log_message(f"Using PyInstaller bundle path: {base_path}", "debug")
    except Exception:
        base_path = os.path.abspath(".")
        log_message(f"Using development path: {base_path}", "debug")
    
    full_path = os.path.join(base_path, relative_path)
    log_message(f"Resource path for '{relative_path}': {full_path}", "debug")
    return full_path

def get_wps_application():
    """Gets the running WPS Writer Application object."""
    try:
        log_message("Attempting to get WPS Application object", "debug")
        app = win32com.client.GetActiveObject("kwps.Application")
        log_message("Successfully obtained WPS Application object")
        return app
    except Exception as e:
        log_message(f"Error getting WPS Application object: {e}", "error")
        log_message(f"Exception traceback: {traceback.format_exc()}", "debug")
        return None

def insert_text_at_cursor(text):
    """Inserts text into the active document at the current cursor position."""
    log_message(f"Attempting to insert text: {text[:50]}...", "debug")
    wps_app = get_wps_application()
    if wps_app and wps_app.Documents.Count > 0:
        try:
            wps_app.Selection.TypeText(Text=text)
            log_message("Text successfully inserted into active WPS document.")
        except Exception as e:
            log_message(f"Error inserting text into WPS document: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
    else:
        log_message("Warning: Could not find an active WPS document to insert text into.", "warning")

class WPSAddin:
    _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}" # python -c "import uuid; print(uuid.uuid4())"
    _reg_desc_ = "AI Office Automation"
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
    _reg_class_spec_ = __name__ + ".WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
        'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter'
        'GetCustomUI' # This is the method that WPS Office expects to load the XML
    ]
    # Attribute used by Microosoft office but not WPS Office
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("=== Add-in __init__ started ===")
        try:
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Attempting to load ribbon from: {ribbon_path}")
            
            if not os.path.exists(ribbon_path):
                log_message(f"FATAL: Ribbon XML file does NOT exist at the path.", "critical")
                raise FileNotFoundError(f"Ribbon XML not found at {ribbon_path}")
            
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                self.ribbon = f.read()
            log_message(f"Ribbon XML loaded successfully. Length: {len(self.ribbon)} characters")
            
    

            self.translations = {
                1033: {
                    "tab": "AI Assistant", "group": "AI Tools", "run_prompt": "Run General Prompt",
                    "analyze_doc": "Analyze Document", "summarize_doc": "Summarize Document",
                    "create_memo": "Create Memo", "create_minutes": "Create Minutes", "create_cover_letter": "Create Cover Letter",
                    "prompt_title": "AI Assistant", "prompt_message": "Enter your request (e.g., 'write a report on X'):",
                    "memo_topic": "Enter the memo topic:", "memo_audience": "Enter the memo's audience:",
                    "minutes_topic": "Enter the meeting topic:", "minutes_attendees": "Enter attendees (comma-separated):",
                    "minutes_info": "Enter key discussion points:",
                    "cover_letter_topic": "Enter the job position:", "cover_letter_audience": "Enter the hiring manager/company:",
                    "action_cancelled": "Action cancelled.", "contacting_server": "AI Assistant: Contacting server, please wait...",
                    "connection_error": "\n\nERROR: Could not connect to the backend server. Please ensure the AI Backend is running.\n\n",
                    "unexpected_error": "\n\nAn unexpected error occurred: {e}\n\n",
                    "result_header": "\n\n--- AI Assistant Result ---\n", "result_footer": "\n--- End of Result ---\n\n",
                    "no_active_doc": "No active document found."
                }
            }
            log_message("=== Add-in __init__ completed successfully ===")
        except FileNotFoundError as e:
            log_message(f"FATAL ERROR: Ribbon XML file not found at {ribbon_path}", "critical")
            log_message(f"FileNotFoundError details: {e}", "error")
            self.ribbon = ""
        except Exception as e:
            log_message(f"FATAL ERROR IN __init__: {e}", "critical")
            log_message(f"Full traceback: {traceback.format_exc()}", "error")
            self.ribbon = ""
            
    
    # Explicit Ribbon Loader for WPS
    def GetCustomUI(self, ribbon_id):
        """
        WPS will call this to request the ribbon XML.
        """
        log_message(f"GetCustomUI called with ribbon_id: {ribbon_id}")
        try:
            ribbon_path = resource_path('ribbon.xml')
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            log_message(f"Failed to load ribbon.xml: {e}", "critical")
            return None

    def _get_localized_string(self, key):
        lang_id = 1033
        wps_app = get_wps_application()
        if wps_app:
            try:
                lang_id = wps_app.LanguageSettings.LanguageID(1)
                log_message(f"Detected WPS language ID: {lang_id}", "debug")
            except Exception as e:
                log_message(f"Could not get WPS language ID: {e}", "debug")
        return self.translations.get(lang_id, self.translations[1033]).get(key, key)

    def GetTabLabel(self, c): 
        result = self._get_localized_string("tab")
        log_message(f"GetTabLabel called, returning: {result}", "debug")
        return result
        
    def GetGroupLabel(self, c): 
        result = self._get_localized_string("group")
        log_message(f"GetGroupLabel called, returning: {result}", "debug")
        return result
        
    def GetRunPromptLabel(self, c): 
        result = self._get_localized_string("run_prompt")
        log_message(f"GetRunPromptLabel called, returning: {result}", "debug")
        return result
        
    def GetAnalyzeDocLabel(self, c): 
        result = self._get_localized_string("analyze_doc")
        log_message(f"GetAnalyzeDocLabel called, returning: {result}", "debug")
        return result
        
    def GetSummarizeDocLabel(self, c): 
        result = self._get_localized_string("summarize_doc")
        log_message(f"GetSummarizeDocLabel called, returning: {result}", "debug")
        return result
        
    def GetCreateMemoLabel(self, c): 
        result = self._get_localized_string("create_memo")
        log_message(f"GetCreateMemoLabel called, returning: {result}", "debug")
        return result
        
    def GetCreateMinutesLabel(self, c): 
        result = self._get_localized_string("create_minutes")
        log_message(f"GetCreateMinutesLabel called, returning: {result}", "debug")
        return result
        
    def GetCreateCoverLetterLabel(self, c): 
        result = self._get_localized_string("create_cover_letter")
        log_message(f"GetCreateCoverLetterLabel called, returning: {result}", "debug")
        return result

    def OnLoadImage(self, imageName):
        image_path = resource_path(f"{imageName}.png")
        log_message(f"Attempting to load image: {image_path}")
        try:
            if not os.path.exists(image_path):
                log_message(f"Image file does not exist: {image_path}", "warning")
                return None
                
            img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
            log_message(f"Successfully loaded image '{imageName}'.")
            return img_handle
        except Exception as e:
            log_message(f"ERROR: Failed to load image '{imageName}': {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return None

    def _call_backend_task(self, endpoint: str, payload: dict):
        log_message(f"=== Starting backend call to: {endpoint} ===")
        log_message(f"Payload: {payload}", "debug")
        try:
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
            log_message(f"Making POST request to: {BACKEND_URL}{endpoint}")
            
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
            log_message(f"Received response status: {response.status_code}")
            
            response.raise_for_status()
            result_data = response.json()
            log_message(f"Response data: {result_data}", "debug")
            
            result = result_data.get("result", "")
            header = self._get_localized_string("result_header")
            footer = self._get_localized_string("result_footer")
            insert_text_at_cursor(f"{header}{result}{footer}")
            log_message(f"Successfully received and inserted response from {endpoint}.")
        except requests.exceptions.ConnectionError as e:
            log_message(f"Connection error to {endpoint}: {e}", "error")
            insert_text_at_cursor(self._get_localized_string("connection_error"))
        except requests.exceptions.Timeout as e:
            log_message(f"Timeout error for {endpoint}: {e}", "error")
            insert_text_at_cursor("\n\nERROR: Request timed out. Please try again.\n\n")
        except Exception as e:
            log_message(f"Error calling {endpoint}: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))

    def OnRunPrompt(self, c):
        log_message("=== OnRunPrompt called ===")
        try:
            root = Tk(); root.withdraw()
            prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("prompt_message"))
            root.destroy()
            
            if not prompt:
                log_message("User cancelled prompt input")
                return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
            
            log_message(f"User entered prompt: {prompt}")
            threading.Thread(target=self._call_backend_task, 
                            args=("/process", {"prompt": prompt})).start()
        except Exception as e:
            log_message(f"Error in OnRunPrompt: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnAnalyzeDocument(self, c):
        log_message("=== OnAnalyzeDocument called ===")
        try:
            wps_app = get_wps_application()
            if not wps_app or wps_app.Documents.Count == 0: 
                log_message("No active WPS document found", "warning")
                return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
            
            content = wps_app.ActiveDocument.Content.Text
            log_message(f"Document content length: {len(content)} characters")
            threading.Thread(target=self._call_backend_task, 
                            args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()
        except Exception as e:
            log_message(f"Error in OnAnalyzeDocument: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnSummarizeDocument(self, c):
        log_message("=== OnSummarizeDocument called ===")
        try:
            wps_app = get_wps_application()
            if not wps_app or wps_app.Documents.Count == 0: 
                log_message("No active WPS document found", "warning")
                return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
            
            content = wps_app.ActiveDocument.Content.Text
            log_message(f"Document content length: {len(content)} characters")
            threading.Thread(target=self._call_backend_task, 
                            args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()
        except Exception as e:
            log_message(f"Error in OnSummarizeDocument: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateMemo(self, c):
        log_message("=== OnCreateMemo called ===")
        try:
            root = Tk(); root.withdraw()
            topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("memo_topic"))
            if not topic:
                root.destroy()
                log_message("User cancelled memo topic input")
                return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
            
            audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("memo_audience"))
            root.destroy()
            
            payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
            log_message(f"Memo payload: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()
        except Exception as e:
            log_message(f"Error in OnCreateMemo: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateMinutes(self, c):
        log_message("=== OnCreateMinutes called ===")
        try:
            root = Tk(); root.withdraw()
            topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("minutes_topic"))
            if not topic:
                root.destroy()
                log_message("User cancelled minutes topic input")
                return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
            
            attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                                self._get_localized_string("minutes_attendees"))
            info = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("minutes_info"))
            root.destroy()
            
            payload = {
                "doc_type": "minutes", "topic": topic, "audience": "Meeting Attendees",
                "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()],
                "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]
            }
            log_message(f"Minutes payload: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()
        except Exception as e:
            log_message(f"Error in OnCreateMinutes: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateCoverLetter(self, c):
        log_message("=== OnCreateCoverLetter called ===")
        try:
            root = Tk(); root.withdraw()
            topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("cover_letter_topic"))
            if not topic:
                root.destroy()
                log_message("User cancelled cover letter topic input")
                return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
            
            audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("cover_letter_audience"))
            root.destroy()
            
            payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
            log_message(f"Cover letter payload: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()
        except Exception as e:
            log_message(f"Error in OnCreateCoverLetter: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")


if __name__ == '__main__':
    
    log_message("=== Executing main block ===")

    def is_pyinstaller_bundle():
        """Check if running as PyInstaller bundle"""
        result = getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')
        log_message(f"Is PyInstaller bundle: {result}", "debug")
        return result
    
    def register_com_server_pyinstaller(cls):
        """Register COM server for PyInstaller executable"""
        log_message("=== Starting PyInstaller COM server registration ===")
        import winreg
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_
        
        # Get the executable path
        exe_path = sys.executable if is_pyinstaller_bundle() else __file__
        
        log_message(f"Registering PyInstaller COM server with executable: {exe_path}")
        
        try:
            # Register main CLSID
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                log_message(f"Created CLSID entry: {clsid}")
            
            # Register LocalServer32 (not InprocServer32 for exe)
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{exe_path}"')
                log_message(f"Created LocalServer32 entry")
            
            # Register ProgID
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, progid)
                log_message(f"Created ProgID entry: {progid}")
            
            # Register ProgID mapping
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                with winreg.CreateKeyEx(key, "CLSID") as clsid_key:
                    winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
                log_message(f"Created ProgID mapping")
            
            log_message("PyInstaller COM server registered successfully")
            return True
        except Exception as e:
            log_message(f"Failed to register PyInstaller COM server: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return False
    
    def register_com_server_python(cls):
        """Register COM server for regular Python execution"""
        log_message("=== Starting Python COM server registration ===")
        try:
            import win32com.server.register
            win32com.server.register.UseCommandLine(cls)
            log_message("Python COM server registered successfully")
            return True
        except Exception as e:
            log_message(f"Failed to register Python COM server: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return False
    
    def register_wps_addin_entry(clsid, progid, description):
        """
        Creates the specific registry entry that WPS Office looks for.
        """
        log_message("=== Starting WPS add-in entry registration ===")
        import winreg
        log_message(f"Attempting to create WPS-specific entry for ProgID: {progid}")
        
        # Try multiple possible WPS registry paths
        wps_addin_paths = [
            r"Software\Kingsoft\Office\Addins",
            r"Software\WPS\Office\Addins", 
            r"Software\WPS Office\Addins"
        ]
        
        registration_succeeded = False
        
        for base_path in wps_addin_paths:
            try:
                log_message(f"Trying registry path: HKCU\\{base_path}")
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_CREATE_SUB_KEY) as parent_key:
                    log_message(f"Successfully opened/created parent key at HKCU\\{base_path}")
                    with winreg.CreateKeyEx(parent_key, progid) as key:
                        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, description)
                        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                        winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                        winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
                        winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
                        log_message(f"Successfully created WPS Add-in entry '{progid}' at {base_path}")
                        registration_succeeded = True
                        break  # Success, no need to try other paths
            except Exception as e:
                log_message(f"Could not register at {base_path}: {e}", "warning")
                continue
        
        return registration_succeeded
    
    def register_server(cls):
        """Register COM server - handles both Python and PyInstaller"""
        log_message("=== Starting enhanced registration process ===")
        
        # Step 1: Register COM server (different methods for Python vs PyInstaller)
        if is_pyinstaller_bundle():
            com_success = register_com_server_pyinstaller(cls)
        else:
            com_success = register_com_server_python(cls)
            
        if not com_success:
            log_message("FATAL: COM server registration failed", "critical")
            print("FATAL: Failed to register the COM server. Please run as Administrator.")
            return False
            
        # Step 2: Register WPS Office add-in entries
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_
        
        if register_wps_addin_entry(clsid, progid, desc):
            log_message("WPS add-in registration successful")
            print("SUCCESS: Add-in registered successfully!")
            print("Please restart WPS Office to see the add-in.")
            return True
        else:
            log_message("WPS add-in registration failed", "error")
            print("FAILED: Could not register WPS Office add-in entry.")
            return False

    def unregister_server(cls):
        """Enhanced unregistration"""
        log_message("=== Starting unregistration process ===")
        import winreg
        
        # Unregister COM server
        if not is_pyinstaller_bundle():
            try:
                import win32com.server.register
                win32com.server.register.UnregisterServer(cls._reg_clsid_)
                log_message("Python COM server unregistered successfully")
            except Exception as e:
                log_message(f"Could not unregister Python COM server: {e}", "warning")

        # Remove WPS Office entries
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        
        # Remove COM server entries
        com_paths_to_remove = [
            f"CLSID\\{clsid}\\LocalServer32",
            f"CLSID\\{clsid}\\InprocServer32", 
            f"CLSID\\{clsid}\\ProgID",
            f"CLSID\\{clsid}",
            f"{progid}\\CLSID",
            progid
        ]
        
        for path in com_paths_to_remove:
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, path)
                log_message(f"Removed COM entry: {path}")
            except FileNotFoundError:
                log_message(f"COM entry not found (already removed): {path}", "debug")
            except Exception as e:
                log_message(f"Could not remove COM entry {path}: {e}", "warning")
        
        # Remove WPS Office entries
        wps_addin_paths = [
            f"Software\\Kingsoft\\Office\\Addins\\{progid}",
            f"Software\\WPS\\Office\\Addins\\{progid}",
            f"Software\\WPS Office\\Addins\\{progid}"
        ]
        
        for path in wps_addin_paths:
            try:
                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0)
                log_message(f"Removed: HKCU\\{path}")
            except FileNotFoundError:
                log_message(f"WPS entry not found (already removed): {path}", "debug")
            except Exception as e:
                log_message(f"Could not remove {path}: {e}", "warning")

        log_message("Unregistration process completed")
        print("Unregistration complete.")
    
    def run_com_server():
        """Run as COM server when called by Windows"""
        log_message("=== Starting COM server ===")
        try:
            import win32com.server.localserver
            log_message(f"Serving COM server with CLSID: {WPSAddin._reg_clsid_}")
            win32com.server.localserver.serve([WPSAddin._reg_clsid_])
        except Exception as e:
            log_message(f"Error running COM server: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
    
    def check_environment():
        """Check and setup environment"""
        log_message("=== Checking environment ===")
        
        if not is_pyinstaller_bundle():
            script_dir = os.path.dirname(os.path.abspath(__file__))
            if script_dir not in sys.path:
                sys.path.insert(0, script_dir)
                log_message(f"Added script directory to Python path: {script_dir}")
        
        log_message(f"Running as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        log_message(f"Executable: {sys.executable}")
        log_message(f"Current directory: {os.getcwd()}")
        log_message(f"Python path: {sys.path[:3]}...")  # Show first 3 paths
        
        # Check for required modules
        required_modules = ['win32com.client', 'win32api', 'winreg', 'pythoncom', 'requests', 'tkinter']
        for module in required_modules:
            try:
                __import__(module)
                log_message(f"Module {module}: OK", "debug")
            except ImportError as e:
                log_message(f"Module {module}: MISSING - {e}", "error")

    # Main command logic
    log_message("=== Starting main execution logic ===")
    check_environment()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        log_message(f"Command received: {command}")
        
        if command == '/regserver':
            log_message("Executing registration command")
            register_server(WPSAddin)
        elif command == '/unregserver':
            log_message("Executing unregistration command")
            unregister_server(WPSAddin)
        elif command == '/embedding':
            # This is called by Windows when WPS tries to instantiate the COM object
            log_message("Executing embedding command (COM server mode)")
            run_com_server()
        else:
            log_message(f"Unknown command: {command}", "warning")
            print(f"Unknown command: {command}")
    else:
        log_message("No command line arguments provided, showing usage")
        print("WPS Office AI Assistant Add-in")
        print("Usage:")
        print("  /regserver   - Register the add-in")
        print("  /unregserver - Unregister the add-in")
        print(f"Running as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        print(f"Log file: {current_log_file}")
        input("\nPress Enter to exit...")
    
    log_message("=== Script execution completed ===")
    
    # Final log flush to ensure all messages are written
    try:
        for handler in logging.getLogger().handlers:
            if hasattr(handler, 'flush'):
                handler.flush()
    except Exception:
        pass