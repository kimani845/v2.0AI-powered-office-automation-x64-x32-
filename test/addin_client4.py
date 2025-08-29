"""
This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
Fixed version addressing common WPS Office add-in loading issues and PyInstaller compatibility.
Enhanced with persistent single-file logging and detailed XML/WPS failure tracking, better than addin_client3.py but has issues.
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

# Single persistent log file setup
PERSISTENT_LOG_FILE = "wps_ai_addin_debug.log"

def setup_persistent_logging():
    """Setup persistent logging to a single file with detailed WPS/XML tracking"""
    try:
        # Base directory (adjust path if needed)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(base_dir, "logs")
        
        # Create logs directory with full permissions
        try:
            os.makedirs(log_dir, exist_ok=True)
            os.chmod(log_dir, 0o777)
        except PermissionError:
            print(f"Warning: Could not set full permissions on log directory: {log_dir}")
        except Exception as e:
            print(f"Warning: Error setting up log directory: {e}")
        
        # Single persistent log file path
        log_file = os.path.join(log_dir, PERSISTENT_LOG_FILE)
        
        # Test write access before setting up logging
        try:
            with open(log_file, 'a') as test_file:
                test_file.write(f"\n{'='*80}\n")
                test_file.write(f"NEW SESSION STARTED: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                test_file.write(f"{'='*80}\n")
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
                    test_log_file = os.path.join(fallback_dir, PERSISTENT_LOG_FILE)
                    with open(test_log_file, 'a') as test_file:
                        test_file.write(f"\n{'='*80}\n")
                        test_file.write(f"NEW SESSION STARTED (FALLBACK): {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                        test_file.write(f"{'='*80}\n")
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
        
        # Create formatter with more detailed information
        formatter = logging.Formatter(
            '[%(asctime)s] - PID:%(process)d - TID:%(thread)d - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Add file handler if we have a valid log file
        if log_file:
            try:
                file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
                print(f"Persistent logging enabled: {log_file}")
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

# Initialize persistent logging
current_log_file = setup_persistent_logging()

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
        
        # Force flush to ensure message is written immediately
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

# Enhanced WPS diagnostic logging
def log_wps_environment():
    """Log detailed WPS Office environment information"""
    log_message("=== WPS OFFICE ENVIRONMENT DIAGNOSTICS ===")
    
    # Check WPS Office installation
    try:
        import winreg
        wps_install_paths = [
            r"SOFTWARE\Kingsoft\Office",
            r"SOFTWARE\WPS\Office", 
            r"SOFTWARE\WPS Office"
        ]
        
        for path in wps_install_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                    try:
                        install_path = winreg.QueryValueEx(key, "InstallRoot")[0]
                        log_message(f"WPS Office found at: {path} -> {install_path}")
                    except FileNotFoundError:
                        log_message(f"WPS Office registry key exists but no InstallRoot: {path}")
            except FileNotFoundError:
                log_message(f"WPS Office registry path not found: {path}", "debug")
            except Exception as e:
                log_message(f"Error checking WPS registry path {path}: {e}", "warning")
                
    except Exception as e:
        log_message(f"Error during WPS installation check: {e}", "error")
    
    # Test WPS COM object availability
    try:
        log_message("Testing WPS COM object availability...")
        app = win32com.client.GetActiveObject("kwps.Application")
        if app:
            log_message(f"✓ WPS Application COM object found - Version: {getattr(app, 'Version', 'Unknown')}")
            log_message(f"✓ WPS Documents count: {app.Documents.Count}")
            log_message(f"✓ WPS Visible: {getattr(app, 'Visible', 'Unknown')}")
        else:
            log_message("✗ WPS Application COM object is None", "warning")
    except Exception as e:
        log_message(f"✗ Cannot access WPS COM object: {e}", "warning")
        log_message(f"Full traceback: {traceback.format_exc()}", "debug")

def log_xml_loading_attempt(ribbon_path):
    """Log detailed XML loading diagnostics"""
    log_message("=== XML RIBBON LOADING DIAGNOSTICS ===")
    log_message(f"Attempting to load ribbon XML from: {ribbon_path}")
    
    # Check file existence
    if not os.path.exists(ribbon_path):
        log_message(f"✗ CRITICAL: Ribbon XML file does NOT exist at path: {ribbon_path}", "critical")
        return False
    else:
        log_message(f"✓ Ribbon XML file exists at: {ribbon_path}")
    
    # Check file permissions
    try:
        if os.access(ribbon_path, os.R_OK):
            log_message("✓ Ribbon XML file is readable")
        else:
            log_message("✗ CRITICAL: Ribbon XML file is NOT readable", "critical")
            return False
    except Exception as e:
        log_message(f"✗ Error checking ribbon XML file permissions: {e}", "error")
        return False
    
    # Check file size
    try:
        file_size = os.path.getsize(ribbon_path)
        log_message(f"✓ Ribbon XML file size: {file_size} bytes")
        if file_size == 0:
            log_message("✗ CRITICAL: Ribbon XML file is empty!", "critical")
            return False
    except Exception as e:
        log_message(f"✗ Error getting ribbon XML file size: {e}", "error")
        return False
    
    # Try to read and validate XML content
    try:
        with open(ribbon_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        log_message(f"✓ Successfully read ribbon XML content ({len(content)} characters)")
        
        # Basic XML validation
        if not content.strip():
            log_message("✗ CRITICAL: Ribbon XML content is empty or whitespace only!", "critical")
            return False
        
        if '<customUI' not in content:
            log_message("✗ WARNING: Ribbon XML does not contain expected <customUI> element", "warning")
        else:
            log_message("✓ Ribbon XML contains expected <customUI> element")
        
        if 'xmlns=' not in content:
            log_message("✗ WARNING: Ribbon XML may be missing namespace declaration", "warning")
        else:
            log_message("✓ Ribbon XML appears to have namespace declaration")
            
        # Log first 200 characters of XML for debugging
        log_message(f"✓ Ribbon XML preview (first 200 chars): {content[:200]}...")
        
        return True
        
    except UnicodeDecodeError as e:
        log_message(f"✗ CRITICAL: Ribbon XML has encoding issues: {e}", "critical")
        return False
    except Exception as e:
        log_message(f"✗ CRITICAL: Cannot read ribbon XML content: {e}", "critical")
        log_message(f"Full traceback: {traceback.format_exc()}", "debug")
        return False

# Initial session log messages
log_message("=== NEW ADDIN SESSION STARTED ===")
log_message(f"Python Version: {sys.version}")
log_message(f"Executable Path: {sys.executable}")
log_message(f"Command Line Arguments: {sys.argv}")
log_message(f"Current Working Directory: {os.getcwd()}")
log_message(f"Persistent log file location: {current_log_file}")

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
    """Gets the running WPS Writer Application object with enhanced diagnostics."""
    log_message("=== WPS APPLICATION CONNECTION ATTEMPT ===")
    try:
        log_message("Attempting to get WPS Application COM object...")
        app = win32com.client.GetActiveObject("kwps.Application")
        
        if app:
            log_message("✓ Successfully obtained WPS Application object")
            
            # Enhanced WPS diagnostics
            try:
                version = getattr(app, 'Version', 'Unknown')
                log_message(f"✓ WPS Version: {version}")
            except Exception as e:
                log_message(f"Cannot get WPS version: {e}", "debug")
            
            try:
                doc_count = app.Documents.Count
                log_message(f"✓ WPS Documents count: {doc_count}")
            except Exception as e:
                log_message(f"Cannot get WPS documents count: {e}", "debug")
            
            try:
                visible = getattr(app, 'Visible', 'Unknown')
                log_message(f"✓ WPS Visible: {visible}")
            except Exception as e:
                log_message(f"Cannot get WPS visibility: {e}", "debug")
                
            return app
        else:
            log_message("✗ WPS Application object is None", "error")
            return None
            
    except Exception as e:
        log_message(f"✗ CRITICAL: Cannot access WPS Application COM object: {e}", "critical")
        log_message(f"This indicates WPS Office is either:")
        log_message("  1. Not running")
        log_message("  2. Not properly installed")  
        log_message("  3. COM interface is not available")
        log_message("  4. Permission issues")
        log_message(f"Full exception traceback: {traceback.format_exc()}", "debug")
        return None

def insert_text_at_cursor(text):
    """Inserts text into the active document at the current cursor position with enhanced diagnostics."""
    log_message(f"=== TEXT INSERTION ATTEMPT ===")
    log_message(f"Attempting to insert text (length: {len(text)} chars): {text[:100]}...", "debug")
    
    wps_app = get_wps_application()
    if not wps_app:
        log_message("✗ Cannot insert text: No WPS Application object available", "error")
        return
        
    if wps_app.Documents.Count == 0:
        log_message("✗ Cannot insert text: No active WPS document found", "warning")
        return
        
    try:
        # Additional document diagnostics
        active_doc = wps_app.ActiveDocument
        log_message(f"✓ Active document found: {getattr(active_doc, 'Name', 'Unknown')}")
        
        # Check selection
        selection = wps_app.Selection
        if selection:
            log_message("✓ Selection object available")
        else:
            log_message("✗ Selection object not available", "warning")
            
        # Perform text insertion
        wps_app.Selection.TypeText(Text=text)
        log_message("✓ Text successfully inserted into active WPS document.")
        
    except Exception as e:
        log_message(f"✗ CRITICAL: Error inserting text into WPS document: {e}", "critical")
        log_message(f"This could indicate:")
        log_message("  1. Document is read-only")
        log_message("  2. Selection is not available")
        log_message("  3. COM interface issues")
        log_message("  4. WPS security restrictions")
        log_message(f"Full traceback: {traceback.format_exc()}", "debug")

class WPSAddin:
    _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}" # python -c "import uuid; print(uuid.uuid4())"
    _reg_desc_ = "AI Office Automation"
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
    _reg_class_spec_ = __name__ + ".WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
        'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter',
        'GetCustomUI' # This is the method that WPS Office expects to load the XML
    ]
    # Attribute used by Microsoft office but not WPS Office
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("=== ADDIN INITIALIZATION STARTED ===")
        
        # Log WPS environment first
        log_wps_environment()
        
        try:
            ribbon_path = resource_path('ribbon.xml')
            
            # Enhanced XML loading diagnostics
            if not log_xml_loading_attempt(ribbon_path):
                log_message("FATAL: Ribbon XML loading failed during diagnostics", "critical")
                self.ribbon = ""
                return
            
            # Load the XML content
            try:
                with open(ribbon_path, 'r', encoding='utf-8') as f:
                    self.ribbon = f.read()
                log_message(f"✓ Ribbon XML loaded successfully. Content length: {len(self.ribbon)} characters")
                log_message("✓ Add-in ribbon property initialized")
            except Exception as e:
                log_message(f"✗ CRITICAL: Failed to load ribbon XML content: {e}", "critical")
                self.ribbon = ""
                return

            # Initialize translations
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
            
            log_message("✓ Translations initialized")
            log_message("=== ADDIN INITIALIZATION COMPLETED SUCCESSFULLY ===")
            
        except FileNotFoundError as e:
            log_message(f"✗ FATAL ERROR: Ribbon XML file not found: {e}", "critical")
            self.ribbon = ""
        except Exception as e:
            log_message(f"✗ FATAL ERROR IN ADDIN __init__: {e}", "critical")
            log_message(f"Full traceback: {traceback.format_exc()}", "error")
            self.ribbon = ""
            
    
    # Explicit Ribbon Loader for WPS
    def GetCustomUI(self, ribbon_id):
        """
        WPS will call this to request the ribbon XML.
        This is the critical method for WPS Office integration.
        """
        log_message(f"=== WPS GETCUSTOMUI METHOD CALLED ===")
        log_message(f"Ribbon ID requested by WPS: {ribbon_id}")
        log_message(f"This method is CRITICAL for WPS Office add-in loading")
        
        try:
            # Check if we have cached ribbon content
            if hasattr(self, 'ribbon') and self.ribbon:
                log_message(f"✓ Returning cached ribbon XML (length: {len(self.ribbon)} chars)")
                log_message("✓ WPS GetCustomUI call SUCCESSFUL")
                return self.ribbon
            else:
                log_message("✗ No cached ribbon content, attempting to reload from file", "warning")
                
            # Fallback: reload from file
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Reloading ribbon XML from: {ribbon_path}")
            
            if not log_xml_loading_attempt(ribbon_path):
                log_message("✗ CRITICAL: Ribbon XML reload failed in GetCustomUI", "critical")
                return None
                
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            log_message(f"✓ Successfully reloaded ribbon XML (length: {len(content)} chars)")
            log_message("✓ WPS GetCustomUI call SUCCESSFUL (reloaded)")
            return content
            
        except Exception as e:
            log_message(f"✗ CRITICAL FAILURE in GetCustomUI: {e}", "critical")
            log_message("✗ This will prevent the WPS add-in from loading properly!")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
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
        log_message(f"WPS GetTabLabel called, returning: '{result}'", "debug")
        return result
        
    def GetGroupLabel(self, c): 
        result = self._get_localized_string("group")
        log_message(f"WPS GetGroupLabel called, returning: '{result}'", "debug")
        return result
        
    def GetRunPromptLabel(self, c): 
        result = self._get_localized_string("run_prompt")
        log_message(f"WPS GetRunPromptLabel called, returning: '{result}'", "debug")
        return result
        
    def GetAnalyzeDocLabel(self, c): 
        result = self._get_localized_string("analyze_doc")
        log_message(f"WPS GetAnalyzeDocLabel called, returning: '{result}'", "debug")
        return result
        
    def GetSummarizeDocLabel(self, c): 
        result = self._get_localized_string("summarize_doc")
        log_message(f"WPS GetSummarizeDocLabel called, returning: '{result}'", "debug")
        return result
        
    def GetCreateMemoLabel(self, c): 
        result = self._get_localized_string("create_memo")
        log_message(f"WPS GetCreateMemoLabel called, returning: '{result}'", "debug")
        return result
        
    def GetCreateMinutesLabel(self, c): 
        result = self._get_localized_string("create_minutes")
        log_message(f"WPS GetCreateMinutesLabel called, returning: '{result}'", "debug")
        return result
        
    def GetCreateCoverLetterLabel(self, c): 
        result = self._get_localized_string("create_cover_letter")
        log_message(f"WPS GetCreateCoverLetterLabel called, returning: '{result}'", "debug")
        return result

    def OnLoadImage(self, imageName):
        log_message(f"=== WPS IMAGE LOADING REQUEST ===")
        image_path = resource_path(f"{imageName}.png")
        log_message(f"WPS requested image: '{imageName}'")
        log_message(f"Looking for image at: {image_path}")
        
        try:
            if not os.path.exists(image_path):
                log_message(f"✗ Image file does not exist: {image_path}", "warning")
                log_message("This will cause the button to appear without an icon")
                return None
                
            log_message(f"✓ Image file exists: {image_path}")
            img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
            
            if img_handle:
                log_message(f"✓ Successfully loaded image '{imageName}' with handle: {img_handle}")
            else:
                log_message(f"✗ Failed to get image handle for '{imageName}'", "warning")
                
            return img_handle
            
        except Exception as e:
            log_message(f"✗ ERROR: Failed to load image '{imageName}': {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return None

    def _call_backend_task(self, endpoint: str, payload: dict):
        log_message(f"=== BACKEND API CALL STARTED ===")
        log_message(f"Endpoint: {endpoint}")
        log_message(f"Payload: {payload}", "debug")
        
        try:
            # Notify user that we're contacting the server
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
            log_message(f"Making POST request to: {BACKEND_URL}{endpoint}")
            
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
            log_message(f"✓ Received response status: {response.status_code}")
            
            response.raise_for_status()
            result_data = response.json()
            log_message(f"✓ Response data received", "debug")
            
            result = result_data.get("result", "")
            if result:
                log_message(f"✓ Backend returned result (length: {len(result)} chars)")
            else:
                log_message("✗ Backend returned empty result", "warning")
                
            header = self._get_localized_string("result_header")
            footer = self._get_localized_string("result_footer")
            insert_text_at_cursor(f"{header}{result}{footer}")
            log_message(f"✓ Successfully processed backend response from {endpoint}")
            
        except requests.exceptions.ConnectionError as e:
            log_message(f"✗ CONNECTION ERROR to {endpoint}: {e}", "error")
            log_message("Backend server is likely not running or not accessible")
            insert_text_at_cursor(self._get_localized_string("connection_error"))
        except requests.exceptions.Timeout as e:
            log_message(f"✗ TIMEOUT ERROR for {endpoint}: {e}", "error")
            insert_text_at_cursor("\n\nERROR: Request timed out. Please try again.\n\n")
        except Exception as e:
            log_message(f"✗ UNEXPECTED ERROR calling {endpoint}: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))

    def OnRunPrompt(self, c):
        log_message("=== WPS ONRUNPROMPT BUTTON CLICKED ===")
        try:
            root = Tk(); root.withdraw()
            prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                            self._get_localized_string("prompt_message"))
            root.destroy()
            
            if not prompt:
                log_message("User cancelled prompt input")
                return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
            
            log_message(f"✓ User entered prompt: {prompt}")
            threading.Thread(target=self._call_backend_task, 
                            args=("/process", {"prompt": prompt})).start()
        except Exception as e:
            log_message(f"✗ Error in OnRunPrompt: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnAnalyzeDocument(self, c):
        log_message("=== WPS ONANALYZEDOCUMENT BUTTON CLICKED ===")
        try:
            wps_app = get_wps_application()
            if not wps_app or wps_app.Documents.Count == 0: 
                log_message("✗ No active WPS document found", "warning")
                return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
            
            content = wps_app.ActiveDocument.Content.Text
            log_message(f"✓ Document content extracted (length: {len(content)} characters)")
            threading.Thread(target=self._call_backend_task, 
                            args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()
        except Exception as e:
            log_message(f"✗ Error in OnAnalyzeDocument: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnSummarizeDocument(self, c):
        log_message("=== WPS ONSUMMARIZEDOCUMENT BUTTON CLICKED ===")
        try:
            wps_app = get_wps_application()
            if not wps_app or wps_app.Documents.Count == 0: 
                log_message("✗ No active WPS document found", "warning")
                return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
            
            content = wps_app.ActiveDocument.Content.Text
            log_message(f"✓ Document content extracted (length: {len(content)} characters)")
            threading.Thread(target=self._call_backend_task, 
                            args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()
        except Exception as e:
            log_message(f"✗ Error in OnSummarizeDocument: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateMemo(self, c):
        log_message("=== WPS ONCREATEMEMO BUTTON CLICKED ===")
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
            log_message(f"✓ Memo payload prepared: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()
        except Exception as e:
            log_message(f"✗ Error in OnCreateMemo: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateMinutes(self, c):
        log_message("=== WPS ONCREATEMINUTES BUTTON CLICKED ===")
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
            log_message(f"✓ Minutes payload prepared: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()
        except Exception as e:
            log_message(f"✗ Error in OnCreateMinutes: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")

    def OnCreateCoverLetter(self, c):
        log_message("=== WPS ONCREATE COVERLETTER BUTTON CLICKED ===")
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
            log_message(f"✓ Cover letter payload prepared: {payload}")
            threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()
        except Exception as e:
            log_message(f"✗ Error in OnCreateCoverLetter: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")


if __name__ == '__main__':
    
    log_message("=== MAIN EXECUTION BLOCK STARTED ===")

    def is_pyinstaller_bundle():
        """Check if running as PyInstaller bundle"""
        result = getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')
        log_message(f"Is PyInstaller bundle: {result}", "debug")
        return result
    
    def register_com_server_pyinstaller(cls):
        """Register COM server for PyInstaller executable with enhanced logging"""
        log_message("=== STARTING PYINSTALLER COM SERVER REGISTRATION ===")
        import winreg
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_
        
        # Get the executable path
        exe_path = sys.executable if is_pyinstaller_bundle() else __file__
        
        log_message(f"Registering PyInstaller COM server:")
        log_message(f"  CLSID: {clsid}")
        log_message(f"  ProgID: {progid}")
        log_message(f"  Description: {desc}")
        log_message(f"  Executable: {exe_path}")
        
        try:
            # Register main CLSID
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                log_message(f"✓ Created CLSID entry: {clsid}")
            
            # Register LocalServer32 (not InprocServer32 for exe)
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{exe_path}"')
                log_message(f"✓ Created LocalServer32 entry: {exe_path}")
            
            # Register ProgID
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, progid)
                log_message(f"✓ Created ProgID entry: {progid}")
            
            # Register ProgID mapping
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                with winreg.CreateKeyEx(key, "CLSID") as clsid_key:
                    winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
                log_message(f"✓ Created ProgID mapping")
            
            log_message("✓ PyInstaller COM server registered successfully")
            return True
            
        except PermissionError as e:
            log_message(f"✗ PERMISSION ERROR during COM registration: {e}", "critical")
            log_message("This usually means the script needs to be run as Administrator", "critical")
            return False
        except Exception as e:
            log_message(f"✗ Failed to register PyInstaller COM server: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return False
    
    def register_com_server_python(cls):
        """Register COM server for regular Python execution with enhanced logging"""
        log_message("=== STARTING PYTHON COM SERVER REGISTRATION ===")
        try:
            import win32com.server.register
            log_message(f"Registering Python COM server:")
            log_message(f"  CLSID: {cls._reg_clsid_}")
            log_message(f"  ProgID: {cls._reg_progid_}")
            log_message(f"  Class spec: {cls._reg_class_spec_}")
            
            win32com.server.register.UseCommandLine(cls)
            log_message("✓ Python COM server registered successfully")
            return True
        except PermissionError as e:
            log_message(f"✗ PERMISSION ERROR during Python COM registration: {e}", "critical")
            log_message("This usually means the script needs to be run as Administrator", "critical")
            return False
        except Exception as e:
            log_message(f"✗ Failed to register Python COM server: {e}", "error")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
            return False
    
    def register_wps_addin_entry(clsid, progid, description):
        """
        Creates the specific registry entry that WPS Office looks for with enhanced logging.
        """
        log_message("=== STARTING WPS ADDIN REGISTRY ENTRY CREATION ===")
        import winreg
        log_message(f"Creating WPS-specific registry entries:")
        log_message(f"  CLSID: {clsid}")
        log_message(f"  ProgID: {progid}")
        log_message(f"  Description: {description}")
        
        # Try multiple possible WPS registry paths
        wps_addin_paths = [
            r"Software\Kingsoft\Office\Addins",
            r"Software\WPS\Office\Addins", 
            r"Software\WPS Office\Addins"
        ]
        
        registration_succeeded = False
        
        for base_path in wps_addin_paths:
            try:
                log_message(f"Attempting registry path: HKCU\\{base_path}")
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_CREATE_SUB_KEY) as parent_key:
                    log_message(f"✓ Successfully opened/created parent key at HKCU\\{base_path}")
                    
                    with winreg.CreateKeyEx(parent_key, progid) as key:
                        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, description)
                        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                        winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                        winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
                        winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
                        
                        log_message(f"✓ Successfully created WPS Add-in entry '{progid}' at {base_path}")
                        log_message(f"  - Description: {description}")
                        log_message(f"  - FriendlyName: AI Assistant")
                        log_message(f"  - LoadBehavior: 3 (Load at startup)")
                        log_message(f"  - CLSID: {clsid}")
                        log_message(f"  - CommandLineSafe: 0")
                        
                        registration_succeeded = True
                        break  # Success, no need to try other paths
                        
            except PermissionError as e:
                log_message(f"✗ PERMISSION ERROR at {base_path}: {e}", "warning")
                log_message("This path requires Administrator privileges", "warning")
                continue
            except Exception as e:
                log_message(f"✗ Could not register at {base_path}: {e}", "warning")
                continue
        
        if registration_succeeded:
            log_message("✓ WPS add-in registry entries created successfully")
        else:
            log_message("✗ FAILED to create any WPS add-in registry entries", "critical")
            
        return registration_succeeded
    
    def register_server(cls):
        """Register COM server with comprehensive logging - handles both Python and PyInstaller"""
        log_message("=== STARTING COMPREHENSIVE REGISTRATION PROCESS ===")
        
        # Step 1: Register COM server (different methods for Python vs PyInstaller)
        log_message("STEP 1: Registering COM server...")
        if is_pyinstaller_bundle():
            com_success = register_com_server_pyinstaller(cls)
        else:
            com_success = register_com_server_python(cls)
            
        if not com_success:
            log_message("✗ FATAL: COM server registration failed", "critical")
            log_message("Cannot proceed without COM server registration", "critical")
            print("FATAL: Failed to register the COM server. Please run as Administrator.")
            return False
        
        log_message("✓ COM server registration completed successfully")
            
        # Step 2: Register WPS Office add-in entries
        log_message("STEP 2: Registering WPS Office add-in entries...")
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_
        
        if register_wps_addin_entry(clsid, progid, desc):
            log_message("✓ WPS add-in registration successful")
            log_message("=== REGISTRATION PROCESS COMPLETED SUCCESSFULLY ===")
            print("SUCCESS: Add-in registered successfully!")
            print("Please restart WPS Office to see the add-in.")
            print(f"Check log file for details: {current_log_file}")
            return True
        else:
            log_message("✗ WPS add-in registration failed", "error")
            log_message("COM server is registered but WPS won't see the add-in", "error")
            print("FAILED: Could not register WPS Office add-in entry.")
            print("The COM server is registered but WPS Office won't see the add-in.")
            print(f"Check log file for details: {current_log_file}")
            return False

    def unregister_server(cls):
        """Enhanced unregistration with comprehensive logging"""
        log_message("=== STARTING COMPREHENSIVE UNREGISTRATION PROCESS ===")
        import winreg
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        
        log_message(f"Unregistering add-in:")
        log_message(f"  CLSID: {clsid}")
        log_message(f"  ProgID: {progid}")
        
        # Step 1: Unregister COM server
        log_message("STEP 1: Unregistering COM server...")
        if not is_pyinstaller_bundle():
            try:
                import win32com.server.register
                win32com.server.register.UnregisterServer(clsid)
                log_message("✓ Python COM server unregistered successfully")
            except Exception as e:
                log_message(f"✗ Could not unregister Python COM server: {e}", "warning")
        else:
            log_message("PyInstaller bundle - COM server will be removed via registry cleanup")

        # Step 2: Remove COM server registry entries
        log_message("STEP 2: Removing COM server registry entries...")
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
                log_message(f"✓ Removed COM entry: HKCR\\{path}")
            except FileNotFoundError:
                log_message(f"COM entry not found (already removed): HKCR\\{path}", "debug")
            except Exception as e:
                log_message(f"✗ Could not remove COM entry HKCR\\{path}: {e}", "warning")
        
        # Step 3: Remove WPS Office registry entries
        log_message("STEP 3: Removing WPS Office registry entries...")
        wps_addin_paths = [
            f"Software\\Kingsoft\\Office\\Addins\\{progid}",
            f"Software\\WPS\\Office\\Addins\\{progid}",
            f"Software\\WPS Office\\Addins\\{progid}"
        ]
        
        for path in wps_addin_paths:
            try:
                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0)
                log_message(f"✓ Removed WPS entry: HKCU\\{path}")
            except FileNotFoundError:
                log_message(f"WPS entry not found (already removed): HKCU\\{path}", "debug")
            except Exception as e:
                log_message(f"✗ Could not remove HKCU\\{path}: {e}", "warning")

        log_message("=== UNREGISTRATION PROCESS COMPLETED ===")
        print("Unregistration complete.")
        print(f"Check log file for details: {current_log_file}")
    
    def run_com_server():
        """Run as COM server when called by Windows with enhanced logging"""
        log_message("=== STARTING COM SERVER FOR WPS OFFICE ===")
        log_message("This process will handle WPS Office requests for the add-in")
        
        try:
            import win32com.server.localserver
            log_message(f"✓ Serving COM server with CLSID: {WPSAddin._reg_clsid_}")
            log_message(f"✓ ProgID: {WPSAddin._reg_progid_}")
            log_message("✓ COM server is now running and waiting for WPS Office requests...")
            log_message("WPS Office should now be able to load the add-in")
            
            win32com.server.localserver.serve([WPSAddin._reg_clsid_])
            
        except Exception as e:
            log_message(f"✗ CRITICAL ERROR running COM server: {e}", "critical")
            log_message("WPS Office will not be able to load the add-in", "critical")
            log_message(f"Full traceback: {traceback.format_exc()}", "debug")
    
    def check_environment():
        """Check and setup environment with comprehensive diagnostics"""
        log_message("=== COMPREHENSIVE ENVIRONMENT CHECK ===")
        
        # Python environment
        log_message("PYTHON ENVIRONMENT:")
        log_message(f"  Python version: {sys.version}")
        log_message(f"  Python executable: {sys.executable}")
        log_message(f"  Running as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        
        # Path information
        if not is_pyinstaller_bundle():
            script_dir = os.path.dirname(os.path.abspath(__file__))
            if script_dir not in sys.path:
                sys.path.insert(0, script_dir)
                log_message(f"✓ Added script directory to Python path: {script_dir}")
        
        log_message(f"  Current directory: {os.getcwd()}")
        log_message(f"  Python path (first 3): {sys.path[:3]}")
        
        # Check for required modules
        log_message("MODULE AVAILABILITY CHECK:")
        required_modules = {
            'win32com.client': 'Windows COM interface',
            'win32api': 'Windows API access', 
            'winreg': 'Windows registry access',
            'pythoncom': 'Python COM support',
            'requests': 'HTTP requests to backend',
            'tkinter': 'User input dialogs'
        }
        
        all_modules_ok = True
        for module, description in required_modules.items():
            try:
                __import__(module)
                log_message(f"  ✓ {module}: Available ({description})")
            except ImportError as e:
                log_message(f"  ✗ {module}: MISSING ({description}) - {e}", "error")
                all_modules_ok = False
        
        if all_modules_ok:
            log_message("✓ All required modules are available")
        else:
            log_message("✗ Some required modules are missing - add-in may not work properly", "warning")
        
        # Resource files check
        log_message("RESOURCE FILES CHECK:")
        ribbon_path = resource_path('ribbon.xml')
        if os.path.exists(ribbon_path):
            log_message(f"  ✓ ribbon.xml found at: {ribbon_path}")
        else:
            log_message(f"  ✗ ribbon.xml NOT found at: {ribbon_path}", "critical")
        
        log_message("=== ENVIRONMENT CHECK COMPLETED ===")

    # Main command execution logic
    log_message("=== STARTING MAIN EXECUTION LOGIC ===")
    check_environment()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        log_message(f"Command line argument received: '{command}'")
        
        if command == '/regserver':
            log_message("=== EXECUTING REGISTRATION COMMAND ===")
            register_server(WPSAddin)
        elif command == '/unregserver':
            log_message("=== EXECUTING UNREGISTRATION COMMAND ===")
            unregister_server(WPSAddin)
        elif command == '/embedding':
            # This is called by Windows when WPS tries to instantiate the COM object
            log_message("=== EXECUTING EMBEDDING COMMAND (COM SERVER MODE) ===")
            log_message("WPS Office is requesting the COM server - this is good!")
            run_com_server()
        else:
            log_message(f"✗ Unknown command line argument: '{command}'", "warning")
            print(f"Unknown command: {command}")
    else:
        log_message("=== NO COMMAND LINE ARGUMENTS - SHOWING USAGE ===")
        print("WPS Office AI Assistant Add-in")
        print("Usage:")
        print("  /regserver   - Register the add-in")
        print("  /unregserver - Unregister the add-in")
        print(f"\nRunning as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        print(f"Log file: {current_log_file}")
        input("\nPress Enter to exit...")
    
    log_message("=== MAIN SCRIPT EXECUTION COMPLETED ===")
    
    # Final log flush to ensure all messages are written
    try:
        for handler in logging.getLogger().handlers:
            if hasattr(handler, 'flush'):
                handler.flush()
        log_message("Final log flush completed")
    except Exception as e:
        print(f"Error during final log flush: {e}")