"""
This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
Fixed version addressing common WPS Office add-in loading issues and PyInstaller compatibility.
"""
import datetime
import traceback
import os
import sys
import threading
import requests
import win32com.client
import win32api
import winreg
import pythoncom
import logging
from tkinter import simpledialog, Tk


# file-based logging 
try:
    # log_dir = os.path.join(os.getenv('APPDATA'), 'WPS_AI_Addin_Logs')
    # os.makedirs(log_dir, exist_ok=True)
    # log_file = os.path.join(log_dir, 'addin_debug.log')
    
# Base directory (adjust path if needed)
    base_dir = os.path.dirname(os.path.abspath(__file__)) # Logs directory inside project
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    # Log file path
    log_file = os.path.join(log_dir, "addin_debug.log")

    # Configure logging
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='[%(asctime)s] - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # A function to replace the old print-based logging
    def log_message(message):
        logging.info(message)
        # Also print to console if it exists, for command-line use
        print(f"LOG: {message}")

    log_message("--- SCRIPT EXECUTION STARTED ---")
    log_message(f"Python Version: {sys.version}")
    log_message(f"Executable Path: {sys.executable}")
    log_message(f"Command Line Arguments: {sys.argv}")

except Exception as e:
    # If logging setup fails, we can't do much, but we try to inform the user.
    # This is a last resort.
    print(f"FATAL: Could not set up logging. Error: {e}")
    
    
# Configuration - IP address
BACKEND_URL = "http://127.0.0.1:8000"

# Consistent naming
WPS_ADDIN_ENTRY_NAME = "WPSAIAddin.Connect"


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller bundling """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_wps_application():
    """Gets the running WPS Writer Application object."""
    try:
        return win32com.client.GetActiveObject("kwps.Application")
    except Exception as e:
        log_message(f"Error getting WPS Application object: {e}")
        return None

def insert_text_at_cursor(text):
    """Inserts text into the active document at the current cursor position."""
    wps_app = get_wps_application()
    if wps_app and wps_app.Documents.Count > 0:
        try:
            wps_app.Selection.TypeText(Text=text)
            log_message("Text successfully inserted into active WPS document.")
        except Exception as e:
            log_message(f"Error inserting text into WPS document: {e}\n{traceback.format_exc()}")
    else:
        log_message("Warning: Could not find an active WPS document to insert text into.")

class WPSAddin:
    _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
    _reg_desc_ = "AI Office Automation"
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
    _reg_class_spec_ = __name__ + ".WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
        'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter'
    ]
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("--- WPSAdd-in __init__ started ---")
        logging.getLogger().handlers[0].flush()
        
        try:
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Attempting to load ribbon from: {ribbon_path}")
            logging.getLogger().handlers[0].flush()
            
            if not os.path.exists(ribbon_path):
                log_message(f"FATAL: Ribbon XML file does NOT exist at the path.")
                raise FileNotFoundError(f"Ribbon XML not found at {ribbon_path}")
            logging.getLogger().handlers[0].flush()
            
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                self.ribbon = f.read()
            log_message("Ribbon XML loaded successfully.")
            logging.getLogger().handlers[0].flush()

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
            
            log_message("--- Add-in __init__ completed successfully. ---")
            logging.getLogger().handlers[0].flush()
        except FileNotFoundError:
            log_message(f"FATAL ERROR: Ribbon XML file not found at {ribbon_path}")
            self.ribbon = ""
        except Exception as e:
            log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}")
            self.ribbon = ""

    def _get_localized_string(self, key):
        lang_id = 1033
        wps_app = get_wps_application()
        if wps_app:
            try:
                lang_id = wps_app.LanguageSettings.LanguageID(1)
            except Exception:
                pass
        return self.translations.get(lang_id, self.translations[1033]).get(key, key)

    def GetTabLabel(self, c): return self._get_localized_string("tab")
    def GetGroupLabel(self, c): return self._get_localized_string("group")
    def GetRunPromptLabel(self, c): return self._get_localized_string("run_prompt")
    def GetAnalyzeDocLabel(self, c): return self._get_localized_string("analyze_doc")
    def GetSummarizeDocLabel(self, c): return self._get_localized_string("summarize_doc")
    def GetCreateMemoLabel(self, c): return self._get_localized_string("create_memo")
    def GetCreateMinutesLabel(self, c): return self._get_localized_string("create_minutes")
    def GetCreateCoverLetterLabel(self, c): return self._get_localized_string("create_cover_letter")

    def OnLoadImage(self, imageName):
        image_path = resource_path(f"{imageName}.png")
        log_message(f"Attempting to load image: {image_path}")
        try:
            img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
            log_message(f"Successfully loaded image '{imageName}'.")
            return img_handle
        except Exception as e:
            log_message(f"ERROR: Failed to load image '{imageName}': {e}")
            return None

    def _call_backend_task(self, endpoint: str, payload: dict):
        log_message(f"Calling backend endpoint: {endpoint}")
        try:
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
            response.raise_for_status()
            result = response.json().get("result", "")
            header = self._get_localized_string("result_header")
            footer = self._get_localized_string("result_footer")
            insert_text_at_cursor(f"{header}{result}{footer}")
            log_message(f"Successfully received response from {endpoint}.")
        except requests.exceptions.ConnectionError:
            log_message(f"Connection error to {endpoint}")
            insert_text_at_cursor(self._get_localized_string("connection_error"))
        except Exception as e:
            log_message(f"Error calling {endpoint}: {e}")
            insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))

    def OnRunPrompt(self, c):
        root = Tk(); root.withdraw()
        prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("prompt_message"))
        root.destroy()
        if not prompt: 
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        threading.Thread(target=self._call_backend_task, 
                        args=("/process", {"prompt": prompt})).start()

    def OnAnalyzeDocument(self, c):
        wps_app = get_wps_application()
        if not wps_app or wps_app.Documents.Count == 0: 
            return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
        content = wps_app.ActiveDocument.Content.Text
        threading.Thread(target=self._call_backend_task, 
                        args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()

    def OnSummarizeDocument(self, c):
        wps_app = get_wps_application()
        if not wps_app or wps_app.Documents.Count == 0: 
            return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
        content = wps_app.ActiveDocument.Content.Text
        threading.Thread(target=self._call_backend_task, 
                        args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()

    def OnCreateMemo(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("memo_topic"))
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("memo_audience"))
        root.destroy()
        payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
        threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()

    def OnCreateMinutes(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("minutes_topic"))
        if not topic:
            root.destroy()
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
        threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()

    def OnCreateCoverLetter(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("cover_letter_topic"))
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
                                        self._get_localized_string("cover_letter_audience"))
        root.destroy()
        payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
        threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()


if __name__ == '__main__':
    
    log_message("Executing main block.")
    
    def is_pyinstaller_bundle():
        """Check if running as PyInstaller bundle"""
        return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

    def register_server(cls):
        """Registers the COM server and WPS Add-in entry."""
        import winreg
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_
        
        is_64bit_process = sys.maxsize > 2**32
        reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY
        
        log_message(f"Starting registration for {'64-bit' if is_64bit_process else '32-bit'} view.")
        
        try:
            # Create the CLSID key with its subkeys
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, winreg.KEY_WRITE | reg_view_flag) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
            
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", 0, winreg.KEY_WRITE | reg_view_flag) as server_key:
                executable_path = sys.executable
                winreg.SetValueEx(server_key, "", 0, winreg.REG_SZ, f'"{executable_path}"')
            
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, winreg.KEY_WRITE | reg_view_flag) as progid_key:
                winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
            
            # Create the ProgID-to-CLSID mapping
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, winreg.KEY_WRITE | reg_view_flag) as key:
                with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
                    winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
            
            log_message("Core COM server registered successfully.")
        except Exception as e:
            log_message(f"FATAL: Core COM registration failed: {e}")
            print("FATAL: Failed to register the COM server. Please run as Administrator.")
            return

        # Create the specific WPS Office add-in entry
        wps_addin_paths = [ f"Software\\Kingsoft\\Office\\Addins\\{progid}" ]
        for path in wps_addin_paths:
            try:
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
                    winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
                    winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                    winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                    winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
                log_message(f"SUCCESS: WPS add-in entry created at HKCU\\{path}")
                print("SUCCESS: Add-in registered successfully!")
                print("Please restart WPS Office to see the add-in.")
                return
            except Exception:
                continue
        
        log_message("FAILED: Could not create WPS add-in entry in any known location.")
        print("FAILED: Could not register the WPS Office add-in entry.")

    def unregister_server(cls):
        """FIXED: Robustly unregisters the COM server and WPS Add-in entry."""
        import winreg
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        is_64bit_process = sys.maxsize > 2**32
        reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY
        
        log_message(f"Starting unregistration for {'64-bit' if is_64bit_process else '32-bit'} view.")

        # Remove WPS Office entry
        try:
            winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, f"Software\\Kingsoft\\Office\\Addins\\{progid}", 0, 0)
            log_message("Removed WPS Add-in entry.")
        except FileNotFoundError:
            pass

        # Remove COM server entries
        try:
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, progid, reg_view_flag, 0)
            log_message("Removed core COM server entries.")
        except FileNotFoundError:
            pass
        except Exception as e:
            log_message(f"Error during COM unregistration: {e}")
            
        print("Unregistration complete.")


    def run_com_server():
        log_message("Preparing to start COM server...")
        try:
            # This is where add-in gets created when WPS connects
            pythoncom.CoInitialize()
            factory = pythoncom.MakePyFactory(WPSAddin)
            clsid = WPSAddin._reg_clsid_
            pythoncom.CoRegisterClassObject(clsid, factory, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.REGCLS_MULTIPLEUSE)
            log_message("COM Class Object registered. Starting message pump.")
            pythoncom.PumpMessages()
            pythoncom.CoUninitialize()
            log_message("COM server shut down gracefully.")
        except Exception as e:
            log_message(f"FATAL ERROR while running COM server: {traceback.format_exc()}")

    # Main execution logic
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            log_message("Command: /regserver")
            register_server(WPSAddin)
        elif sys.argv[1].lower() == '/unregserver':
            log_message("Command: /unregserver")
            unregister_server(WPSAddin)
        else:
            log_message(f"Command: {sys.argv[1]} - Starting COM Server.")
            run_com_server()
    else:
        log_message("Command: No arguments - Starting COM Server.")
        run_com_server()