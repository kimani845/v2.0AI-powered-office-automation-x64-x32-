# addin_common.py

import datetime
import traceback
import os
import sys
import threading
import requests
import win32com.client
import win32api
import logging
import pythoncom
from tkinter import simpledialog, Tk

# file-based logging setup
# (Keep the logging setup from your original code)
try:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "addin_debug.log")
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='[%(asctime)s] - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    def log_message(message):
        logging.info(message)
        print(f"LOG: {message}")

    log_message("--- COMMON SCRIPT IMPORTED ---")
    log_message(f"Python Version: {sys.version}")

except Exception as e:
    print(f"FATAL: Could not set up logging. Error: {e}")

# Configuration - BACKEND IP address
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

# Core COM class
class WPSAddin:
    # Use different CLSIDs for 32-bit and 64-bit versions
    _reg_clsid_ = None 
    _reg_desc_ = "AI Office Automation"
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
    _reg_class_spec_ = "wps_addin.WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
        'GetCreateCoverLetter', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter',
        'GetCustomUI' # This is the method that WPS Office expects to load the XML

    ]
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("--- Add-in __init__ started ---")
        try:
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Attempting to load ribbon from: {ribbon_path}")
            
            if not os.path.exists(ribbon_path):
                log_message(f"FATAL: Ribbon XML file does NOT exist at the path.")
                raise FileNotFoundError(f"Ribbon XML not found at {ribbon_path}")
            
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                self.ribbon = f.read()
            log_message("Ribbon XML loaded successfully.")
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
        except FileNotFoundError:
            log_message(f"FATAL ERROR: Ribbon XML file not found at {ribbon_path}")
            self.ribbon = ""
        except Exception as e:
            log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}")
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
            
            # if not log_xml_loading_attempt(ribbon_path):
            #     log_message("✗ CRITICAL: Ribbon XML reload failed in GetCustomUI", "critical")
            #     return None
                
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
                self.ribbon = content
                log_message(f"✓ Successfully reloaded ribbon XML (length: {len(content)} chars)")
                return content
                    
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
        
        
    def register_wps_addin_entry(clsid, progid, description):
        """
        Creates the specific registry entry that WPS Office looks for.
        """
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
                log_message(f"Could not register at {base_path}: {e}")
                continue
        
        return registration_succeeded
    