# """
# This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
# Fixed version addressing common WPS Office add-in loading issues and PyInstaller compatibility.
# Enhanced with better 32/64-bit registry handling and comprehensive logging.
# """
# import datetime
# import traceback
# import os
# import sys
# import threading
# import requests
# import win32com.client
# import win32api
# import winreg
# import pythoncom
# import logging
# from tkinter import simpledialog, Tk
# import ctypes
# from ctypes import wintypes


# def is_admin():
#     """Check if the current process has administrator privileges."""
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False

# def run_as_admin():
#     """Re-run the current script with administrator privileges."""
#     if is_admin():
#         return True
#     else:
#         try:
#             # Re-run the program with admin rights
#             ctypes.windll.shell32.ShellExecuteW(
#                 None, "runas", sys.executable, " ".join(sys.argv), None, 1
#             )
#             return False
#         except:
#             print("Failed to elevate privileges. Please run as Administrator.")
#             return False

# # Enhanced file-based logging 
# try:
#     base_dir = os.path.dirname(os.path.abspath(__file__))
#     log_dir = os.path.join(base_dir, "logs")
#     os.makedirs(log_dir, exist_ok=True)
#     log_file = os.path.join(log_dir, "addin_debug.log")

#     # Configure logging with more detailed format
#     logging.basicConfig(
#         filename=log_file,
#         level=logging.DEBUG,  # Changed to DEBUG for more details
#         format='[%(asctime)s] - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
#         datefmt='%Y-%m-%d %H:%M:%S'
#     )

#     # Also log to console
#     console_handler = logging.StreamHandler()
#     console_handler.setLevel(logging.INFO)
#     console_formatter = logging.Formatter('[%(levelname)s] %(message)s')
#     console_handler.setFormatter(console_formatter)
#     logging.getLogger().addHandler(console_handler)

#     def log_message(message, level="INFO"):
#         if level.upper() == "DEBUG":
#             logging.debug(message)
#         elif level.upper() == "WARNING":
#             logging.warning(message)
#         elif level.upper() == "ERROR":
#             logging.error(message)
#         else:
#             logging.info(message)

#     log_message("--- SCRIPT EXECUTION STARTED ---")
#     log_message(f"Python Version: {sys.version}")
#     log_message(f"Executable Path: {sys.executable}")
#     log_message(f"Command Line Arguments: {sys.argv}")
#     log_message(f"Process Architecture: {'64-bit' if sys.maxsize > 2**32 else '32-bit'}")
#     log_message(f"Admin Rights: {is_admin()}")

# except Exception as e:
#     print(f"FATAL: Could not set up logging. Error: {e}")
    
    
# # Configuration - IP address
# BACKEND_URL = "http://127.0.0.1:8000"

# # Consistent naming
# WPS_ADDIN_ENTRY_NAME = "WPSAIAddin.Connect"


# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller bundling """
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)

# def is_pyinstaller_bundle():
#     """Check if running as PyInstaller bundle"""
#     return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

# def get_wps_application():
#     """Gets the running WPS Writer Application object."""
#     try:
#         return win32com.client.GetActiveObject("kwps.Application")
#     except Exception as e:
#         log_message(f"Error getting WPS Application object: {e}")
#         return None

# def insert_text_at_cursor(text):
#     """Inserts text into the active document at the current cursor position."""
#     wps_app = get_wps_application()
#     if wps_app and wps_app.Documents.Count > 0:
#         try:
#             wps_app.Selection.TypeText(Text=text)
#             log_message("Text successfully inserted into active WPS document.")
#         except Exception as e:
#             log_message(f"Error inserting text into WPS document: {e}\n{traceback.format_exc()}")
#     else:
#         log_message("Warning: Could not find an active WPS document to insert text into.")

# class WPSAddin:
#     _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}"
#     _reg_desc_ = "AI Office Automation"
#     _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
#     _reg_class_spec_ = "wps_addin_fixed.WPSAddin"

#     _public_methods_ = [
#         'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
#         'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
#         'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
#         'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter',
#         'GetCustomUI'
#     ]
#     _public_attrs_ = ['ribbon']

#     def __init__(self):
#         log_message("--- WPSAdd-in __init__ started ---")
        
#         try:
#             ribbon_path = resource_path('ribbon.xml')
#             log_message(f"Attempting to load ribbon from: {ribbon_path}")
            
#             if not os.path.exists(ribbon_path):
#                 log_message(f"FATAL: Ribbon XML file does NOT exist at the path.", "ERROR")
#                 # Create a basic ribbon XML if file doesn't exist
#                 self.ribbon = '''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
#                     <ribbon>
#                         <tabs>
#                             <tab id="AITab" label="AI Assistant">
#                                 <group id="AIGroup" label="AI Tools">
#                                     <button id="RunPrompt" label="Run Prompt" onAction="OnRunPrompt"/>
#                                     <button id="AnalyzeDoc" label="Analyze Document" onAction="OnAnalyzeDocument"/>
#                                     <button id="SummarizeDoc" label="Summarize Document" onAction="OnSummarizeDocument"/>
#                                 </group>
#                             </tab>
#                         </tabs>
#                     </ribbon>
#                 </customUI>'''
#                 log_message("Using default ribbon XML")
#             else:
#                 with open(ribbon_path, 'r', encoding='utf-8') as f:
#                     self.ribbon = f.read()
#                 log_message("Ribbon XML loaded successfully.")
#                 log_message(f"Ribbon XML content preview: {self.ribbon[:200]}...")

#             self.translations = {
#                 1033: {
#                     "tab": "AI Assistant", "group": "AI Tools", "run_prompt": "Run General Prompt",
#                     "analyze_doc": "Analyze Document", "summarize_doc": "Summarize Document",
#                     "create_memo": "Create Memo", "create_minutes": "Create Minutes", "create_cover_letter": "Create Cover Letter",
#                     "prompt_title": "AI Assistant", "prompt_message": "Enter your request (e.g., 'write a report on X'):",
#                     "memo_topic": "Enter the memo topic:", "memo_audience": "Enter the memo's audience:",
#                     "minutes_topic": "Enter the meeting topic:", "minutes_attendees": "Enter attendees (comma-separated):",
#                     "minutes_info": "Enter key discussion points:",
#                     "cover_letter_topic": "Enter the job position:", "cover_letter_audience": "Enter the hiring manager/company:",
#                     "action_cancelled": "Action cancelled.", "contacting_server": "AI Assistant: Contacting server, please wait...",
#                     "connection_error": "\n\nERROR: Could not connect to the backend server. Please ensure the AI Backend is running.\n\n",
#                     "unexpected_error": "\n\nAn unexpected error occurred: {e}\n\n",
#                     "result_header": "\n\n--- AI Assistant Result ---\n", "result_footer": "\n--- End of Result ---\n\n",
#                     "no_active_doc": "No active document found."
#                 }
#             }
            
#             log_message("--- Add-in __init__ completed successfully. ---")
#         except Exception as e:
#             log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}", "ERROR")
#             self.ribbon = ""

#     def GetCustomUI(self, ribbonID):
#         """Return the ribbon XML for WPS Office"""
#         log_message(f"GetCustomUI called with ribbonID: {ribbonID}")
#         return self.ribbon

#     def _get_localized_string(self, key):
#         lang_id = 1033
#         wps_app = get_wps_application()
#         if wps_app:
#             try:
#                 lang_id = wps_app.LanguageSettings.LanguageID(1)
#             except Exception:
#                 pass
#         return self.translations.get(lang_id, self.translations[1033]).get(key, key)

#     def GetTabLabel(self, c): return self._get_localized_string("tab")
#     def GetGroupLabel(self, c): return self._get_localized_string("group")
#     def GetRunPromptLabel(self, c): return self._get_localized_string("run_prompt")
#     def GetAnalyzeDocLabel(self, c): return self._get_localized_string("analyze_doc")
#     def GetSummarizeDocLabel(self, c): return self._get_localized_string("summarize_doc")
#     def GetCreateMemoLabel(self, c): return self._get_localized_string("create_memo")
#     def GetCreateMinutesLabel(self, c): return self._get_localized_string("create_minutes")
#     def GetCreateCoverLetterLabel(self, c): return self._get_localized_string("create_cover_letter")

#     def OnLoadImage(self, imageName):
#         image_path = resource_path(f"{imageName}.png")
#         log_message(f"Attempting to load image: {image_path}")
#         try:
#             img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
#             log_message(f"Successfully loaded image '{imageName}'.")
#             return img_handle
#         except Exception as e:
#             log_message(f"ERROR: Failed to load image '{imageName}': {e}")
#             return None

#     def _call_backend_task(self, endpoint: str, payload: dict):
#         log_message(f"Calling backend endpoint: {endpoint}")
#         try:
#             insert_text_at_cursor(self._get_localized_string("contacting_server"))
#             response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
#             response.raise_for_status()
#             result = response.json().get("result", "")
#             header = self._get_localized_string("result_header")
#             footer = self._get_localized_string("result_footer")
#             insert_text_at_cursor(f"{header}{result}{footer}")
#             log_message(f"Successfully received response from {endpoint}.")
#         except requests.exceptions.ConnectionError:
#             log_message(f"Connection error to {endpoint}")
#             insert_text_at_cursor(self._get_localized_string("connection_error"))
#         except Exception as e:
#             log_message(f"Error calling {endpoint}: {e}")
#             insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))

#     def OnRunPrompt(self, c):
#         log_message("OnRunPrompt called")
#         root = Tk(); root.withdraw()
#         prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("prompt_message"))
#         root.destroy()
#         if not prompt: 
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/process", {"prompt": prompt})).start()

#     def OnAnalyzeDocument(self, c):
#         log_message("OnAnalyzeDocument called")
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()

#     def OnSummarizeDocument(self, c):
#         log_message("OnSummarizeDocument called")
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()

#     def OnCreateMemo(self, c):
#         log_message("OnCreateMemo called")
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("memo_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("memo_audience"))
#         root.destroy()
#         payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
#         threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()

#     def OnCreateMinutes(self, c):
#         log_message("OnCreateMinutes called")
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("minutes_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("minutes_attendees"))
#         info = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                     self._get_localized_string("minutes_info"))
#         root.destroy()
#         payload = {
#             "doc_type": "minutes", "topic": topic, "audience": "Meeting Attendees",
#             "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()],
#             "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]
#         }
#         threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()

#     def OnCreateCoverLetter(self, c):
#         log_message("OnCreateCoverLetter called")
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("cover_letter_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("cover_letter_audience"))
#         root.destroy()
#         payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
#         threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()


# def get_registry_access_flags():
#     """Get the appropriate registry access flags for current architecture"""
#     is_64bit_process = sys.maxsize > 2**32
#     is_64bit_os = 'PROGRAMFILES(X86)' in os.environ
    
#     log_message(f"Process: {'64-bit' if is_64bit_process else '32-bit'}, OS: {'64-bit' if is_64bit_os else '32-bit'}")
    
#     # Use WOW64 flags appropriately
#     if is_64bit_os:
#         if is_64bit_process:
#             # 64-bit process on 64-bit OS - access 64-bit registry view
#             return winreg.KEY_WOW64_64KEY
#         else:
#             # 32-bit process on 64-bit OS - access 32-bit registry view
#             return winreg.KEY_WOW64_32KEY
#     else:
#         # 32-bit OS - no WOW64 flags needed
#         return 0

# def safe_delete_registry_key(root_key, subkey_path, access_flags=0):
#     """Safely delete a registry key, handling all error cases"""
#     try:
#         if access_flags:
#             winreg.DeleteKeyEx(root_key, subkey_path, access_flags, 0)
#         else:
#             winreg.DeleteKey(root_key, subkey_path)
#         log_message(f"Successfully deleted registry key: {subkey_path}")
#         return True
#     except FileNotFoundError:
#         log_message(f"Registry key not found (already deleted): {subkey_path}", "DEBUG")
#         return True
#     except PermissionError as e:
#         log_message(f"Permission denied deleting key {subkey_path}: {e}", "ERROR")
#         return False
#     except Exception as e:
#         log_message(f"Error deleting registry key {subkey_path}: {e}", "ERROR")
#         return False

# def register_server(cls):
#     """Enhanced COM server registration with better error handling"""
#     if not is_admin():
#         log_message("Registration requires administrator privileges. Attempting to elevate...", "WARNING")
#         if not run_as_admin():
#             return False
#         return True  # Script will restart with admin rights
    
#     clsid = cls._reg_clsid_
#     progid = cls._reg_progid_
#     desc = cls._reg_desc_
    
#     # Get both 32-bit and 64-bit access flags
#     access_flags_32 = winreg.KEY_WOW64_32KEY
#     access_flags_64 = winreg.KEY_WOW64_64KEY
#     write_flags_32 = winreg.KEY_WRITE | access_flags_32
#     write_flags_64 = winreg.KEY_WRITE | access_flags_64
    
#     log_message(f"Starting registration for both 32-bit and 64-bit registry views")
    
#     try:
#         # Determine server type and executable path
#         if is_pyinstaller_bundle():
#             server_type = "LocalServer32"
#             executable_path = sys.executable
#             log_message(f"Using LocalServer32 for bundled executable: {executable_path}")
#         else:
#             server_type = "LocalServer32"  # Always use LocalServer32 for Python COM servers
#             executable_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
#             log_message(f"Using LocalServer32 for Python script: {executable_path}")
        
#         # Register in both 32-bit and 64-bit registry views to ensure compatibility
#         for view_name, write_flags in [("32-bit", write_flags_32), ("64-bit", write_flags_64)]:
#             try:
#                 log_message(f"Creating CLSID key in {view_name} registry view: CLSID\\{clsid}")
#                 with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, write_flags) as key:
#                     winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                
#                 log_message(f"Creating {server_type} key in {view_name} view")
#                 with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\{server_type}", 0, write_flags) as server_key:
#                     winreg.SetValueEx(server_key, "", 0, winreg.REG_SZ, executable_path)
                
#                 log_message(f"Creating ProgID key in {view_name} view")
#                 with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, write_flags) as progid_key:
#                     winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
                
#                 # Create the ProgID-to-CLSID mapping
#                 log_message(f"Creating ProgID mapping in {view_name} view: {progid}")
#                 with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, write_flags) as key:
#                     winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
#                     with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
#                         winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
                
#                 log_message(f"Successfully registered COM server in {view_name} registry view")
                
#             except Exception as e:
#                 log_message(f"Failed to register in {view_name} registry view: {e}", "WARNING")
#                 # Continue with other view - don't fail completely
#                 continue
        
#         log_message("Core COM server registered successfully.")
#     except Exception as e:
#         log_message(f"FATAL: Core COM registration failed: {e}\n{traceback.format_exc()}", "ERROR")
#         print("FATAL: Failed to register the COM server. Please run as Administrator.")
#         return False

#     # Create WPS Office add-in entries (try multiple locations)
#     wps_addin_paths = [
#         f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#         f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}",
#         f"Software\\Kingsoft\\WPS Office\\Addins\\{progid}"
#     ]
    
#     registration_success = False
#     for path in wps_addin_paths:
#         try:
#             log_message(f"Attempting to create WPS add-in entry at: HKCU\\{path}")
#             with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
#                 winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
#                 winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
#                 winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
#                 winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
#                 # Add CLSID to the WPS add-in entry
#                 winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
#             log_message(f"SUCCESS: WPS add-in entry created at HKCU\\{path}")
#             registration_success = True
#         except Exception as e:
#             log_message(f"Failed to create add-in entry at {path}: {e}", "WARNING")
#             continue
    
#     if registration_success:
#         print("SUCCESS: Add-in registered successfully!")
#         print("Please restart WPS Office to see the add-in.")
#         print("\nTo verify registration, run these commands:")
#         print(f'reg query "HKCR\\CLSID\\{clsid}" /s')
#         print(f'reg query "HKCU\\Software\\Kingsoft\\Office\\Addins\\{progid}" /s')
#         log_message("Registration completed successfully.")
#         return True
#     else:
#         log_message("FAILED: Could not create WPS add-in entry in any known location.", "ERROR")
#         print("FAILED: Could not register the WPS Office add-in entry.")
#         return False

# def unregister_server(cls):
#     """Enhanced COM server unregistration with proper error handling"""
#     if not is_admin():
#         log_message("Unregistration requires administrator privileges. Attempting to elevate...", "WARNING")
#         if not run_as_admin():
#             return False
#         return True  # Script will restart with admin rights
    
#     clsid = cls._reg_clsid_
#     progid = cls._reg_progid_
    
#     # Use both 32-bit and 64-bit access flags for thorough cleanup
#     access_flags_32 = winreg.KEY_WOW64_32KEY
#     access_flags_64 = winreg.KEY_WOW64_64KEY
    
#     log_message("Starting unregistration from both 32-bit and 64-bit registry views")
    
#     # Remove WPS Office entries (try all possible locations)
#     wps_addin_paths = [
#         f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#         f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}",
#         f"Software\\Kingsoft\\WPS Office\\Addins\\{progid}"
#     ]
    
#     for path in wps_addin_paths:
#         try:
#             winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
#             log_message(f"Removed WPS Add-in entry at: HKCU\\{path}")
#         except FileNotFoundError:
#             log_message(f"WPS Add-in entry not found at: HKCU\\{path}", "DEBUG")
#         except Exception as e:
#             log_message(f"Error removing WPS add-in entry at {path}: {e}", "WARNING")

#     # Remove COM server entries from both registry views
#     for view_name, access_flags in [("32-bit", access_flags_32), ("64-bit", access_flags_64)]:
#         log_message(f"Cleaning up {view_name} registry view")
        
#         # Remove COM server entries in correct order
#         server_types = ["LocalServer32", "InprocServer32"]  # Try both
        
#         for server_type in server_types:
#             safe_delete_registry_key(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\{server_type}", access_flags)
        
#         # Remove other CLSID subkeys
#         safe_delete_registry_key(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", access_flags)
#         safe_delete_registry_key(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", access_flags)
        
#         # Remove ProgID entries
#         safe_delete_registry_key(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID", access_flags)
#         safe_delete_registry_key(winreg.HKEY_CLASSES_ROOT, progid, access_flags)
    
#     log_message("Unregistration completed.")
#     print("Unregistration complete.")
#     return True

# def run_com_server():
#     """Run the COM server with enhanced error handling"""
#     log_message("Preparing to start COM server...")
#     try:
#         pythoncom.CoInitialize()
#         log_message("COM initialized successfully.")
        
#         factory = pythoncom.MakePyFactory(WPSAddin)
#         clsid = WPSAddin._reg_clsid_
        
#         log_message(f"Registering class object with CLSID: {clsid}")
#         pythoncom.CoRegisterClassObject(
#             clsid, 
#             factory, 
#             pythoncom.CLSCTX_LOCAL_SERVER, 
#             pythoncom.REGCLS_MULTIPLEUSE
#         )
        
#         log_message("COM Class Object registered successfully. Starting message pump...")
#         print("COM Server is running. Press Ctrl+C to stop.")
        
#         # Message pump
#         pythoncom.PumpMessages()
        
#     except KeyboardInterrupt:
#         log_message("COM server stopped by user.")
#         print("\nCOM Server stopped.")
#     except Exception as e:
#         log_message(f"FATAL ERROR while running COM server: {e}\n{traceback.format_exc()}", "ERROR")
#         print(f"FATAL ERROR: {e}")
#     finally:
#         try:
#             pythoncom.CoUninitialize()
#             log_message("COM uninitialized successfully.")
#         except:
#             pass

# if __name__ == '__main__':
#     log_message("Executing main block.")
    
#     # Main execution logic
#     if len(sys.argv) > 1:
#         if sys.argv[1].lower() == '/regserver':
#             log_message("Command: /regserver")
#             success = register_server(WPSAddin)
#             sys.exit(0 if success else 1)
#         elif sys.argv[1].lower() == '/unregserver':
#             log_message("Command: /unregserver")
#             success = unregister_server(WPSAddin)
#             sys.exit(0 if success else 1)
#         else:
#             log_message(f"Command: {sys.argv[1]} - Starting COM Server.")
#             run_com_server()
#     else:
#         log_message("Command: No arguments - Starting COM Server.")
#         run_com_server()

# """
# This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
# Fixed version addressing common WPS Office add-in loading issues and PyInstaller compatibility.
# Enhanced with better 32/64-bit registry handling and comprehensive logging.
# """
# import datetime
# import traceback
# import os
# import sys
# import threading
# import requests
# import win32com.client
# import win32com.server.register
# import win32api
# import winreg
# import pythoncom
# import logging
# from tkinter import simpledialog, Tk
# import ctypes
# from ctypes import wintypes


# def is_admin():
#     """Check if the current process has administrator privileges."""
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False

# def run_as_admin():
#     """Re-run the current script with administrator privileges."""
#     if is_admin():
#         return True
#     else:
#         try:
#             # Re-run the program with admin rights
#             ctypes.windll.shell32.ShellExecuteW(
#                 None, "runas", sys.executable, " ".join(sys.argv), None, 1
#             )
#             return False
#         except:
#             print("Failed to elevate privileges. Please run as Administrator.")
#             return False

# # Enhanced file-based logging 
# try:
#     base_dir = os.path.dirname(os.path.abspath(__file__))
#     log_dir = os.path.join(base_dir, "logs")
#     os.makedirs(log_dir, exist_ok=True)
#     log_file = os.path.join(log_dir, "addin_debug.log")

#     # Configure logging with more detailed format
#     logging.basicConfig(
#         filename=log_file,
#         level=logging.DEBUG,  # Changed to DEBUG for more details
#         format='[%(asctime)s] - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
#         datefmt='%Y-%m-%d %H:%M:%S'
#     )

#     # Also log to console
#     console_handler = logging.StreamHandler()
#     console_handler.setLevel(logging.INFO)
#     console_formatter = logging.Formatter('[%(levelname)s] %(message)s')
#     console_handler.setFormatter(console_formatter)
#     logging.getLogger().addHandler(console_handler)

#     def log_message(message, level="INFO"):
#         if level.upper() == "DEBUG":
#             logging.debug(message)
#         elif level.upper() == "WARNING":
#             logging.warning(message)
#         elif level.upper() == "ERROR":
#             logging.error(message)
#         else:
#             logging.info(message)

#     log_message("--- SCRIPT EXECUTION STARTED ---")
#     log_message(f"Python Version: {sys.version}")
#     log_message(f"Executable Path: {sys.executable}")
#     log_message(f"Command Line Arguments: {sys.argv}")
#     log_message(f"Process Architecture: {'64-bit' if sys.maxsize > 2**32 else '32-bit'}")
#     log_message(f"Admin Rights: {is_admin()}")

# except Exception as e:
#     print(f"FATAL: Could not set up logging. Error: {e}")
    
    
# # Configuration - IP address
# BACKEND_URL = "http://127.0.0.1:8000"

# # Consistent naming
# WPS_ADDIN_ENTRY_NAME = "WPSAIAddin.Connect"


# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller bundling """
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)

# def is_pyinstaller_bundle():
#     """Check if running as PyInstaller bundle"""
#     return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

# def get_wps_application():
#     """Gets the running WPS Writer Application object."""
#     try:
#         return win32com.client.GetActiveObject("kwps.Application")
#     except Exception as e:
#         log_message(f"Error getting WPS Application object: {e}")
#         return None

# def insert_text_at_cursor(text):
#     """Inserts text into the active document at the current cursor position."""
#     wps_app = get_wps_application()
#     if wps_app and wps_app.Documents.Count > 0:
#         try:
#             wps_app.Selection.TypeText(Text=text)
#             log_message("Text successfully inserted into active WPS document.")
#         except Exception as e:
#             log_message(f"Error inserting text into WPS document: {e}\n{traceback.format_exc()}")
#     else:
#         log_message("Warning: Could not find an active WPS document to insert text into.")

# class WPSAddin:
#     # Proper COM registration attributes
#     _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
#     _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}"
#     _reg_desc_ = "AI Office Automation"
#     _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
#     _reg_class_spec_ = "addin_client5.WPSAddin"

#     _public_methods_ = [
#         'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
#         'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
#         'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
#         'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter',
#         'GetCustomUI'
#     ]
#     _public_attrs_ = ['ribbon']

#     def __init__(self):
#         log_message("--- WPSAdd-in __init__ started ---")
        
#         try:
#             ribbon_path = resource_path('ribbon.xml')
#             log_message(f"Attempting to load ribbon from: {ribbon_path}")
            
#             if not os.path.exists(ribbon_path):
#                 log_message(f"FATAL: Ribbon XML file does NOT exist at the path.", "ERROR")
#                 # Create a basic ribbon XML if file doesn't exist
#                 self.ribbon = '''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
#                     <ribbon>
#                         <tabs>
#                             <tab id="AITab" label="AI Assistant">
#                                 <group id="AIGroup" label="AI Tools">
#                                     <button id="RunPrompt" label="Run Prompt" onAction="OnRunPrompt"/>
#                                     <button id="AnalyzeDoc" label="Analyze Document" onAction="OnAnalyzeDocument"/>
#                                     <button id="SummarizeDoc" label="Summarize Document" onAction="OnSummarizeDocument"/>
#                                 </group>
#                             </tab>
#                         </tabs>
#                     </ribbon>
#                 </customUI>'''
#                 log_message("Using default ribbon XML")
#             else:
#                 with open(ribbon_path, 'r', encoding='utf-8') as f:
#                     self.ribbon = f.read()
#                 log_message("Ribbon XML loaded successfully.")
#                 log_message(f"Ribbon XML content preview: {self.ribbon[:200]}...")

#             self.translations = {
#                 1033: {
#                     "tab": "AI Assistant", "group": "AI Tools", "run_prompt": "Run General Prompt",
#                     "analyze_doc": "Analyze Document", "summarize_doc": "Summarize Document",
#                     "create_memo": "Create Memo", "create_minutes": "Create Minutes", "create_cover_letter": "Create Cover Letter",
#                     "prompt_title": "AI Assistant", "prompt_message": "Enter your request (e.g., 'write a report on X'):",
#                     "memo_topic": "Enter the memo topic:", "memo_audience": "Enter the memo's audience:",
#                     "minutes_topic": "Enter the meeting topic:", "minutes_attendees": "Enter attendees (comma-separated):",
#                     "minutes_info": "Enter key discussion points:",
#                     "cover_letter_topic": "Enter the job position:", "cover_letter_audience": "Enter the hiring manager/company:",
#                     "action_cancelled": "Action cancelled.", "contacting_server": "AI Assistant: Contacting server, please wait...",
#                     "connection_error": "\n\nERROR: Could not connect to the backend server. Please ensure the AI Backend is running.\n\n",
#                     "unexpected_error": "\n\nAn unexpected error occurred: {e}\n\n",
#                     "result_header": "\n\n--- AI Assistant Result ---\n", "result_footer": "\n--- End of Result ---\n\n",
#                     "no_active_doc": "No active document found."
#                 }
#             }
            
#             log_message("--- Add-in __init__ completed successfully. ---")
#         except Exception as e:
#             log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}", "ERROR")
#             self.ribbon = ""

#     def GetCustomUI(self, ribbonID):
#         """Return the ribbon XML for WPS Office"""
#         log_message(f"GetCustomUI called with ribbonID: {ribbonID}")
#         return self.ribbon

#     def _get_localized_string(self, key):
#         lang_id = 1033
#         wps_app = get_wps_application()
#         if wps_app:
#             try:
#                 lang_id = wps_app.LanguageSettings.LanguageID(1)
#             except Exception:
#                 pass
#         return self.translations.get(lang_id, self.translations[1033]).get(key, key)

#     def GetTabLabel(self, c): return self._get_localized_string("tab")
#     def GetGroupLabel(self, c): return self._get_localized_string("group")
#     def GetRunPromptLabel(self, c): return self._get_localized_string("run_prompt")
#     def GetAnalyzeDocLabel(self, c): return self._get_localized_string("analyze_doc")
#     def GetSummarizeDocLabel(self, c): return self._get_localized_string("summarize_doc")
#     def GetCreateMemoLabel(self, c): return self._get_localized_string("create_memo")
#     def GetCreateMinutesLabel(self, c): return self._get_localized_string("create_minutes")
#     def GetCreateCoverLetterLabel(self, c): return self._get_localized_string("cover_letter")

#     def OnLoadImage(self, imageName):
#         image_path = resource_path(f"{imageName}.png")
#         log_message(f"Attempting to load image: {image_path}")
#         try:
#             img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
#             log_message(f"Successfully loaded image '{imageName}'.")
#             return img_handle
#         except Exception as e:
#             log_message(f"ERROR: Failed to load image '{imageName}': {e}")
#             return None

#     def _call_backend_task(self, endpoint: str, payload: dict, retries=3):
#         """Enhanced backend task with retry logic"""
#         log_message(f"Calling backend endpoint: {endpoint}")
        
#         for attempt in range(retries):
#             try:
#                 if attempt == 0:  # Only show "contacting server" on first attempt
#                     insert_text_at_cursor(self._get_localized_string("contacting_server"))
                
#                 response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
#                 response.raise_for_status()
#                 result = response.json().get("result", "")
#                 header = self._get_localized_string("result_header")
#                 footer = self._get_localized_string("result_footer")
#                 insert_text_at_cursor(f"{header}{result}{footer}")
#                 log_message(f"Successfully received response from {endpoint}.")
#                 return  # Success, exit retry loop
                
#             except requests.exceptions.ConnectionError as e:
#                 log_message(f"Connection error to {endpoint}, attempt {attempt + 1}/{retries}: {e}")
#                 if attempt == retries - 1:  # Last attempt
#                     insert_text_at_cursor(self._get_localized_string("connection_error"))
#                 else:
#                     import time
#                     time.sleep(2 ** attempt)  # Exponential backoff
#             except Exception as e:
#                 log_message(f"Error calling {endpoint}: {e}")
#                 insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))
#                 break  # Don't retry for non-connection errors

#     def OnRunPrompt(self, c):
#         log_message("OnRunPrompt called")
#         root = Tk(); root.withdraw()
#         try:
#             prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("prompt_message"))
#         finally:
#             root.destroy()
            
#         if not prompt: 
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/process", {"prompt": prompt}), daemon=True).start()

#     def OnAnalyzeDocument(self, c):
#         log_message("OnAnalyzeDocument called")
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/analyze", {"content": content, "prompt": "Analyze the document content."}), 
#                         daemon=True).start()

#     def OnSummarizeDocument(self, c):
#         log_message("OnSummarizeDocument called")
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/summarize", {"content": content, "prompt": "Summarize the document content."}), 
#                         daemon=True).start()

#     def OnCreateMemo(self, c):
#         log_message("OnCreateMemo called")
#         root = Tk(); root.withdraw()
#         try:
#             topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("memo_topic"))
#             if not topic:
#                 return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#             audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("memo_audience"))
#         finally:
#             root.destroy()
            
#         payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
#         threading.Thread(target=self._call_backend_task, args=("/create_memo", payload), daemon=True).start()

#     def OnCreateMinutes(self, c):
#         log_message("OnCreateMinutes called")
#         root = Tk(); root.withdraw()
#         try:
#             topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("minutes_topic"))
#             if not topic:
#                 return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#             attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                                 self._get_localized_string("minutes_attendees"))
#             info = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("minutes_info"))
#         finally:
#             root.destroy()
            
#         payload = {
#             "doc_type": "minutes", "topic": topic, "audience": "Meeting Attendees",
#             "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()],
#             "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]
#         }
#         threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload), daemon=True).start()

#     def OnCreateCoverLetter(self, c):
#         log_message("OnCreateCoverLetter called")
#         root = Tk(); root.withdraw()
#         try:
#             topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("cover_letter_topic"))
#             if not topic:
#                 return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#             audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                             self._get_localized_string("cover_letter_audience"))
#         finally:
#             root.destroy()
            
#         payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
#         threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload), daemon=True).start()


#     def get_registry_access_flags():
#         """Get the appropriate registry access flags for current architecture"""
#         is_64bit_process = sys.maxsize > 2**32
#         is_64bit_os = 'PROGRAMFILES(X86)' in os.environ
        
#         log_message(f"Process: {'64-bit' if is_64bit_process else '32-bit'}, OS: {'64-bit' if is_64bit_os else '32-bit'}")
        
#         # Use WOW64 flags appropriately
#         if is_64bit_os:
#             if is_64bit_process:
#                 # 64-bit process on 64-bit OS - access 64-bit registry view
#                 return winreg.KEY_WOW64_64KEY
#             else:
#                 # 32-bit process on 64-bit OS - access 32-bit registry view
#                 return winreg.KEY_WOW64_32KEY
#         else:
#             # 32-bit OS - no WOW64 flags needed
#             return 0

#     def safe_delete_registry_key(root_key, subkey_path, access_flags=0):
#         """Safely delete a registry key, handling all error cases"""
#         try:
#             if access_flags:
#                 winreg.DeleteKeyEx(root_key, subkey_path, access_flags, 0)
#             else:
#                 winreg.DeleteKey(root_key, subkey_path)
#             log_message(f"Successfully deleted registry key: {subkey_path}")
#             return True
#         except FileNotFoundError:
#             log_message(f"Registry key not found (already deleted): {subkey_path}", "DEBUG")
#             return True
#         except PermissionError as e:
#             log_message(f"Permission denied deleting key {subkey_path}: {e}", "ERROR")
#             return False
#         except Exception as e:
#             log_message(f"Error deleting registry key {subkey_path}: {e}", "ERROR")
#             return False

#     def register_server():
#         """Register the COM server manually using registry operations"""
#         if not is_admin():
#             log_message("Registration requires administrator privileges. Attempting to elevate...", "WARNING")
#             if not run_as_admin():
#                 return False
#             return True
        
#         clsid = WPSAddin._reg_clsid_
#         progid = WPSAddin._reg_progid_
#         desc = WPSAddin._reg_desc_
        
#         # Get both 32-bit and 64-bit access flags
#         access_flags_32 = winreg.KEY_WOW64_32KEY
#         access_flags_64 = winreg.KEY_WOW64_64KEY
#         write_flags_32 = winreg.KEY_WRITE | access_flags_32
#         write_flags_64 = winreg.KEY_WRITE | access_flags_64
        
#         log_message(f"Starting registration for both 32-bit and 64-bit registry views")
        
#         try:
#             # Determine server type and executable path
#             if is_pyinstaller_bundle():
#                 server_type = "LocalServer32"
#                 executable_path = sys.executable
#                 log_message(f"Using LocalServer32 for bundled executable: {executable_path}")
#             else:
#                 server_type = "LocalServer32"
#                 executable_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
#                 log_message(f"Using LocalServer32 for Python script: {executable_path}")
            
#             # Register in both 32-bit and 64-bit registry views to ensure compatibility
#             for view_name, write_flags in [("32-bit", write_flags_32), ("64-bit", write_flags_64)]:
#                 try:
#                     log_message(f"Creating CLSID key in {view_name} registry view: CLSID\\{clsid}")
#                     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, write_flags) as key:
#                         winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                    
#                     log_message(f"Creating {server_type} key in {view_name} view")
#                     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\{server_type}", 0, write_flags) as server_key:
#                         winreg.SetValueEx(server_key, "", 0, winreg.REG_SZ, executable_path)
                    
#                     log_message(f"Creating ProgID key in {view_name} view")
#                     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, write_flags) as progid_key:
#                         winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
                    
#                     # Create the ProgID-to-CLSID mapping
#                     log_message(f"Creating ProgID mapping in {view_name} view: {progid}")
#                     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, write_flags) as key:
#                         winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
#                         with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
#                             winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
                    
#                     log_message(f"Successfully registered COM server in {view_name} registry view")
                    
#                 except Exception as e:
#                     log_message(f"Failed to register in {view_name} registry view: {e}", "WARNING")
#                     continue
            
#             log_message("Core COM server registered successfully.")
#         except Exception as e:
#             log_message(f"FATAL: Core COM registration failed: {e}\n{traceback.format_exc()}", "ERROR")
#             print("FATAL: Failed to register the COM server. Please run as Administrator.")
#             return False
        
#     # Register WPS add-in entry
#         clsid = WPSAddin._reg_clsid_
#         progid = WPSAddin._reg_progid_
#         desc = WPSAddin._reg_desc_
            
#         wps_addin_paths = [
#             f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#             f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}",
#             f"Software\\Kingsoft\\WPS Office\\Addins\\{progid}"
#         ]
            
#         registration_success = False
#         for path in wps_addin_paths:
#             try:
#                 log_message(f"Attempting to create WPS add-in entry at: HKCU\\{path}")
#                 with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
#                     winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
#                     winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
#                     winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
#                     winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
#                     winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
#                     log_message(f"SUCCESS: WPS add-in entry created at HKCU\\{path}")
#                 registration_success = True
#                 break  # Success, no need to try other paths
#             except Exception as e:
#                 log_message(f"Failed to create add-in entry at {path}: {e}", "WARNING")
#                 continue
            
#         if registration_success:
#             print("SUCCESS: Add-in registered successfully!")
#             print("Please restart WPS Office to see the add-in.")
#             return True
#         else:
#             log_message("FAILED: Could not create WPS add-in entry in any known location.", "ERROR")
#             return False
                
#     except Exception as e:
#         log_message(f"FATAL: COM registration failed: {e}\n{traceback.format_exc()}", "ERROR")
#         print(f"FATAL: Failed to register the COM server: {e}")
#         return False

# def unregister_server():
#     """Unregister the COM server"""
#     if not is_admin():
#         log_message("Unregistration requires administrator privileges. Attempting to elevate...", "WARNING")
#         if not run_as_admin():
#             return False
#         return True
    
#     try:
#         import win32com.server.register
#         # Fixed: Pass the class, not just the CLSID
#         win32com.server.register.UnregisterServer(WPSAddin)
#         log_message("COM server unregistered successfully")
        
#         # Remove WPS add-in entries
#         progid = WPSAddin._reg_progid_
#         wps_addin_paths = [
#             f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#             f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}",
#             f"Software\\Kingsoft\\WPS Office\\Addins\\{progid}"
#         ]
        
#         for path in wps_addin_paths:
#             try:
#                 winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
#                 log_message(f"Removed WPS Add-in entry at: HKCU\\{path}")
#             except FileNotFoundError:
#                 log_message(f"WPS Add-in entry not found at: HKCU\\{path}", "DEBUG")
#             except Exception as e:
#                 log_message(f"Error removing WPS add-in entry at {path}: {e}", "WARNING")
        
#         print("Unregistration complete.")
#         return True
        
#     except Exception as e:
#         log_message(f"Error during unregistration: {e}", "ERROR")
#         print(f"Error during unregistration: {e}")
#         return False

# def run_com_server():
#     """Run the COM server with proper factory creation"""
#     log_message("Preparing to start COM server...")
#     try:
#         pythoncom.CoInitialize()
#         log_message("COM initialized successfully.")
        
#         # Fixed: Use win32com.server.factory for proper factory creation
#         from win32com.server.factory import Factory
#         factory = Factory(pythoncom.CLSID(WPSAddin._reg_clsid_), WPSAddin)
        
#         log_message(f"Registering class object with CLSID: {WPSAddin._reg_clsid_}")
#         reg_handle = pythoncom.CoRegisterClassObject(
#             WPSAddin._reg_clsid_, 
#             factory, 
#             pythoncom.CLSCTX_LOCAL_SERVER, 
#             pythoncom.REGCLS_MULTIPLEUSE
#         )
        
#         log_message("COM Class Object registered successfully. Starting message pump...")
#         print("COM Server is running. Press Ctrl+C to stop.")
        
#         # Message pump
#         pythoncom.PumpMessages()
        
#     except KeyboardInterrupt:
#         log_message("COM server stopped by user.")
#         print("\nCOM Server stopped.")
#     except Exception as e:
#         log_message(f"FATAL ERROR while running COM server: {e}\n{traceback.format_exc()}", "ERROR")
#         print(f"FATAL ERROR: {e}")
        
#         # Additional debugging info
#         log_message(f"WPSAddin._reg_clsid_ = {WPSAddin._reg_clsid_}", "DEBUG")
#         log_message(f"Type of _reg_clsid_: {type(WPSAddin._reg_clsid_)}", "DEBUG")
        
#     finally:
#         try:
#             pythoncom.CoUninitialize()
#             log_message("COM uninitialized successfully.")
#         except:
#             pass

# if __name__ == '__main__':
#     log_message("Executing main block.")
    
#     # Main execution logic
#     if len(sys.argv) > 1:
#         if sys.argv[1].lower() == '/regserver':
#             log_message("Command: /regserver")
#             success = register_server()
#             sys.exit(0 if success else 1)
#         elif sys.argv[1].lower() == '/unregserver':
#             log_message("Command: /unregserver")
#             success = unregister_server()
#             sys.exit(0 if success else 1)
#         else:
#             log_message(f"Command: {sys.argv[1]} - Starting COM Server.")
#             run_com_server()
#     else:
#         log_message("Command: No arguments - Starting COM Server.")
#         run_com_server()
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
import subprocess # For admin elevation
import ctypes     # For admin check

# --- Logging Setup (from your original code, kept for context) ---
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
    def log_message(message, level="INFO"):
        if level == "INFO": logging.info(message)
        elif level == "WARNING": logging.warning(message)
        elif level == "ERROR": logging.error(message)
        elif level == "DEBUG": logging.debug(message)
        elif level == "CRITICAL": logging.critical(message)
        print(f"LOG: {message}") # Also print to console
    log_message("--- SCRIPT EXECUTION STARTED ---")
    log_message(f"Python Version: {sys.version}")
    log_message(f"Executable Path: {sys.executable}")
    log_message(f"Command Line Arguments: {sys.argv}")
except Exception as e:
    print(f"FATAL: Could not set up logging. Error: {e}")

# --- Configuration (from your original code, kept for context) ---
BACKEND_URL = "http://127.0.0.1:8000"
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
        log_message(f"Error getting WPS Application object: {e}", "WARNING")
        return None

def insert_text_at_cursor(text):
    """Inserts text into the active document at the current cursor position."""
    wps_app = get_wps_application()
    if wps_app and wps_app.Documents.Count > 0:
        try:
            wps_app.Selection.TypeText(Text=text)
            log_message("Text successfully inserted into active WPS document.")
        except Exception as e:
            log_message(f"Error inserting text into WPS document: {e}\n{traceback.format_exc()}", "ERROR")
    else:
        log_message("Warning: Could not find an active WPS document to insert text into.", "WARNING")

# --- Dummy WPSAddin Class (for context) ---
class WPSAddin:
    _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}"
    _reg_desc_ = "AI Office Automation"
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME  
    _reg_class_spec_ = __name__ + ".WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel',
        'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter',
        'GetCustomUI'
    ]
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("--- Add-in __init__ started ---")
        try:
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Attempting to load ribbon from: {ribbon_path}")
            if not os.path.exists(ribbon_path):
                log_message(f"FATAL: Ribbon XML file does NOT exist at the path.", "CRITICAL")
                raise FileNotFoundError(f"Ribbon XML not found at {ribbon_path}")
            with open(ribbon_path, 'r', encoding='utf-8') as f:
                self.ribbon = f.read()
            log_message("Ribbon XML loaded successfully.")
            self.translations = {1033: {"tab": "AI Assistant", "group": "AI Tools", "run_prompt": "Run General Prompt", "analyze_doc": "Analyze Document", "summarize_doc": "Summarize Document", "create_memo": "Create Memo", "create_minutes": "Create Minutes", "create_cover_letter": "Create Cover Letter", "prompt_title": "AI Assistant", "prompt_message": "Enter your request (e.g., 'write a report on X'):", "memo_topic": "Enter the memo topic:", "memo_audience": "Enter the memo's audience:", "minutes_topic": "Enter the meeting topic:", "minutes_attendees": "Enter attendees (comma-separated):", "minutes_info": "Enter key discussion points:", "cover_letter_topic": "Enter the job position:", "cover_letter_audience": "Enter the hiring manager/company:", "action_cancelled": "Action cancelled.", "contacting_server": "AI Assistant: Contacting server, please wait...", "connection_error": "\n\nERROR: Could not connect to the backend server. Please ensure the AI Backend is running.\n\n", "unexpected_error": "\n\nAn unexpected error occurred: {e}\n\n", "result_header": "\n\n--- AI Assistant Result ---\n", "result_footer": "\n--- End of Result ---\n\n", "no_active_doc": "No active document found."}}
            log_message("--- Add-in __init__ completed successfully. ---")
        except Exception as e:
            log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}", "CRITICAL")
            self.ribbon = ""
            raise # Re-raise to indicate severe initialization failure

    def GetCustomUI(self, ribbon_id):
        log_message(f"=== WPS GETCUSTOMUI METHOD CALLED (ID: {ribbon_id}) ===")
        return self.ribbon if hasattr(self, 'ribbon') else ""
    def _get_localized_string(self, key):
        lang_id = 1033
        wps_app = get_wps_application()
        if wps_app:
            try: lang_id = wps_app.LanguageSettings.LanguageID(1)
            except: pass
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
        try: return win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
        except Exception as e: log_message(f"ERROR: Failed to load image '{imageName}': {e}", "ERROR")
        return None
    def _call_backend_task(self, endpoint: str, payload: dict):
        try:
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300)
            response.raise_for_status()
            result = response.json().get("result", "")
            header = self._get_localized_string("result_header")
            footer = self._get_localized_string("result_footer")
            insert_text_at_cursor(f"{header}{result}{footer}")
            log_message(f"Successfully received response from {endpoint}.")
        except requests.exceptions.ConnectionError: insert_text_at_cursor(self._get_localized_string("connection_error"))
        except Exception as e: insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))
    def OnRunPrompt(self, c):
        root = Tk(); root.withdraw()
        prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("prompt_message"))
        root.destroy()
        if not prompt: return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        threading.Thread(target=self._call_backend_task, args=("/process", {"prompt": prompt})).start()
    def OnAnalyzeDocument(self, c):
        wps_app = get_wps_application()
        if not wps_app or wps_app.Documents.Count == 0: return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
        content = wps_app.ActiveDocument.Content.Text
        threading.Thread(target=self._call_backend_task, args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()
    def OnSummarizeDocument(self, c):
        wps_app = get_wps_application()
        if not wps_app or wps_app.Documents.Count == 0: return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
        content = wps_app.ActiveDocument.Content.Text
        threading.Thread(target=self._call_backend_task, args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()
    def OnCreateMemo(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_topic"))
        if not topic: root.destroy(); return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_audience"))
        root.destroy()
        payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
        threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()
    def OnCreateMinutes(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_topic"))
        if not topic: root.destroy(); return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_attendees"))
        info = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_info"))
        root.destroy()
        payload = {"doc_type": "minutes", "topic": topic, "audience": "Meeting Attendees", "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()], "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]}
        threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()
    def OnCreateCoverLetter(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_topic"))
        if not topic: root.destroy(); return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_audience"))
        root.destroy()
        payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
        threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()


# --- Helper functions for admin checks ---
def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin(command=None):
    """Re-run the current script or a given command with administrator privileges."""
    if command is None:
        command = [sys.executable] + sys.argv
    
    # Ensure the command is a list of strings
    if not isinstance(command, list):
        command = [command]

    try:
        # Use shell execute to elevate
        # 'runas' verb requests elevation
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", command[0], " ".join(command[1:]), None, 1
        )
        # ShellExecuteW returns > 32 for success.
        return result > 32
    except Exception as e:
        log_message(f"Failed to elevate process: {e}", "ERROR")
        return False

def is_pyinstaller_bundle():
    """Check if running as PyInstaller bundle"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


# --- Core Functions (Updated) ---
def register_server():
    """Register the COM server manually using registry operations."""
    if not is_admin():
        log_message("Registration requires administrator privileges. Attempting to elevate...", "WARNING")
        if run_as_admin():
            log_message("Successfully launched elevated process. Original process exiting.", "INFO")
            sys.exit(0) # Exit the non-elevated process
        else:
            log_message("Failed to elevate privileges. Registration aborted.", "ERROR")
            print("ERROR: Failed to obtain administrator privileges for registration.")
            return False
    
    clsid = WPSAddin._reg_clsid_
    progid = WPSAddin._reg_progid_
    desc = WPSAddin._reg_desc_
    
    # Get relevant access flags
    write_access_key = winreg.KEY_WRITE
    
    # Determine executable path for LocalServer32
    if is_pyinstaller_bundle():
        executable_path = f'"{sys.executable}" /embedding'
    else:
        # For a Python script, we need to launch python.exe with the script path and /embedding
        executable_path = f'"{sys.executable}" "{os.path.abspath(__file__)}" /embedding'
    
    log_message(f"Registering COM server for executable: {executable_path}")
    
    # Register in both 32-bit and 64-bit registry views to ensure compatibility
    # HKEY_CLASSES_ROOT automatically merges HKLM\Software\Classes and HKCU\Software\Classes
    # For LocalServer32 for EXEs/scripts, it's generally stored under the main CLSID path,
    # and Windows redirection handles WOW64 if necessary based on the executable.
    # We will ensure both explicit WOW64 views are targeted for robustness.
    
    views_to_register = []
    if sys.maxsize > 2**32: # 64-bit Python
        views_to_register.append((winreg.KEY_WOW64_64KEY, "64-bit"))
        views_to_register.append((winreg.KEY_WOW64_32KEY, "32-bit")) # Try to register 32-bit view if WPS is 32-bit
    else: # 32-bit Python
        views_to_register.append((winreg.KEY_WOW64_32KEY, "32-bit")) # Primarily 32-bit view

    com_registration_successful = False
    for view_flag, view_name in views_to_register:
        try:
            log_message(f"Attempting COM registration in {view_name} view (flag: {view_flag})...")
            # Create CLSID key
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, write_access_key | view_flag) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                
            # Create LocalServer32 key
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32", 0, write_access_key | view_flag) as server_key:
                winreg.SetValueEx(server_key, "", 0, winreg.REG_SZ, executable_path)
            
            # Create ProgID key under CLSID
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, write_access_key | view_flag) as progid_key:
                winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
            
            # Create ProgID mapping (ProgID -> CLSID)
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, write_access_key | view_flag) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                with winreg.CreateKeyEx(key, "CLSID", 0, write_access_key) as clsid_key: # CLSID under ProgID does not need view_flag explicitly here
                    winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
            
            log_message(f"Successfully registered COM server in {view_name} registry view.", "INFO")
            com_registration_successful = True
            # For simplicity, if one view succeeds, we consider core COM registration successful.
            # More complex scenarios might require both.
            # break 
        except Exception as e:
            log_message(f"Failed to register COM server in {view_name} view: {e}", "WARNING")
    
    if not com_registration_successful:
        log_message("FAILED: Core COM server registration failed in all attempted views.", "ERROR")
        print("FAILED: Core COM server registration failed.")
        return False

    # --- Register WPS add-in entry (separate step) ---
    log_message("Attempting to register WPS add-in entry in HKCU...")
    wps_addin_paths = [
        f"Software\\Kingsoft\\Office\\Addins\\{progid}",
        f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}", # Common for older/specific versions
        f"Software\\WPS\\Office\\Addins\\{progid}",
        f"Software\\WPS Office\\Addins\\{progid}"
    ]
    
    wps_addin_registration_successful = False
    for path in wps_addin_paths:
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
                winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
                winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
            log_message(f"SUCCESS: WPS add-in entry created at HKCU\\{path}", "INFO")
            wps_addin_registration_successful = True
            break # Success, no need to try other paths
        except Exception as e:
            log_message(f"Failed to create WPS add-in entry at HKCU\\{path}: {e}", "WARNING")
            continue
    
    if not wps_addin_registration_successful:
        log_message("FAILED: Could not create WPS add-in entry in any known location.", "ERROR")
        print("FAILED: Could not create WPS Office add-in entry.")
        return False
    
    print("SUCCESS: Add-in registered successfully!")
    print("Please restart WPS Office to see the add-in.")
    return True

def unregister_server():
    """Unregister the COM server and WPS add-in entries."""
    if not is_admin():
        log_message("Unregistration requires administrator privileges. Attempting to elevate...", "WARNING")
        if run_as_admin():
            log_message("Successfully launched elevated process for unregistration. Original process exiting.", "INFO")
            sys.exit(0) # Exit the non-elevated process
        else:
            log_message("Failed to elevate privileges. Unregistration aborted.", "ERROR")
            print("ERROR: Failed to obtain administrator privileges for unregistration.")
            return False

    clsid = WPSAddin._reg_clsid_
    progid = WPSAddin._reg_progid_

    unregistration_successful = True

    # --- Remove core COM server entries from HKEY_CLASSES_ROOT ---
    log_message("Attempting to remove core COM server entries from HKEY_CLASSES_ROOT...")
    
    # Paths to attempt deleting
    com_paths_to_remove = [
        f"CLSID\\{clsid}\\LocalServer32",
        f"CLSID\\{clsid}\\ProgID",
        f"CLSID\\{clsid}",
        f"{progid}\\CLSID",
        f"{progid}"
    ]

    # Try deleting from both 64-bit and 32-bit views
    views_to_delete = []
    if sys.maxsize > 2**32: # 64-bit Python
        views_to_delete.append((winreg.KEY_WOW64_64KEY, "64-bit"))
        views_to_delete.append((winreg.KEY_WOW64_32KEY, "32-bit"))
    else: # 32-bit Python
        views_to_delete.append((winreg.KEY_WOW64_32KEY, "32-bit"))

    for view_flag, view_name in views_to_delete:
        log_message(f"Attempting to delete COM entries in {view_name} view...")
        for path in com_paths_to_remove:
            try:
                # Use HKEY_LOCAL_MACHINE or HKEY_CURRENT_USER with views if HKEY_CLASSES_ROOT is problematic
                # For HKEY_CLASSES_ROOT, rely on default behavior or use DeleteKeyEx carefully
                # We'll stick to HKEY_CLASSES_ROOT for consistency with creation, but note the caveat.
                winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, path, view_flag, 0)
                log_message(f"Removed COM entry: HKEY_CLASSES_ROOT\\{path} (view: {view_name})", "INFO")
            except FileNotFoundError:
                log_message(f"COM entry not found: HKEY_CLASSES_ROOT\\{path} (view: {view_name})", "DEBUG")
            except Exception as e:
                log_message(f"Error removing COM entry HKEY_CLASSES_ROOT\\{path} (view: {view_name}): {e}", "WARNING")
                unregistration_successful = False

    # --- Remove WPS add-in entries from HKEY_CURRENT_USER ---
    log_message("Attempting to remove WPS add-in entries from HKEY_CURRENT_USER...")
    wps_addin_paths = [
        f"Software\\Kingsoft\\Office\\Addins\\{progid}",
        f"Software\\Kingsoft\\Office\\6.0\\Addins\\{progid}",
        f"Software\\WPS\\Office\\Addins\\{progid}",
        f"Software\\WPS Office\\Addins\\{progid}"
    ]
    
    for path in wps_addin_paths:
        try:
            # Use 0 for default access and view for HKEY_CURRENT_USER
            winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0) 
            log_message(f"Removed WPS Add-in entry at: HKCU\\{path}", "INFO")
        except FileNotFoundError:
            log_message(f"WPS Add-in entry not found at: HKCU\\{path}", "DEBUG")
        except Exception as e:
            log_message(f"Error removing WPS add-in entry at HKCU\\{path}: {e}", "WARNING")
            unregistration_successful = False
            
    if unregistration_successful:
        print("Unregistration complete and successful.")
        return True
    else:
        print("Unregistration finished with some errors. Check logs for details.")
        return False


def run_com_server():
    """Run the COM server with proper factory creation."""
    log_message("Preparing to start COM server...", "INFO")
    try:
        pythoncom.CoInitialize()
        log_message("COM initialized successfully.", "INFO")

        from win32com.server import localserver

        log_message(f"Starting local COM server for class: {WPSAddin.__name__}", "INFO")
        print("COM Server is running. Press Ctrl+C to stop.")

        # This runs the COM message pump until stopped
        localserver.serve([WPSAddin])

    except KeyboardInterrupt:
        log_message("COM server stopped by user.", "INFO")
        print("\nCOM Server stopped.")
    except Exception as e:
        log_message(f"FATAL ERROR while running COM server: {e}\n{traceback.format_exc()}", "ERROR")
        print(f"FATAL ERROR: {e}")
        log_message(f"WPSAddin._reg_clsid_ = {WPSAddin._reg_clsid_}", "DEBUG")
        log_message(f"Type of _reg_clsid_: {type(WPSAddin._reg_clsid_)}", "DEBUG")
    finally:
        try:
            pythoncom.CoUninitialize()
            log_message("COM uninitialized successfully.", "INFO")
        except Exception as e:
            log_message(f"Error during CoUninitialize: {e}", "ERROR")


# --- Main execution block ---
if __name__ == '__main__':
    log_message("Executing main block.", "INFO")
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        if command == '/regserver':
            log_message("Command: /regserver", "INFO")
            success = register_server()
            sys.exit(0 if success else 1)
        elif command == '/unregserver':
            log_message("Command: /unregserver", "INFO")
            success = unregister_server()
            sys.exit(0 if success else 1)
        elif command == '/embedding': # This is how Windows launches a LocalServer32
            log_message("Command: /embedding - Starting COM Server.", "INFO")
            run_com_server()
        else:
            log_message(f"Unknown command line argument: {sys.argv[1]}. Starting COM Server by default.", "WARNING")
            run_com_server()
    else:
        log_message("No command line arguments provided. Starting COM Server.", "INFO")
        run_com_server()