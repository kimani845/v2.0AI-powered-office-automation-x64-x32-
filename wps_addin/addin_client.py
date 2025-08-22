# """
# This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
# It provides the ribbon UI, captures user actions, and sends requests to a separate backend server.
# It supports multiple languages for the UI elements.
# """

# import datetime
# import traceback
# import os
# import sys
# import threading
# import requests
# import win32com.client
# import win32api
# from tkinter import simpledialog, Tk

# # --- Resource Path Helper ---
# def resource_path(filename):
#     """
#     Returns the absolute path to a resource in the wps_addin directory.
#     """
#     base_dir = os.path.dirname(os.path.abspath(__file__))
#     return os.path.join(base_dir, filename)

# # --- Configuration ---
# BACKEND_URL = "http://127.0.0.1:8000"

# # --- Terminal Logging for Debugging ---
# # def log_message(message):
# #     """Writes a message with a timestamp to the terminal (stdout)."""
# #     print(f"[{datetime.datetime.now()}] {message}")


# # PORTABLE File Logging for Debugging 

# def get_log_file_path():
#     """
#     Determines the correct path for the log file.
#     - If running as a bundled .exe, the log appears next to the .exe.
#     - If running as a .py script, the log appears in the project root.
#     """
#     if getattr(sys, 'frozen', False):
#         # We are running in a bundled executable (e.g., from the 'dist' folder).
#         # sys.executable is the full path to the .exe file.
#         # os.path.dirname() gets the directory where the .exe is located.
#         log_directory = os.path.dirname(sys.executable)
#     else:
#         # We are running as a normal .py script.
#         # os.path.abspath(__file__) is the path to this script (addin_client.py).
#         # We go up two levels to get to the project root:
#         # ..\package\addin_client.py -> ..\package -> \project_root
#         log_directory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
#     return os.path.join(log_directory, "wps_addin_log.txt")

# # Define the log path using our smart function
# LOG_FILE_PATH = get_log_file_path()

# def log_message(message):
#     """Writes a message with a timestamp to a log file in the application's directory."""
#     try:
#         # The "a" means append, so it won't overwrite the file each time.
#         with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
#             f.write(f"[{datetime.datetime.now()}] {message}\n")
#     except Exception as e:
#         # If logging fails for any reason (like permissions), print the error.
#         # This won't be visible during WPS startup but is useful for other debugging.
#         print(f"!!! FAILED TO WRITE TO LOG FILE AT {LOG_FILE_PATH}: {e} !!!")
# # --- WPS Helper Functions ---
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

# # --- Ribbon Callback Class ---
# class WPSAddin:
#     _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
#     _reg_desc_ = "AI Office Automation Add-in"
#     _reg_progid_ = "AI.OfficeAddin"
#     _reg_class_spec_ = "addin_client.WPSAddin"

#     _public_methods_ = [
#         'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
#         'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
#         'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel', 
#         'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter'
#     ]
#     _public_attrs_ = ['ribbon']

#     def __init__(self):
#         log_message("--- Add-in __init__ started ---")
#         try:
#             ribbon_path = resource_path('ribbon.xml')
#             log_message(f"Attempting to load ribbon from: {ribbon_path}")
#             if not os.path.exists(ribbon_path):
#                 raise FileNotFoundError(f"Ribbon XML file not found at {ribbon_path}")
#             with open(ribbon_path, 'r', encoding='utf-8') as f:
#                 self.ribbon = f.read()
#             log_message("Ribbon XML loaded successfully.")

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
#                 },
#                 2052: {
#                     "tab": "人工智能助手", "group": "人工智能工具", "run_prompt": "运行通用提示",
#                     "analyze_doc": "分析文档", "summarize_doc": "总结文档",
#                     "create_memo": "创建备忘录", "create_minutes": "创建会议纪要", "create_cover_letter": "创建求职信",
#                     "prompt_title": "人工智能助手", "prompt_message": "请输入您的请求:",
#                     "memo_topic": "请输入备忘录主题:", "memo_audience": "请输入备忘录收件人:",
#                     "minutes_topic": "请输入会议主题:", "minutes_attendees": "请输入与会者(逗号分隔):",
#                     "minutes_info": "请输入关键讨论点:",
#                     "cover_letter_topic": "请输入职位名称:", "cover_letter_audience": "请输入招聘经理/公司:",
#                     "action_cancelled": "操作已取消。", "contacting_server": "人工智能助手：正在连接服务器...",
#                     "connection_error": "\n\n错误：无法连接到后端服务器。\n\n",
#                     "unexpected_error": "\n\n发生意外错误: {e}\n\n",
#                     "result_header": "\n\n--- 人工智能助手结果 ---\n", "result_footer": "\n--- 结果结束 ---\n\n",
#                     "no_active_doc": "未找到活动文档。"
#                 }
#             }
#             log_message("--- Add-in __init__ completed successfully. ---")
#         except FileNotFoundError:
#             log_message(f"!!!!!!!!!! FATAL ERROR IN __init__ !!!!!!!!!!!\nERROR: Ribbon XML file not found at {ribbon_path}.\nTRACEBACK:\n{traceback.format_exc()}\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
#             self.ribbon = ""
#         except Exception as e:
#             error_info = traceback.format_exc()
#             log_message(f"!!!!!!!!!! FATAL ERROR IN __init__ !!!!!!!!!!!\nERROR: {e}\nTRACEBACK:\n{error_info}\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
#             self.ribbon = ""

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
#         if not os.path.exists(image_path):
#             log_message(f"ERROR: Image file not found for '{imageName}' at {image_path}.")
#             return None
#         try:
#             img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
#             log_message(f"Successfully loaded image '{imageName}'.")
#             return img_handle
#         except Exception as e:
#             log_message(f"ERROR: Failed to load image '{imageName}' from {image_path}: {e}\n{traceback.format_exc()}")
#             return None

#     def _call_backend_task(self, endpoint: str, payload: dict):
#         log_message(f"Calling backend endpoint: {endpoint} with payload: {list(payload.keys())}")
#         try:
#             insert_text_at_cursor(self._get_localized_string("contacting_server"))
#             response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=180)
#             response.raise_for_status()
#             result = response.json().get("result", "")
#             header = self._get_localized_string("result_header")
#             footer = self._get_localized_string("result_footer")
#             insert_text_at_cursor(f"{header}{result}{footer}")
#             log_message(f"Successfully received and inserted response from {endpoint}.")
#         except requests.exceptions.ConnectionError as e:
#             log_message(f"ConnectionError calling {endpoint}: {e}")
#             insert_text_at_cursor(self._get_localized_string("connection_error"))
#         except requests.exceptions.HTTPError as e:
#             log_message(f"HTTPError calling {endpoint}: {e}\nServer response: {e.response.text}")
#             insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))
#         except Exception as e:
#             log_message(f"Unexpected Exception calling {endpoint}: {e}\n{traceback.format_exc()}")
#             error_msg = self._get_localized_string("unexpected_error").format(e=e)
#             insert_text_at_cursor(error_msg)

#     def OnRunPrompt(self, c):
#         root = Tk(); root.withdraw()
#         prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("prompt_message"))
#         root.destroy()
#         if not prompt: return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         threading.Thread(target=self._call_backend_task, args=("/process", {"prompt": prompt})).start()

#     def OnAnalyzeDocument(self, c):
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()
        
#     def OnSummarizeDocument(self, c):
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()

#     def OnCreateMemo(self, c):
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_audience"))
#         root.destroy()
#         payload = {
#             "doc_type": "memo",
#             "topic": topic,
#             "audience": audience or "Internal Team",
#             "length": "medium",
#             "tone": "formal"
#         }
#         threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()

#     def OnCreateMinutes(self, c):
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_attendees"))
#         info = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_info"))
#         root.destroy()
#         payload = {
#             "doc_type": "minutes",
#             "topic": topic,
#             "audience": "Meeting Attendees",
#             "length": "medium",
#             "tone": "formal",
#             "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()],
#             "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]
#         }
#         threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()

#     def OnCreateCoverLetter(self, c):
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_audience"))
#         root.destroy()
#         payload = {
#             "doc_type": "cover_letter",
#             "topic": topic,
#             "audience": audience or "Hiring Manager",
#             "length": "medium",
#             "tone": "formal"
#         }
#         threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()
# # COM Server Registration Logic ---
# if __name__ == '__main__':
#     if len(sys.argv) > 1:
#         import win32com.server.register
        
#         # class object for the registration functions
#         classes_to_register = [WPSAddin]
        
#         if sys.argv[1].lower() == '/regserver':
#             log_message("Direct registration command received.")
#             print("Registering AI Office Add-in Client...")
#             try:
#                 # Use the correct, direct registration function
#                 win32com.server.register.RegisterClasses(*classes_to_register)
#                 print("Registration complete.")
#                 log_message("Registration successful.")
#             except Exception as e:
#                 print(f"Registration failed: {e}")
#                 log_message(f"Registration failed: {traceback.format_exc()}")
#                 input("Press Enter to continue...")

#         elif sys.argv[1].lower() == '/unregserver':
#             log_message("Direct unregistration command received.")
#             print("Unregistering AI Office Add-in Client...")
#             try:
#                 # direct unregistration function
#                 win32com.server.register.UnregisterClasses(*classes_to_register)
#                 print("Unregistration complete.")
#                 log_message("Unregistration successful.")
#             except Exception as e:
#                 print(f"Unregistration failed: {e}")
#                 log_message(f"Unregistration failed: {traceback.format_exc()}")
#                 input("Press Enter to continue...")
#     else:
#         print("This is a COM server client for an add-in. Use '/regserver' or '/unregserver'.")
#         input("Press Enter to exit.")
        
        
        
        
        
"""
This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
It provides the ribbon UI, captures user actions, and sends requests to a separate backend server.
It supports multiple languages for the UI elements.
"""
import datetime
import traceback
import os
import sys
import threading
import requests
import win32com.client
import win32api
import winreg # Added for registry manipulation
from tkinter import simpledialog, Tk

# --- Configuration ---
BACKEND_URL = "http://127.0.0.1:8000"

# --- Terminal Logging for Debugging ---
def log_message(message):
    """Writes a message with a timestamp to the terminal (stdout)."""
    print(f"[{datetime.datetime.now()}] {message}")

# --- PyInstaller Resource Path Helper ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller bundling """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- WPS Helper Functions ---
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

# --- Registry Functions for WPS Office Add-in Entry ---

# The name WPS Office uses for your add-in entry.
# This is distinct from the COM component's ProgID but points to its CLSID.
WPS_ADDIN_ENTRY_NAME = "AIAddin.Connect"

def register_wps_addin_entry(guid, wps_entry_name):
    """
    Registers the explicit add-in entry for WPS Office under HKEY_CURRENT_USER.
    This function targets the current user's registry.
    """
    log_message(f"Attempting to register WPS Office Add-in entry for: {wps_entry_name}")
    wps_addin_key_path = f"Software\\Kingsoft\\Office\\Addins\\{wps_entry_name}"
    
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, wps_addin_key_path)
        log_message(f"[SUCCESS] Created/Opened registry key: HKEY_CURRENT_USER\\{wps_addin_key_path}")

        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "AI Office Automation Add-in")
        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Addin")
        winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3) # Load at startup
        winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0) # Not command line safe
        winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, guid)
        log_message(f"[SUCCESS] Set values for {wps_entry_name}. Linked to CLSID: {guid}")

        winreg.CloseKey(key)
        print(f"WPS Office Add-in '{wps_entry_name}' registration complete for current user.")
        return True
    except PermissionError:
        print("\n[FAILURE] Permission denied for WPS Office Add-in entry. Run script as Administrator.")
        return False
    except Exception as e:
        log_message(f"[ERROR] WPS Office Add-in entry registration failed: {e}\n{traceback.format_exc()}")
        print(f"\n[ERROR] WPS Office Add-in entry registration failed: {e}")
        return False

def unregister_wps_addin_entry(wps_entry_name):
    """
    Unregisters the explicit add-in entry for WPS Office under HKEY_CURRENT_USER.
    """
    log_message(f"Attempting to unregister WPS Office Add-in entry for: {wps_entry_name}")
    wps_addin_key_path = f"Software\\Kingsoft\\Office\\Addins\\{wps_entry_name}"

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, wps_addin_key_path)
        print(f"WPS Office Add-in '{wps_entry_name}' unregistration complete for current user.")
        log_message(f"[SUCCESS] Unregistered WPS Office Add-in entry: {wps_entry_name}")
        return True
    except FileNotFoundError:
        print(f"WPS Office Add-in entry for '{wps_entry_name}' not found. Already unregistered or never existed.")
        log_message(f"[INFO] WPS Office Add-in entry '{wps_entry_name}' not found for unregistration.")
        return True # Considered successful if it's already gone
    except PermissionError:
        print("\n[FAILURE] Permission denied for WPS Office Add-in entry unregistration. Run script as Administrator.")
        return False
    except Exception as e:
        log_message(f"[ERROR] WPS Office Add-in entry unregistration failed: {e}\n{traceback.format_exc()}")
        print(f"\n[ERROR] WPS Office Add-in entry unregistration failed: {e}")
        return False

def _update_inprocserver32_path(clsid):
    """
    Manually updates the InprocServer32 key for a given CLSID to point to the
    currently running executable. This is crucial for PyInstaller bundles.
    """
    log_message(f"Attempting to update InprocServer32 for CLSID: {clsid}")
    executable_path = sys.executable
    
    # Determine the registry view based on the current process's bitness
    # This assumes sys.executable matches the bitness of the registry view we want to modify.
    # For a 64-bit Python/EXE, it defaults to 64-bit view.
    # For a 32-bit Python/EXE on 64-bit Windows, it defaults to 32-bit view via WoW64.
    if sys.maxsize > 2**32:
        # Running as a 64-bit process
        reg_flags = winreg.KEY_READ | winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
        arch_name = "64-bit"
    else:
        # Running as a 32-bit process
        reg_flags = winreg.KEY_READ | winreg.KEY_SET_VALUE | winreg.KEY_WOW64_32KEY
        arch_name = "32-bit"

    try:
        # Open the CLSID key in HKLM
        clsid_key_path = f"SOFTWARE\\Classes\\CLSID\\{clsid}"
        clsid_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, clsid_key_path, 0, reg_flags)

        # Open or create the InprocServer32 subkey
        inproc_key = winreg.CreateKey(clsid_key, "InprocServer32")
        
        # Set the default value of InprocServer32 to the executable path
        winreg.SetValueEx(inproc_key, None, 0, winreg.REG_SZ, executable_path) # Set default value
        winreg.SetValueEx(inproc_key, "ThreadingModel", 0, winreg.REG_SZ, "Both") # Standard threading model

        winreg.CloseKey(inproc_key)
        winreg.CloseKey(clsid_key)
        log_message(f"[SUCCESS] InprocServer32 for {clsid} updated to '{executable_path}' ({arch_name} view).")
        print(f"InprocServer32 path set to: '{executable_path}' ({arch_name} view).")
        return True
    except PermissionError:
        log_message(f"[FAILURE] Permission denied when updating InprocServer32 for {clsid}. Run as Administrator.")
        print(f"[FAILURE] Permission denied when updating InprocServer32. Ensure executable is run as Administrator.")
        return False
    except Exception as e:
        log_message(f"[ERROR] Failed to update InprocServer32 for {clsid}: {e}\n{traceback.format_exc()}")
        print(f"[ERROR] Failed to update InprocServer32: {e}")
        return False


# --- Ribbon Callback Class ---
class WPSAddin:
    # IMPORTANT: Ensure this CLSID matches the one in your .py file and other registration scripts.
    _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
    _reg_desc_ = "AI Office Automation Add-in"
    # SUGGESTED CHANGE: Explicitly use the constant here for consistency
    _reg_progid_ = WPS_ADDIN_ENTRY_NAME 
    _reg_class_spec_ = "addin_client.WPSAddin"

    _public_methods_ = [
        'OnRunPrompt', 'OnAnalyzeDocument', 'OnSummarizeDocument', 'OnLoadImage',
        'GetTabLabel', 'GetGroupLabel', 'GetRunPromptLabel', 'GetAnalyzeDocLabel',
        'GetSummarizeDocLabel', 'GetCreateMemoLabel', 'GetCreateMinutesLabel', 
        'GetCreateCoverLetterLabel', 'OnCreateMemo', 'OnCreateMinutes', 'OnCreateCoverLetter'
    ]
    _public_attrs_ = ['ribbon']

    def __init__(self):
        log_message("--- Add-in __init__ started ---")
        try:
            ribbon_path = resource_path('ribbon.xml')
            log_message(f"Attempting to load ribbon from: {ribbon_path}")
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
                },
                2052: { # Chinese (Simplified)
                    "tab": "人工智能助手", "group": "人工智能工具", "run_prompt": "运行通用提示",
                    "analyze_doc": "分析文档", "summarize_doc": "总结文档",
                    "create_memo": "创建备忘录", "create_minutes": "创建会议纪要", "create_cover_letter": "创建求职信",
                    "prompt_title": "人工智能助手", "prompt_message": "请输入您的请求:",
                    "memo_topic": "请输入备忘录主题:", "memo_audience": "请输入备忘录收件人:",
                    "minutes_topic": "请输入会议主题:", "minutes_attendees": "请输入与会者(逗号分隔):",
                    "minutes_info": "请输入关键讨论点:",
                    "cover_letter_topic": "请输入职位名称:", "cover_letter_audience": "请输入招聘经理/公司:",
                    "action_cancelled": "操作已取消。", "contacting_server": "人工智能助手：正在连接服务器...",
                    "connection_error": "\n\n错误：无法连接到后端服务器。\n\n",
                    "unexpected_error": "\n\n发生意外错误: {e}\n\n",
                    "result_header": "\n\n--- 人工智能助手结果 ---\n", "result_footer": "\n--- 结果结束 ---\n\n",
                    "no_active_doc": "未找到活动文档。"
                }
            }
            log_message("--- Add-in __init__ completed successfully. ---")
        except FileNotFoundError:
            log_message(f"!!!!!!!!!! FATAL ERROR IN __init__ !!!!!!!!!!!\nERROR: Ribbon XML file not found at {ribbon_path}.\nTRACEBACK:\n{traceback.format_exc()}\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            self.ribbon = ""
        except Exception as e:
            error_info = traceback.format_exc()
            log_message(f"!!!!!!!!!! FATAL ERROR IN __init__ !!!!!!!!!!!\nERROR: {e}\nTRACEBACK:\n{error_info}\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            self.ribbon = ""

    def _get_localized_string(self, key):
        lang_id = 1033 # Default to English
        wps_app = get_wps_application()
        if wps_app:
            try:
                # Get the current language ID from WPS Office
                lang_id = wps_app.LanguageSettings.LanguageID(1) # msoAppLanguageIDInstall
            except Exception:
                pass # Fallback to default lang_id
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
            # win32api.LoadImage requires an absolute path. The resource_path helper handles this.
            # 0x10 is LR_LOADFROMFILE
            img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10)
            log_message(f"Successfully loaded image '{imageName}'.")
            return img_handle
        except FileNotFoundError:
            log_message(f"ERROR: Image file not found for '{imageName}' at {image_path}.\n{traceback.format_exc()}")
            return None
        except Exception as e:
            log_message(f"ERROR: Failed to load image '{imageName}' from {image_path}: {e}\n{traceback.format_exc()}")
            return None

    def _call_backend_task(self, endpoint: str, payload: dict):
        log_message(f"Calling backend endpoint: {endpoint} with payload: {list(payload.keys())}")
        try:
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
            # Increased timeout for potentially long-running AI tasks
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=300) 
            response.raise_for_status()
            result = response.json().get("result", "")
            header = self._get_localized_string("result_header")
            footer = self._get_localized_string("result_footer")
            insert_text_at_cursor(f"{header}{result}{footer}")
            log_message(f"Successfully received and inserted response from {endpoint}.")
        except requests.exceptions.ConnectionError as e:
            log_message(f"ConnectionError calling {endpoint}: {e}")
            insert_text_at_cursor(self._get_localized_string("connection_error"))
        except requests.exceptions.HTTPError as e:
            log_message(f"HTTPError calling {endpoint}: {e}\nServer response: {e.response.text}")
            insert_text_at_cursor(self._get_localized_string("unexpected_error").format(e=e))
        except Exception as e:
            log_message(f"Unexpected Exception calling {endpoint}: {e}\n{traceback.format_exc()}")
            error_msg = self._get_localized_string("unexpected_error").format(e=e)
            insert_text_at_cursor(error_msg)

    def OnRunPrompt(self, c):
        root = Tk(); root.withdraw() # Create and immediately hide Tkinter root window
        prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("prompt_message"))
        root.destroy() # Destroy the root window after interaction
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
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_audience"))
        root.destroy()
        payload = {
            "doc_type": "memo",
            "topic": topic,
            "audience": audience or "Internal Team",
            "length": "medium",
            "tone": "formal"
        }
        threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()

    def OnCreateMinutes(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_topic"))
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_attendees"))
        info = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("minutes_info"))
        root.destroy()
        payload = {
            "doc_type": "minutes",
            "topic": topic,
            "audience": "Meeting Attendees",
            "length": "medium",
            "tone": "formal",
            "members_present": [name.strip() for name in (attendees or "").split(',') if name.strip()],
            "data_sources": [data.strip() for data in (info or "").split(',') if data.strip()]
        }
        threading.Thread(target=self._call_backend_task, args=("/create_minutes", payload)).start()

    def OnCreateCoverLetter(self, c):
        root = Tk(); root.withdraw()
        topic = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_topic"))
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("cover_letter_audience"))
        root.destroy()
        payload = {
            "doc_type": "cover_letter",
            "topic": topic,
            "audience": audience or "Hiring Manager",
            "length": "medium",
            "tone": "formal"
        }
        threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()

# --- COM Server Registration Logic ---
if __name__ == '__main__':
    # Ensure pywin32 is installed: pip install pywin32
    # Ensure pywin32 is correctly registered: python -m win32com.client.makepy -d
    # Then run this script with /regserver or /unregserver as Administrator
    # Example: python addin_client.py /regserver
    # For a 32-bit COM server, use a 32-bit Python interpreter.
    # For a 64-bit COM server, use a 64-bit Python interpreter.
    
    if len(sys.argv) > 1:
        import win32com.server.register
        
        # Class object for the registration functions
        classes_to_register = [WPSAddin]
        
        if sys.argv[1].lower() == '/regserver':
            log_message("Direct registration command received.")
            print("Registering AI Office Add-in Client (COM component and WPS entry)...")
            try:
                # 1. Register the COM component itself
                win32com.server.register.RegisterClasses(*classes_to_register)
                print(f"COM component '{WPSAddin._reg_progid_}' registered successfully.")
                log_message("COM component registration successful.")

                # 2. Manually update InprocServer32 to point to the current executable path
                _update_inprocserver32_path(WPSAddin._reg_clsid_)

                # 3. Register the WPS Office specific add-in entry
                register_wps_addin_entry(WPSAddin._reg_clsid_, WPS_ADDIN_ENTRY_NAME)
                
                print("All registrations complete.")
            except Exception as e:
                print(f"Registration failed: {e}") 
                log_message(f"Registration failed: {traceback.format_exc()}")
            input("Press Enter to continue...")

        elif sys.argv[1].lower() == '/unregserver':
            log_message("Direct unregistration command received.")
            print("Unregistering AI Office Add-in Client (COM component and WPS entry)...")
            try:
                # 1. Unregister the COM component itself
                win32com.server.register.UnregisterClasses(*classes_to_register)
                print(f"COM component '{WPSAddin._reg_progid_}' unregistered successfully.")
                log_message("COM component unregistration successful.")

                # 2. Unregister the WPS Office specific add-in entry
                unregister_wps_addin_entry(WPS_ADDIN_ENTRY_NAME)
                
                print("All unregistrations complete.")
            except Exception as e:
                print(f"Unregistration failed: {e}")
                log_message(f"Unregistration failed: {traceback.format_exc()}")
            input("Press Enter to continue...")
        else:
            print("This is a COM server client for an add-in. Use '/regserver' or '/unregserver'.")
            input("Press Enter to exit.")
    else:
        # If no arguments, and not running as a COM server, keep the script alive for debugging/inspection
        # You might want to remove this block in a final bundled application
        print("Running in COM server mode. Awaiting calls from WPS Office.")
        log_message("addin_client.py started in direct execution mode (no /regserver or /unregserver).")
        # Keep the script running so WPS can connect if it's already registered and trying to load.
        # This is typically only useful for debugging; in a deployed scenario, WPS would launch it.
        try:
            # Prevent script from exiting immediately if not registering/unregistering
            import pythoncom
            pythoncom.PumpMessages()
        except KeyboardInterrupt:
            print("\nExiting COM server client.")
        except Exception as e:
            log_message(f"Error in COM message pump: {e}")
            print(f"\nError in COM message pump: {e}")
        input("Press Enter to exit.")
