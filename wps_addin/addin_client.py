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
import winreg # For registry manipulation
from tkinter import simpledialog, Tk
import pythoncom # New import for COM object handling

# Configuration
BACKEND_URL = "http://127.0.0.1:8000"

def get_log_file_path():
    """Determines the correct path for the log file (next to the exe or script)."""
    if getattr(sys, 'frozen', False):
        log_directory = os.path.dirname(sys.executable)
    else:
        # Assumes this script is in a subfolder like 'package' or 'wps_addin'
        log_directory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(log_directory, "wps_addin_log.txt")

LOG_FILE_PATH = get_log_file_path()

def log_message(message):
    """Writes a message with a timestamp to the log file."""
    try:
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now()}] {message}\n")
    except Exception as e:
        print(f"!!! FAILED TO WRITE TO LOG FILE AT {LOG_FILE_PATH}: {e} !!!")

# # --- Terminal Logging for Debugging ---
# def log_message(message):
#     """Writes a message with a timestamp to the terminal (stdout)."""
#     print(f"[{datetime.   datetime.now()}] {message}")

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

# The name WPS Office uses for your add-in entry.
WPS_ADDIN_ENTRY_NAME = "AIAddinConnect"

# Ribbon Callback Class
class WPSAddin:
    # IMPORTANT: Ensure this CLSID matches the one in your .py file and other registration scripts.
    _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
    _reg_desc_ = "AI Office Automation "
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
            img_handle = win32api.LoadImage(0, image_path, 0, 32, 32, 0x10) # 0x10 is LR_LOADFROMFILE
            log_message(f"Successfully loaded image '{imageName}'.")
            return img_handle
        except Exception as e:
            log_message(f"ERROR: Failed to load image '{imageName}' from {image_path}: {e}\n{traceback.format_exc()}")
            return None

    def _call_backend_task(self, endpoint: str, payload: dict):
        log_message(f"Calling backend endpoint: {endpoint} with payload: {list(payload.keys())}")
        try:
            insert_text_at_cursor(self._get_localized_string("contacting_server"))
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
        if not topic:
            root.destroy()
            return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
        audience = simpledialog.askstring(self._get_localized_string("prompt_title"), self._get_localized_string("memo_audience"))
        root.destroy()
        payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
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
            "doc_type": "minutes", "topic": topic, "audience": "Meeting Attendees",
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
        payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
        threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()


if __name__ == '__main__':
    # This block performs a fully manual registration to avoid pywin32 bugs
    # and handles multiple possible WPS Office registry paths.

    # def manual_register_server(cls):
    #     """
    #     Manually creates all necessary registry keys.
    #     This version now tries multiple common WPS Office registry paths.
    #     """
    
    #     clsid = cls._reg_clsid_
    #     progid = cls._reg_progid_
    #     desc = cls._reg_desc_

    #     is_64bit_process = sys.maxsize > 2**32
    #     reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY

    #     log_message(f"Starting manual registration for {'64-bit' if is_64bit_process else '32-bit'} view.")

        # # --- Steps 1-4: Register the core COM Server ---
        # try:
        #     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, winreg.KEY_WRITE | reg_view_flag) as key:
        #         winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
        #     log_message("Created CLSID key with description.")
            
        #     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, winreg.KEY_WRITE | reg_view_flag) as progid_key:
        #         winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
        #     log_message("Created ProgID subkey.")
            
        #     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32", 0, winreg.KEY_WRITE | reg_view_flag) as inproc_key:
        #         pythoncom_path = pythoncom.__file__
        #         winreg.SetValueEx(inproc_key, "", 0, winreg.REG_SZ, pythoncom_path)
        #     log_message(f"Created InprocServer32 subkey pointing to: {pythoncom_path}")
            
        #     with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, winreg.KEY_WRITE | reg_view_flag) as key:
        #         with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
        #             winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
        #     log_message("Created ProgID-to-CLSID mapping.")
            
        # except Exception as e:
        #     log_message(f"FATAL: Failed to register the core COM server: {e}\n{traceback.format_exc()}")
        #     print(f"FATAL: Failed to register the core COM server. Please run as Administrator.")
        #     return

        # # --- Step 5: Create the specific entry for WPS Office (UPDATED LOGIC) ---
        # wps_addin_paths = [
        #     f"Software\\Kingsoft\\Office\\Addins\\{progid}",
        #     f"Software\\WPS\\Office\\Addins\\{progid}",
        #     f"Software\\WPS Office\\Addins\\{progid}"
        # ]
        # # wps_addin_paths = [
        # #     f"Software\\WPS Office\\Addins\\{progid}",
        # #     f"Software\\WPS\\Office\\Addins\\{progid}",
        # #     f"Software\\Kingsoft\\Office\\Addins\\{progid}"
        # # ]
        # # wps_addin_paths = [
        # #     f"Software\\WPS Office\\Addins\\{progid}",
        # #     f"Software\\WPS\\Office\\Addins\\{progid}"
        # # ]
        
        # registration_succeeded = False
        # for path in wps_addin_paths:
        #     log_message(f"Attempting to register WPS add-in at HKCU\\{path}...")
        #     try:
        #         with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
        #             winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
        #             winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
        #             winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        #             winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
                
        #         log_message(f"SUCCESS: Created specific WPS Add-in entry at HKCU\\{path}")
        #         print(f"SUCCESS: Add-in registered at registry path: HKCU\\{path}")
        #         registration_succeeded = True
        #         break
        #     except Exception as e:
        #         log_message(f"INFO: Could not register at path '{path}'. Error: {e}. Trying next path.")

        # if registration_succeeded:
        #     print("Manual registration successful.")
        #     log_message("Manual registration successful.")
        # else:
        #     print("\n[FAILURE] Manual registration FAILED. Could not write to any known WPS Office registry locations.")
        #     log_message("Manual registration FAILED. Could not write to any known WPS Office registry locations.")
        
    def manual_register_server(cls):
        """Manually creates all necessary registry keys for the COM server and WPS."""
        
        clsid = cls._reg_clsid_
        progid = cls._reg_progid_
        desc = cls._reg_desc_

        is_64bit_process = sys.maxsize > 2**32
        reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY

        log_message(f"Starting manual registration for {'64-bit' if is_64bit_process else '32-bit'} view.")

        try:
            # Step 1 & 2: Create CLSID key and ProgID subkey
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, winreg.KEY_WRITE | reg_view_flag) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
                with winreg.CreateKeyEx(key, "ProgID", 0, winreg.KEY_WRITE) as progid_key:
                    winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)

                # Step 3: Create the InprocServer32 subkey with the CORRECT DLL path
                with winreg.CreateKeyEx(key, "InprocServer32", 0, winreg.KEY_WRITE) as inproc_key:
                    # Gets the path to the pythoncom DLL, which is what WPS needs to load.
                    pythoncom_path = pythoncom.__file__
                    
                    # Write the correct DLL path to the (Default) value.
                    winreg.SetValueEx(inproc_key, "", 0, winreg.REG_SZ, pythoncom_path)
                    
                    # The "ThreadingModel" is a best practice for COM servers.
                    winreg.SetValueEx(inproc_key, "ThreadingModel", 0, winreg.REG_SZ, "Apartment")
                    
                    log_message(f"Created InprocServer32 subkey pointing to the correct DLL: {pythoncom_path}")

            # Step 4: Create the ProgID-to-CLSID mapping
            with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, winreg.KEY_WRITE | reg_view_flag) as key:
                with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
                    winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
            log_message("Core COM registration successful.")

            # Step 5: Create the specific entry for WPS Office
            wps_addin_path = f"Software\\Kingsoft\\Office\\Addins\\{progid}"
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, wps_addin_path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
                winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                log_message(f"Created specific WPS Add-in entry at HKCU\\{wps_addin_path}")

            print("Manual registration successful.")
            log_message("Manual registration successful.")

        except Exception as e:
            print(f"Manual registration failed: {e}")
            log_message(f"Manual registration failed:\n{traceback.format_exc()}")
            input("Press Enter to continue...")



    def manual_unregister_server(cls):
        """
        Manually removes all registry keys.
        --- UPDATED ---
        This version now tries to unregister from multiple common WPS Office registry paths.
        """
        import winreg

        clsid = cls._reg_clsid_
        progid = cls._reg_progid_

        is_64bit_process = sys.maxsize > 2**32
        reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY

        log_message(f"Starting manual unregistration for {'64-bit' if is_64bit_process else '32-bit'} view.")

        # --- Step 1: Delete WPS Office entry from all potential locations (UPDATED LOGIC) ---
        wps_addin_paths = [
            f"Software\\Kingsoft\\Office\\Addins\\{progid}",
            f"Software\\WPS\\Office\\Addins\\{progid}",
            f"Software\\WPS Office\\Addins\\{progid}"
        ]
        for path in wps_addin_paths:
            try:
                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0)
                log_message(f"Successfully deleted WPS Add-in key from HKCU\\{path}.")
                print(f"Removed registry key: HKCU\\{path}")
            except FileNotFoundError:
                log_message(f"INFO: WPS Add-in key not found at HKCU\\{path}. Skipping.")
            except Exception as e:
                log_message(f"WARNING: Could not delete WPS Add-in key from {path}. Error: {e}")

        # --- Step 2 & 3: Delete the core COM Server ---
        try:
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, progid, reg_view_flag, 0)
            log_message("Deleted ProgID mapping.")
        except FileNotFoundError:
            pass
        except Exception as e:
            log_message(f"Could not delete ProgID key: {e}")

        try:
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", reg_view_flag, 0)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", reg_view_flag, 0)
            log_message("Deleted CLSID key.")
        except FileNotFoundError:
            pass
        except Exception as e:
            log_message(f"Could not delete CLSID key: {e}")

        print("Manual unregistration complete.")
        log_message("Manual unregistration complete.")


    # Main Command Logic
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            manual_register_server(WPSAddin)
        elif sys.argv[1].lower() == '/unregserver':
            manual_unregister_server(WPSAddin)
    else:
        print("This is a COM server client for a WPS Office add-in.")
        print("Use '/regserver' to register or '/unregserver' to unregister.")
        input("Press Enter to exit.")


# """
# This script runs as a COM server client to provide an AI Assistant add-in for WPS Office.
# Fixed version addressing common WPS Office add-in loading issues.
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
# from tkinter import simpledialog, Tk

# # Configuration - FIXED: Corrected IP address
# BACKEND_URL = "http://127.0.0.1:8000"

# # FIXED: Consistent naming
# WPS_ADDIN_ENTRY_NAME = "AIAddin.Connect"

# def log_message(message):
#     """Writes a message with a timestamp to the terminal (stdout)."""
#     print(f"[{datetime.datetime.now()}] {message}")

# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller bundling """
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)

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
#     _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
#     _reg_desc_ = "AI Office Automation Add-in"
#     _reg_progid_ = WPS_ADDIN_ENTRY_NAME  # FIXED: Now consistent
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
#                 }
#             }
#             log_message("--- Add-in __init__ completed successfully. ---")
#         except FileNotFoundError:
#             log_message(f"FATAL ERROR: Ribbon XML file not found at {ribbon_path}")
#             self.ribbon = ""
#         except Exception as e:
#             log_message(f"FATAL ERROR IN __init__: {e}\n{traceback.format_exc()}")
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
#         root = Tk(); root.withdraw()
#         prompt = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                       self._get_localized_string("prompt_message"))
#         root.destroy()
#         if not prompt: 
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/process", {"prompt": prompt})).start()

#     def OnAnalyzeDocument(self, c):
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/analyze", {"content": content, "prompt": "Analyze the document content."})).start()

#     def OnSummarizeDocument(self, c):
#         wps_app = get_wps_application()
#         if not wps_app or wps_app.Documents.Count == 0: 
#             return insert_text_at_cursor(self._get_localized_string("no_active_doc"))
#         content = wps_app.ActiveDocument.Content.Text
#         threading.Thread(target=self._call_backend_task, 
#                         args=("/summarize", {"content": content, "prompt": "Summarize the document content."})).start()

#     def OnCreateMemo(self, c):
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                      self._get_localized_string("memo_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("memo_audience"))
#         root.destroy()
#         payload = {"doc_type": "memo", "topic": topic, "audience": audience or "Internal Team"}
#         threading.Thread(target=self._call_backend_task, args=("/create_memo", payload)).start()

#     def OnCreateMinutes(self, c):
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                      self._get_localized_string("minutes_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         attendees = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                          self._get_localized_string("minutes_attendees"))
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
#         root = Tk(); root.withdraw()
#         topic = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                      self._get_localized_string("cover_letter_topic"))
#         if not topic:
#             root.destroy()
#             return insert_text_at_cursor(self._get_localized_string("action_cancelled"))
#         audience = simpledialog.askstring(self._get_localized_string("prompt_title"), 
#                                         self._get_localized_string("cover_letter_audience"))
#         root.destroy()
#         payload = {"doc_type": "cover_letter", "topic": topic, "audience": audience or "Hiring Manager"}
#         threading.Thread(target=self._call_backend_task, args=("/create_cover_letter", payload)).start()


# if __name__ == '__main__':
#     def manual_register_server(cls):
#         """FIXED: Enhanced registration with proper InprocServer32 handling"""
#         import winreg
#         import pythoncom

#         clsid = cls._reg_clsid_
#         progid = cls._reg_progid_
#         desc = cls._reg_desc_

#         is_64bit_process = sys.maxsize > 2**32
#         reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY

#         log_message(f"Starting manual registration for {'64-bit' if is_64bit_process else '32-bit'} view.")

#         # Register core COM Server
#         try:
#             # Main CLSID key
#             with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, winreg.KEY_WRITE | reg_view_flag) as key:
#                 winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
            
#             # ProgID subkey
#             with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", 0, winreg.KEY_WRITE | reg_view_flag) as progid_key:
#                 winreg.SetValueEx(progid_key, "", 0, winreg.REG_SZ, progid)
            
#             # FIXED: InprocServer32 with proper threading model
#             with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32", 0, winreg.KEY_WRITE | reg_view_flag) as inproc_key:
#                 # For PyInstaller, use the executable path instead of pythoncom
#                 executable_path = sys.executable
#                 winreg.SetValueEx(inproc_key, "", 0, winreg.REG_SZ, executable_path)
#                 winreg.SetValueEx(inproc_key, "ThreadingModel", 0, winreg.REG_SZ, "Apartment")  # FIXED: Added threading model
            
#             # ProgID-to-CLSID mapping
#             with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid, 0, winreg.KEY_WRITE | reg_view_flag) as key:
#                 winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
#                 with winreg.CreateKeyEx(key, "CLSID", 0, winreg.KEY_WRITE) as clsid_key:
#                     winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
            
#             log_message("Core COM server registration completed successfully.")
#         except Exception as e:
#             log_message(f"FATAL: Failed to register core COM server: {e}")
#             print("FATAL: Failed to register the core COM server. Please run as Administrator.")
#             return

#         # Register WPS Office add-in entry
#         wps_addin_paths = [
#             f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#             f"Software\\WPS\\Office\\Addins\\{progid}",
#             f"Software\\WPS Office\\Addins\\{progid}"
#         ]
        
#         registration_succeeded = False
#         for path in wps_addin_paths:
#             try:
#                 with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
#                     winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, desc)
#                     winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
#                     winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
#                     winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
#                     # FIXED: Added CommandLineSafe
#                     winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
                
#                 log_message(f"SUCCESS: WPS add-in registered at HKCU\\{path}")
#                 print(f"SUCCESS: Add-in registered at: HKCU\\{path}")
#                 registration_succeeded = True
#                 break
#             except Exception as e:
#                 log_message(f"Could not register at {path}: {e}")

#         if registration_succeeded:
#             print("\nRegistration completed successfully!")
#             print("Please restart WPS Office to see the add-in.")
#         else:
#             print("\nFAILED: Could not register WPS Office add-in entry.")

#     def manual_unregister_server(cls):
#         """Enhanced unregistration"""
#         import winreg

#         clsid = cls._reg_clsid_
#         progid = cls._reg_progid_

#         is_64bit_process = sys.maxsize > 2**32
#         reg_view_flag = winreg.KEY_WOW64_64KEY if is_64bit_process else winreg.KEY_WOW64_32KEY

#         # Remove WPS Office entries
#         wps_addin_paths = [
#             f"Software\\Kingsoft\\Office\\Addins\\{progid}",
#             f"Software\\WPS\\Office\\Addins\\{progid}",
#             f"Software\\WPS Office\\Addins\\{progid}"
#         ]
        
#         for path in wps_addin_paths:
#             try:
#                 winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0)
#                 log_message(f"Removed: HKCU\\{path}")
#             except FileNotFoundError:
#                 pass
#             except Exception as e:
#                 log_message(f"Could not remove {path}: {e}")

#         # Remove COM server entries
#         try:
#             winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID", reg_view_flag, 0)
#             winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, progid, reg_view_flag, 0)
#             winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32", reg_view_flag, 0)
#             winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID", reg_view_flag, 0)
#             winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", reg_view_flag, 0)
#             log_message("Removed COM server entries.")
#         except FileNotFoundError:
#             pass
#         except Exception as e:
#             log_message(f"Could not remove COM entries: {e}")

#         print("Unregistration complete.")

#     # Main command logic
#     if len(sys.argv) > 1:
#         if sys.argv[1].lower() == '/regserver':
#             manual_register_server(WPSAddin)
#         elif sys.argv[1].lower() == '/unregserver':
#             manual_unregister_server(WPSAddin)
#     else:
#         print("WPS Office AI Assistant Add-in")
#         print("Usage:")
#         print("  /regserver   - Register the add-in")
#         print("  /unregserver - Unregister the add-in")
#         input("\nPress Enter to exit...")