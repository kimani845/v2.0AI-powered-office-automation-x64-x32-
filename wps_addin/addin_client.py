


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

# --- Ribbon Callback Class ---
class WPSAddin:
    _reg_clsid_ = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
    _reg_desc_ = "AI Office Automation Add-in"
    _reg_progid_ = "AI.OfficeAddin"
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
                2052: {
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
            response = requests.post(f"{BACKEND_URL}{endpoint}", json=payload, timeout=180)
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
# COM Server Registration Logic ---
if __name__ == '__main__':
    if len(sys.argv) > 1:
        import win32com.server.register
        
        # class object for the registration functions
        classes_to_register = [WPSAddin]
        
        if sys.argv[1].lower() == '/regserver':
            log_message("Direct registration command received.")
            print("Registering AI Office Add-in Client...")
            try:
                # Use the correct, direct registration function
                win32com.server.register.RegisterClasses(*classes_to_register)
                print("Registration complete.")
                log_message("Registration successful.")
            except Exception as e:
                print(f"Registration failed: {e}")
                log_message(f"Registration failed: {traceback.format_exc()}")
                input("Press Enter to continue...")

        elif sys.argv[1].lower() == '/unregserver':
            log_message("Direct unregistration command received.")
            print("Unregistering AI Office Add-in Client...")
            try:
                # direct unregistration function
                win32com.server.register.UnregisterClasses(*classes_to_register)
                print("Unregistration complete.")
                log_message("Unregistration successful.")
            except Exception as e:
                print(f"Unregistration failed: {e}")
                log_message(f"Unregistration failed: {traceback.format_exc()}")
                input("Press Enter to continue...")
    else:
        print("This is a COM server client for an add-in. Use '/regserver' or '/unregserver'.")
        input("Press Enter to exit.")