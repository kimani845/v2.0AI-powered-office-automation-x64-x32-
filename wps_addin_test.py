# import threading
# from wps_addin import WPSAddin

# # Dummy function to simulate inserting text (instead of WPS)
# def dummy_insert_text(text):
#     print(f"\n=== INSERTED TEXT ===\n{text}\n=====================")

# # Patch the insert_text_at_cursor function in the add-in
# import wps_addin
# wps_addin.insert_text_at_cursor = dummy_insert_text

# # Instantiate the add-in
# addin = WPSAddin()

# # Simulate backend responses without real server
# def fake_backend(endpoint, payload):
#     print(f"[FAKE BACKEND] Called endpoint: {endpoint}")
#     print(f"[FAKE BACKEND] Payload: {payload}")
#     return "Fake response from AI backend."

# # Patch _call_backend_task to use the fake backend
# def patched_call_backend(self, endpoint, payload):
#     result = fake_backend(endpoint, payload)
#     header = self._get_localized_string("result_header")
#     footer = self._get_localized_string("result_footer")
#     dummy_insert_text(f"{header}{result}{footer}")

# WPSAddin._call_backend_task = patched_call_backend

# # Simulate all button actions
# def run_tests():
#     print("=== TEST: Run General Prompt ===")
#     addin.OnRunPrompt(None)  # This will open dialogs

#     print("=== TEST: Analyze Document ===")
#     addin.OnAnalyzeDocument(None)

#     print("=== TEST: Summarize Document ===")
#     addin.OnSummarizeDocument(None)

#     print("=== TEST: Create Memo ===")
#     addin.OnCreateMemo(None)

#     print("=== TEST: Create Minutes ===")
#     addin.OnCreateMinutes(None)

#     print("=== TEST: Create Cover Letter ===")
#     addin.OnCreateCoverLetter(None)

# # Run in a separate thread to allow dialogs to appear
# threading.Thread(target=run_tests).start()

import sys
import os
import traceback

# Ensure the wps_addin package is on the path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    # from wps_addin import WPSAddin
    from wps_addin.addin_client import WPSAddin

except ImportError as e:
    print(f"[ERROR] Could not import WPSAddin: {e}")
    traceback.print_exc()
    sys.exit(1)

def main():
    print("=== WPS AI Add-in Test ===")
    
    try:
        addin = WPSAddin()
        print("[SUCCESS] WPSAddin instance created.")
    except Exception as e:
        print(f"[ERROR] Failed to create WPSAddin instance: {e}")
        traceback.print_exc()
        return

    try:
        ribbon_content = addin.GetCustomUI("AnyRibbonID")
        if ribbon_content:
            print(f"[SUCCESS] Ribbon loaded successfully. Length: {len(ribbon_content)} characters.")
        else:
            print("[WARNING] Ribbon content is empty.")
    except Exception as e:
        print(f"[ERROR] Failed to load ribbon XML: {e}")
        traceback.print_exc()

    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        if cmd == "/regserver":
            print("Registering COM server...")
            result = addin.register_server()
            print(f"Registration result: {result}")
        elif cmd == "/unregserver":
            print("Unregistering COM server...")
            result = addin.unregister_server()
            print(f"Unregistration result: {result}")
        else:
            print(f"Unknown argument: {cmd}")

    print("\nTest complete.")

if __name__ == "__main__":
    main()
