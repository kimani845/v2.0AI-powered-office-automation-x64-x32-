import argparse
import os
import sys
import subprocess
import requests
from dotenv import load_dotenv

# Try to import WPS-related libraries. Fail gracefully if not on Windows.
try:
    import win32com.client
    import docx
    IS_WINDOWS = True
except ImportError:
    IS_WINDOWS = False

# FastAPI server's base URL
API_BASE_URL = "http://127.0.0.1:8000"

def get_active_wps_document():
    """
    Connects to the active WPS writer instance and returns the document object.
    """
    if not IS_WINDOWS:
        print("WPS integration is only supported in windows!")
        return None
    try:
        wps_app = win32com.client.GetActiveObject("kwps.Application")
        if wps_app.Documents.Count > 0:
            return wps_app.ActiveDocument
        else:
            print("WPS is running but no document is open.")
            return None
    except Exception as e:
        print(f"Error: Could not connect to WPS writer. Is a document open? Details: {e}")
        return None

def get_wps_content(wps_doc):
    """
    Extracts the text content from the active WPS document.
    """
    if not IS_WINDOWS or not wps_doc:
        return ""
    try:
        return wps_doc.Content.Text
    except Exception as e:
        print(f"Error reading content from WPS document: {e}")
        return ""

def insert_content_into_wps(wps_doc, content: str):
    """
    Reads content from a string and inserts it into the active WPS document.
    """
    if not IS_WINDOWS or not wps_doc:
        return
    try:
        print("-> Inserting generated content into active document...")
        wps_doc.Content.InsertAfter("\n\n" + "=" * 25 + " AI Assistant Result " + "=" * 25 + "\n\n")
        wps_doc.Content.InsertAfter(content)
        print("Content successfully inserted into WPS document.")
    except Exception as e:
        print(f"Error inserting content into WPS: {e}")

def save_content_to_docx(content: str, filename_prefix: str) -> str:
    """Saves a string as a new .docx file."""
    try:
        from docx import Document
        doc = Document()
        
        doc.add_paragraph(content)
        
        sanitized_prefix = "".join(c for c in filename_prefix if c.isalnum() or c in " _-").rstrip()
        output_filename = f"{sanitized_prefix.replace(' ', '_')}_{os.urandom(4).hex()}.docx"
        output_filepath = os.path.join(os.getcwd(), "output", output_filename)
        
        os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
        doc.save(output_filepath)
        return output_filepath
    except Exception as e:
        print(f"Error saving content to file: {e}")
        return ""

def open_file_os_agnostic(filepath: str):
    """Opens a file using the default application for the current OS."""
    try:
        abs_path = os.path.abspath(filepath)
        if sys.platform == "win32":
            os.startfile(abs_path)
        elif sys.platform == "darwin":
            subprocess.run(['open', abs_path], check=True)
        elif sys.platform.startswith('linux'):
            subprocess.run(['xdg-open', abs_path], check=True)
        else:
            print(f"Could not automatically open file. Please open it manually at: {abs_path}")
    except Exception as e:
        print(f"Error automatically opening file: {e}")

def main():
    """
    Main function to orchestrate the client-side of the application.
    It handles user input and communicates with the FastAPI server.
    """
    parser = argparse.ArgumentParser(description="AI Office Automation Pipeline. Provide your request in natural language.")
    parser.add_argument("prompt", help="Your natural language request (e.g., 'write a report on Q1 sales' or 'analyze the selected text').")
    parser.add_argument("--wps", action="store_true", help="Enable WPS functionality.")
    
    args = parser.parse_args()

    # Determine context based on mode
    if args.wps:
        print("--- AI Office Assistant Initialized (WPS Mode) ---")
        wps_doc = get_active_wps_document()
        document_context = get_wps_content(wps_doc) if wps_doc else ""
    else:
        print("--- AI Office Assistant Initialized (CLI Mode) ---")
        document_context = ""

    print(f"Request: '{args.prompt}'")
    print("-" * 37)
    
    intent = "general_prompt"
    if "cover letter" in args.prompt.lower():
        intent = "create_cover_letter"
    elif "minutes" in args.prompt.lower():
        intent = "create_minutes"
    elif "memo" in args.prompt.lower():
        intent = "create_memo"
    elif "generate document" in args.prompt.lower() or "create document" in args.prompt.lower():
        intent = "create_document"
    elif "analyze" in args.prompt.lower() or "analysis" in args.prompt.lower():
        intent = "analyze_data"
    elif "report" in args.prompt.lower():
        intent = "create_report"
    elif "summarize" in args.prompt.lower():
        intent = "summarize_content"

    try:
        payload_to_send = {}
        endpoint = ""
        is_document_generation = False # Flag to distinguish between document generation and text generation
        
        # Logic to handle specific document creation intents
        if intent == "create_cover_letter":
            is_document_generation = True
            topic = input(f"? What is the job position for this cover letter? ").strip()
            audience = input("? Who is the hiring manager or company? ").strip()
            length = input("? Desired length? (short, medium, long) [medium]: ").strip() or "medium"
            tone = input("? Desired tone? (formal, casual, technical) [formal]: ").strip() or "formal"

            payload_to_send = {
                "doc_type": "cover_letter",
                "topic": topic,
                "audience": audience,
                "length": length,
                "tone": tone
            }
            endpoint = "/create_cover_letter"

        elif intent == "create_minutes":
            is_document_generation = True
            topic = input(f"? What was the meeting topic? ").strip()
            audience = input("? Who were the meeting attendees? (e.g., Jane, John): ").strip()
            relevant_info = input("? Any key points or data to include? (e.g., Q1 sales figures, project deadlines): ").strip()
            length = input("? Desired length? (short, medium, long) [medium]: ").strip() or "medium"
            tone = input("? Desired tone? (formal, casual, technical) [formal]: ").strip() or "formal"

            payload_to_send = {
                "doc_type": "minutes",
                "topic": topic,
                "audience": audience,
                "length": length,
                "tone": tone,
                "members_present": [name.strip() for name in audience.split(',') if name.strip()],
                "data_sources": [info.strip() for info in relevant_info.split(',') if info.strip()]
            }
            endpoint = "/create_minutes"

        elif intent == "create_memo":
            is_document_generation = True
            topic = input("? What is the subject of the memo? ").strip()
            audience = input("? Who is the memo for? ").strip()
            length = input("? Desired length? (short, medium, long) [medium]: ").strip() or "medium"
            tone = input("? Desired tone? (formal, casual, technical) [formal]: ").strip() or "formal"

            payload_to_send = {
                "doc_type": "memo",
                "topic": topic,
                "audience": audience,
                "length": length,
                "tone": tone
            }
            endpoint = "/create_memo"
            
        elif intent == "create_document":
            is_document_generation = True
            doc_type = input("? What type of document? (e.g., report, blog post, manual): ").strip()
            topic = input(f"? What is the topic for the {doc_type}? (Press Enter to use your prompt): ").strip()
            audience = input("? Who is the audience? (e.g., professional, general public): ").strip()
            length = input("? Desired length? (short, medium, long) [medium]: ").strip() or "medium"
            tone = input("? Desired tone? (formal, casual, technical) [formal]: ").strip() or "formal"

            payload_to_send = {
                "doc_type": doc_type,
                "topic": topic,
                "audience": audience,
                "length": length,
                "tone": tone
            }
            endpoint = "/generate_document"

        elif intent == "analyze_data":
            payload_to_send = {
                "prompt": args.prompt,
                "content": document_context
            }
            endpoint = "/analyze"

        elif intent == "create_report":
            payload_to_send = {
                "prompt": args.prompt,
                "content": ""
            }
            endpoint = "/create_report"

        elif intent == "summarize_content":
            payload_to_send = {
                "prompt": args.prompt,
                "content": document_context
            }
            endpoint = "/summarize"
            
        else: # general_prompt
            payload_to_send = {
                "prompt": args.prompt,
                "content": document_context
            }
            endpoint = "/process"
            
        print(f"-> Sending request to FastAPI server at {API_BASE_URL}{endpoint}")
        
        response = requests.post(f"{API_BASE_URL}{endpoint}", json=payload_to_send)
        response.raise_for_status()
        
        result = response.json()
        output_content = result.get("result")
        
        if output_content:
            if args.wps:
                # For document generation, save to a new file and open it
                if is_document_generation:
                    generated_filepath = save_content_to_docx(output_content, args.prompt)
                    print(f"Generated document saved to: {os.path.abspath(generated_filepath)}")
                    print("Automatically opening the document in WPS...")
                    open_file_os_agnostic(generated_filepath)
                # For text generation, insert into the current document
                else:
                    if wps_doc:
                        insert_content_into_wps(wps_doc, output_content)
                    else:
                        print("WPS is not open or no active document. Saving to a file instead.")
                        generated_filepath = save_content_to_docx(output_content, args.prompt)
                        print(f"Generated content saved to: {os.path.abspath(generated_filepath)}")
            else:
                # CLI Mode: always save and offer to open
                generated_filepath = save_content_to_docx(output_content, args.prompt)
                print(f"Generated content saved to: {os.path.abspath(generated_filepath)}")
                
                open_file = input(f"Do you want to open the generated file now? (y/n): ").lower()
                if open_file == 'y':
                    open_file_os_agnostic(generated_filepath)
        else:
            print("Server returned no content.")

    except requests.exceptions.ConnectionError:
        print(f"\nError: Could not connect to the FastAPI server at {API_BASE_URL}.")
        print("Please ensure the backend server is running.")
    except requests.exceptions.HTTPError as e:
        print(f"\nHTTP Error from server: {e}")
        print(f"Server response: {e.response.text}")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")

if __name__ == "__main__":
    main()