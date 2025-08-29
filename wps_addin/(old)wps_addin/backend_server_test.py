
"""
This script is the backend server for the AI Office Automation add-in.
It uses FastAPI to expose endpoints that perform heavy AI and data processing tasks.
This server is intended to be run as a standalone 64-bit executable.
"""
import os
import sys
import tempfile
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import docx

# This block makes the script "bundle-aware" and correctly loads the .env file.
from dotenv import load_dotenv

if getattr(sys, 'frozen', False):
    # We are running in a bundled executable.
    # sys.executable is the path to the .exe file.
    # We want to find the .env file in the same directory as the .exe.
    application_path = os.path.dirname(sys.executable)
    dotenv_path = os.path.join(application_path, '.env')
    if os.path.exists(dotenv_path):
        load_dotenv(dotenv_path=dotenv_path)
    else:
        # This is a critical error if the .env file is missing in the bundle.
        # You can add file logging here if you want to debug this on a client's machine.
        print("FATAL: Bundled .env file not found!")
        # Create a log file in the same directory as the executable
        log_file = os.path.join(application_path, 'error.log')
        with open(log_file, 'w') as f:
            f.write("FATAL: Bundled .env file not found!\n")
            f.write(f"Looked for: {dotenv_path}\n")
            f.write(f"Files in directory: {os.listdir(application_path)}\n")
else:
    # We are running in a normal Python environment (development mode).
    # load_dotenv() will automatically find the .env in the project root.
    load_dotenv()

# Path Setup - Improved for PyInstaller
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    base_path = sys._MEIPASS
    application_path = os.path.dirname(sys.executable)
else:
    # Running in normal Python environment
    base_path = os.path.dirname(os.path.abspath(__file__))
    application_path = base_path

# Add base path to Python path
sys.path.insert(0, base_path)

# All Heavy Imports and Agent Initializations
try:
    from app.agents.llm_client import LLMClient
    from app.agents.analyzer import StructuredDataAgent
    from app.agents.reports import ReportAgent
    from app.agents.articles import ArticleAgent
    from app.agents.documents import DocumentGenerationAgent, DocumentRequest # DocumentRequest is crucial
except ImportError as e:
    error_msg = f"FATAL: Could not import agent modules. Ensure the 'app' folder is in the same directory. Error: {e}"
    print(error_msg)
    
    # Log the error for debugging
    if getattr(sys, 'frozen', False):
        log_file = os.path.join(application_path, 'import_error.log')
        with open(log_file, 'w') as f:
            f.write(error_msg + "\n")
            f.write(f"Base path: {base_path}\n")
            f.write(f"Application path: {application_path}\n")
            f.write(f"Python path: {sys.path}\n")
            if hasattr(sys, '_MEIPASS'):
                f.write(f"PyInstaller temp path: {sys._MEIPASS}\n")
                f.write(f"Files in temp path: {os.listdir(sys._MEIPASS)}\n")
    
    sys.exit(1)

print("Backend Server: Initializing AI agents...")
try:
    llm_client = LLMClient(provider="deepseek")
    report_agent = ReportAgent(llm_client)
    article_agent = ArticleAgent(llm_client)
    data_agent = StructuredDataAgent(llm_client)
    document_agent = DocumentGenerationAgent(llm_client)
    print("Backend Server: AI agents initialized successfully.")
except Exception as e:
    error_msg = f"FATAL: Failed to initialize AI agents. Check API keys in config.json. Error: {e}"
    print(error_msg)
    
    # Log the error for debugging
    if getattr(sys, 'frozen', False):
        log_file = os.path.join(application_path, 'agent_error.log')
        with open(log_file, 'w') as f:
            f.write(error_msg + "\n")
    
    sys.exit(1)

# Initialize FastAPI Server
app = FastAPI(title="AI Office Automation Backend Server")

# Define Pydantic Models for API Request Bodies
class ProcessRequest(BaseModel):
    prompt: str
    content: str = "" # Optional content field for analysis/summarization

class GeneralResponse(BaseModel):
    result: str

# Helper function to safely handle temporary files
def save_docx_and_extract_text(document_obj, prefix="temp_doc"):
    """Safely save a docx Document object and extract its text content."""
    try:
        # Use tempfile to create a proper temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx', prefix=prefix) as tmp_file:
            temp_file_path = tmp_file.name
        
        # Save the document
        document_obj.save(temp_file_path)
        
        # Read the text content
        doc = docx.Document(temp_file_path)
        full_text = "\n".join([para.text for para in doc.paragraphs])
        
        return full_text
    finally:
        # Clean up the temporary file
        try:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
        except:
            pass  # Ignore cleanup errors

# FastAPI Endpoints
@app.get("/")
def root():
    return {"message": "AI Office Backend Server is running."}

# General Document Generation Endpoint (The fallback)
@app.post("/generate_document", response_model=GeneralResponse)
def generate_document_endpoint(request: DocumentRequest):
    """Generates a general document based on a complete DocumentRequest object."""
    print(f"Backend: Received a general document request for type: '{request.doc_type}'")
    try:
        output_document_obj = document_agent.generate_document(request)
        full_text = save_docx_and_extract_text(output_document_obj, "general_doc_")

        if not full_text:
            raise HTTPException(status_code=500, detail="Failed to extract text from generated document.")
        return GeneralResponse(result=full_text)
    except Exception as e:
        print(f"Error in generate_document_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Document generation failed: {str(e)}")

# Dedicated Cover Letter Endpoint
@app.post("/create_cover_letter", response_model=GeneralResponse)
def create_cover_letter_endpoint(request: DocumentRequest):
    """Generates a cover letter from structured data."""
    print(f"Backend: Received a request to create a cover letter for: '{request.topic}'")
    try:
        request.doc_type = "cover_letter"
        output_document_obj = document_agent.generate_document(request)
        full_text = save_docx_and_extract_text(output_document_obj, "cover_letter_")
        
        return GeneralResponse(result=full_text)
    except Exception as e:
        print(f"Error in create_cover_letter_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Cover letter creation failed: {str(e)}")

# Dedicated Meeting Minutes Endpoint
@app.post("/create_minutes", response_model=GeneralResponse)
def create_minutes_endpoint(request: DocumentRequest):
    """Generates meeting minutes from structured data."""
    print(f"Backend: Received a request to create minutes for: '{request.topic}'")
    try:
        request.doc_type = "minutes"
        output_document_obj = document_agent.generate_document(request)
        full_text = save_docx_and_extract_text(output_document_obj, "minutes_")
        
        return GeneralResponse(result=full_text)
    except Exception as e:
        print(f"Error in create_minutes_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Meeting minutes creation failed: {str(e)}")

# Dedicated Memo Endpoint
@app.post("/create_memo", response_model=GeneralResponse)
def create_memo_endpoint(request: DocumentRequest):
    """Generates a memorandum from structured data."""
    print(f"Backend: Received a request to create a memo on topic: '{request.topic}'")
    try:
        request.doc_type = "memo"
        output_document_obj = document_agent.generate_document(request)
        full_text = save_docx_and_extract_text(output_document_obj, "memo_")
        
        return GeneralResponse(result=full_text)
    except Exception as e:
        print(f"Error in create_memo_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Memo creation failed: {str(e)}")
    
# Dedicated endpoint for Report Generation
@app.post("/create_report", response_model=GeneralResponse)
def create_report_endpoint(request: ProcessRequest):
    """Creates a report based on the provided prompt."""
    print("Backend: Received request to create a report.")
    try:
        # Assuming ReportAgent.create_report_content expects a 'topic'
        # The prompt from ProcessRequest is used as the topic.
        output_content = report_agent.create_report(
            topic=request.prompt, 
            tone="professional", # Default tone
            length="standard"    # Default length
        )
        if not output_content:
            raise HTTPException(status_code=500, detail="Failed to generate report content.")
        return GeneralResponse(result=output_content)
    except Exception as e:
        print(f"Error in create_report_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")

# Dedicated endpoint for Data Analysis
@app.post("/analyze", response_model=GeneralResponse)
def analyze_endpoint(request: ProcessRequest):
    """Analyzes provided text content."""
    print("Backend: Received request to analyze content.")
    try:
        if not request.content:
            raise HTTPException(status_code=400, detail="No content provided for analysis.")
        
        # The data_agent.analyze_input_content generates the analysis text directly
        output_content = data_agent.analyze_input(raw_input=request.content, user_question=request.prompt)
        
        if not output_content:
            raise HTTPException(status_code=500, detail="Failed to generate analysis content.")
        return GeneralResponse(result=output_content)
    except Exception as e:
        print(f"Error in analyze_endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Analysis failed: {str(e)}")

# Dedicated endpoint for Summarization
@app.post("/summarize", response_model=GeneralResponse)
def summarize_endpoint(request: ProcessRequest):
    """Receives text content and returns a summary."""
    print("Backend: Received request to summarize document.")
    try:
        if not request.content:
            raise HTTPException(status_code=400, detail="No content provided for summarization.")
            
        # Corrected: Using generate_response instead of get_completion
        summary = llm_client.generate_response(
            f"Please provide a concise summary of the following document:\n\n{request.content}"
        )
        if not summary:
            raise HTTPException(status_code=500, detail="Failed to generate summary.")
        return GeneralResponse(result=summary)
    except Exception as e:
        print(f"Error during summarization: {e}")
        raise HTTPException(status_code=500, detail=f"Summarization failed: {str(e)}")

# General fallback endpoint for unstructured prompts
@app.post("/process", response_model=GeneralResponse)
def process_general_prompt(request: ProcessRequest):
    """Handles general prompts that don't fit other dedicated endpoints."""
    print(f"Backend: Received general prompt: '{request.prompt}'. Content present: {len(request.content) > 0}")
    try:
        # Corrected: Using generate_response instead of get_completion
        output_content = llm_client.generate_response(request.prompt)
        if not output_content:
            raise HTTPException(status_code=500, detail="Failed to get completion for general prompt.")
        return GeneralResponse(result=output_content)
    except Exception as e:
        print(f"Error in general process endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"General prompt processing failed: {str(e)}")

# This block allows running the server directly for development
if __name__ == "__main__":
    import uvicorn
    print("Starting backend server...")
    
    # Determine if we're in development mode
    reload = os.getenv("ENV") == "development" and not getattr(sys, 'frozen', False)
    
    # Setup logging configuration
    log_config = {
        "version": 1, 
        "disable_existing_loggers": False,
        "formatters": {
            "default": {
                "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            },
        },
        "handlers": {
            "console": {
                "class": "logging.StreamHandler",
                "formatter": "default",
            },
            "file": {
                "class": "logging.FileHandler",
                "filename": os.path.join(application_path, "server.log"),
                "formatter": "default",
            }
        },
        "root": {
            "level": "INFO",
            "handlers": ["console", "file"],
        }
    }
    
    try:
        uvicorn.run(
            app, 
            host="127.0.0.1", 
            port=8000, 
            reload=reload,
            log_config=log_config
        )
    except Exception as e:
        error_msg = f"Failed to start server: {e}"
        print(error_msg)
        if getattr(sys, 'frozen', False):
            log_file = os.path.join(application_path, 'server_startup_error.log')
            with open(log_file, 'w') as f:
                f.write(error_msg + "\n")