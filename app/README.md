 
# AI-Powered Word Document Automation Agent

## Overview

This project is a Python-based automation agent designed to streamline the creation of professional Word documents (`.docx`). It intelligently combines AI-powered content generation with robust data analysis and visualization capabilities to produce complex documents like reports, articles, and memos with minimal human intervention.

The agent is built with a modular, object-oriented architecture, making it easy to understand, maintain, and extend.

## Key Features

-   **Modular AI Content Generation**: Utilizes large language models (like OpenAI's GPT series) to generate structured content sections (e.g., titles, executive summaries, conclusions) based on user-defined requests.
-   **Integrated Data Analysis**: Reads data from various sources (`.csv`, `.xlsx`, `.json`) and performs automated analysis, including summary statistics, insight generation, and actionable recommendations.
-   **Automated Data Visualization**: Creates and embeds relevant charts and graphs (e.g., histograms, heatmaps) directly into the generated documents to support the data analysis.
-   **Structured & Configurable Requests**: Uses a clear `DocumentRequest` structure to define the desired output, allowing for precise control over the document type, topic, tone, length, and audience.
-   **Automated DOCX Document Assembly**: Intelligently builds professional, well-formatted Word documents using the generated content and visualizations, applying custom styles for headings and paragraphs.
-   **Batch Processing**: Capable of processing multiple document requests in a single run, making it ideal for large-scale content creation tasks.

## Architecture

The system is designed around several core components that work together to process a request:

1.  **`WordAutomationAgent`**: The main orchestrator. It receives a `DocumentRequest`, coordinates the other components, and manages the end-to-end workflow.
2.  **`AIContentGenerator`**: The "writer". This class interfaces with an AI model (e.g., OpenAI) to generate the textual content for each section of the document.
3.  **`DataAnalyzer`**: The "analyst". This class is responsible for loading data from files, performing statistical analysis using `pandas`, and creating visualizations with `matplotlib` and `seaborn`.
4.  **`WordDocumentBuilder`**: The "publisher". This class takes all the generated text and charts and assembles them into a final, formatted `.docx` file using `python-docx`.
5.  **`DocumentRequest`**: A dataclass that serves as a structured "brief" for the agent, defining exactly what needs to be created.

## Project Strucrure
```bash
ai-powered-office-automation/
├── app/
│   ├── __init__.py  <-- IMPORTANT: Makes 'app' a package
│   ├── analyzer.py
│   ├── articles.py
│   ├── documents.py
│   ├── llm_client.py
│   ├── reports.py
│   └── main.py      <-- Your original CLI entry point
│
├── api/
│   ├── __init__.py  <-- IMPORTANT: Makes 'api' a package
│   └── server.py    <-- Your new Flask API
│
├── .env
├── requirements.txt
└── generated_documents/ 
```

## Setup and Installation

### Prerequisites

-   Python 3.8 or higher
-   `pip` (Python package installer)
-   An OpenAI API Key (or another supported AI model provider)

### Installation

1.  **Clone the repository (or save the script):**
    ```bash
    git clone <your-repo-url>
    cd <your-repo-directory>
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv env
    source env/bin/activate  # On Windows, use `env\Scripts\activate`
    ```

3.  **Install the required libraries:**
    Create a file named `requirements.txt` with the following content:

    ```text
    pandas
    matplotlib
    seaborn
    openai
    python-docx
    openpyxl  # Required for .xlsx file support
    pywin32   # For win32com, if extended Windows features are used
    ```

    Then, install them using pip:
    ```bash
    pip install -r requirements.txt
    ```
    **Note**: The `win32com.client` import suggests this was designed with Windows in mind. The core DOCX generation is cross-platform, but any direct COM automation would be Windows-only.

## Configuration

Before running the agent, you must configure your AI provider's API key.

Open the script and locate the `main()` function. Replace the placeholder with your actual key:

```python
def main():
    """Example usage of the Word Automation Agent"""

    # Initialize the agent with your API key
    agent = WordAutomationAgent(ai_api_key="sk-YourActualOpenAI_API_KeyHere")

    # ... rest of the function

```
The Final Workflow
The user navigates to your web application in their browser.
They click a "Connect to Microsoft 365" link, go through the Microsoft login and consent flow, and are redirected back. Your app now has a temporary access token for them.
The user interacts with your frontend, perhaps selecting a file from their OneDrive (which you can list using the Graph API) or uploading a new one.
They click a button like "Generate Report." This sends a request to your /generate_report endpoint.
Your Flask server receives the request, uses the appropriate agent to do the work, and then calls the document_generator to create a .docx file in memory.
Flask sends this file back to the user's browser, triggering a download prompt.
The user saves the analysis_report.docx and opens it with their local installation of WPS Office or Microsoft Office.