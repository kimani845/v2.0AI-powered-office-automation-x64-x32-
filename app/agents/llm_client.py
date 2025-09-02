import os
import re
import json
import sys
from typing import Dict, Any, Optional
import requests
import google.generativeai as genai
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# Load API Keys from environment
# OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
# GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
# OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://localhost:11434")
# DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY") # DeepSeek API Key


# System Prompt for Intent Parsing 
SYSTEM_PROMPT_PARSER = """
You are a master task-routing assistant. Your job is to analyze a user's request and convert it into a structured JSON command.
The JSON object must have an "action" key and a "params" key.
- action: Must be one of ["analyse_data", "create_report", "write_article", create_document', "unknown"].
- params: 
    -A dictionary of details.
    - doc_type (string, required): One of 'cover_letter', 'minutes', 'memo'.
    - topic (string, required): The subject of the document.
    - audience (string, required): The recipient of the document.
    - tone (string, optional): e.g., 'formal', 'casual'.
    - length (string, optional): e.g., 'short', 'medium'.
    - data_sources (list[string], optional): Key points to include.
        

Example 1: "Can you analyze the sales_data.csv file for me?"
Output:
{
    "action": "analyse_data",
    "params": {
        "file": "sales_data.csv"
    }
}

Example 2: "draft a formal quarterly report about Q3 sales performance"
Output:
{
    "action": "create_report",
    "params": {
        "topic": "Q3 sales performance",
        "tone": "formal",
        "type": "quarterly"
    }
}

Example 3: "write me a blog post about the future of AI in healthcare"
Output:
{
    "action": "write_article",
    "params": {
        "topic": "The Future of AI in Healthcare",
        "style": "blog post"
    }
}

Example 4: "Summarize our performance"
Output:
{
    "action": "analyse_data",
    "params": {}
}

Example 5: "Draft meeting minutes for the Q3 Strategy Meeting with the Executive Team."
Output: {
    "action": "create_document",
    "params": {
        "doc_type": "minutes",
        "topic": "Q3 Strategy Meeting",
        "audience": "Executive Team"
    }
}

Output ONLY the JSON object and nothing else.
"""
# ---------UPDATED CODE----------------------------------------------------------------------------------
    # new, bundler-friendly way
    # ... inside your LLMClient __init__ method ...
def get_api_key(provider: str) -> str:
    try:
            # This logic helps PyInstaller find the config file
        if getattr(sys, 'frozen', False):
            # we are running in a bundle
            base_path = sys._MEIPASS
        else:
            # we are running in a normal Python environment
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # Project root
            
            config_path = os.path.join(base_path, 'config.json')
            
#         with open(config_path, 'r') as f:
#                 config = json.load(f)
#         return config.get(f"{provider.upper()}_API_KEY")
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                config = json.load(f)
                key = config.get(f"{provider.upper()}_API_KEY")
                if key:
                    return key
                
    except (FileNotFoundError, KeyError):
        print(f"Error: API key for {provider.upper()} not found in config.json")
        return None
    # Fallback for to environment Variables
    env_key = os.getenv(f"{provider.upper()}_API_KEY")
    if env_key:
        return env_key

    print(f"Error: API key for {provider.upper()} not found in config.json or environment variables.")
    return None
# # --------------------------------------------------------------------------------------------------------

class LLMClient:
    """A unified client to interact with multiple LLM providers."""
    # def __init__(self, provider: str = "openai", model: str = None):
    def __init__(self, provider: str = "deepseek", model: str = None):
        """
        provider: 'ollama', 'gemini', 'openai', 'deepseek'
        model: Model name which will depend on the provider.
        """
        self.provider = provider.lower()
        self._validate_provider()
        # self.model = model or self._default_model()
        # self.client = self._setup_client()
        self.api_key = get_api_key(self.provider)
        if not self.api_key:
            raise ValueError(f"API Key for provider '{self.provider}' is missing.")
        self.model = model or self._default_model()
        self.client = self._setup_client()


    def _default_model(self) -> str:
        defaults = {
            "openai": "gpt-4o-mini",
            "gemini": "gemini-1.5-flash",
            "ollama": "mistral",
            # "deepseek": "deepseek-chat"  # DeepSeek default model
            "deepseek": "deepseek/deepseek-r1-0528-qwen3-8b:free"
        }

        # return {"openai": "gpt-4o-mini", "gemini": "gemini-1.5-flash", "ollama": "mistral"}.get(self.provider)
        return defaults.get(self.provider, "deepseek-chat")  # Default fallback if provider is unknown


    def _validate_provider(self):
        supported = ["ollama", "openai", "gemini", "deepseek"]
        if self.provider not in supported:
            raise ValueError(f"Provider '{self.provider}' is not supported. Choose from {supported}.")

    # def _setup_client(self) -> Any:
    #     if self.provider == "openai":
    #         if not OPENAI_API_KEY: raise ValueError("OPENAI_API_KEY not set.")
    #         return OpenAI(api_key=OPENAI_API_KEY)
    #     elif self.provider == "gemini":
    #         if not GEMINI_API_KEY: raise ValueError("GEMINI_API_KEY not set.")
    #         genai.configure(api_key=GEMINI_API_KEY)
    #         return genai.GenerativeModel(self.model)
    #     # Ollama and DeepSeek typically use direct requests, so no client object is returned here.
    #     return 
    def _setup_client(self) -> Any:
        if self.provider == "openai":
            return OpenAI(api_key=self.api_key)
        elif self.provider == "gemini":
            genai.configure(api_key=self.api_key)
            return genai.GenerativeModel(self.model)
        # Ollama and DeepSeek use direct requests, so no client object is returned here.
        return None

    def generate_response(self, prompt: str, system_prompt: Optional[str] = None, json_mode: bool = False) -> str:
        dispatch = {
            "openai": self._call_openai, 
            "gemini": self._call_gemini, 
            "ollama": self._call_ollama,
            "deepseek": self._call_deepseek # DeepSeek dispatch
            
            }
        return dispatch[self.provider](prompt, system_prompt, json_mode)

    def _call_openai(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
        messages = [{"role": "system", "content": system_prompt}] if system_prompt else []
        messages.append({"role": "user", "content": prompt})
        response_format = {"type": "json_object"} if json_mode else {"type": "text"}
        try:
            resp = self.client.chat.completions.create(
                model=self.model, 
                messages=messages, 
                temperature=0.1,
                response_format=response_format
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            raise RuntimeError(f"OpenAI API request failed: {e}")

    def _call_gemini(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
        full_prompt = f"{system_prompt}\n\n{prompt}" if system_prompt else prompt
        config = {"response_mime_type": "application/json"} if json_mode else {}
        try:
            resp = self.client.generate_content(
                full_prompt, 
                generation_config=genai.types.GenerationConfig(**config))
            return resp.text.strip()
        except Exception as e:
            raise RuntimeError(f"Gemini API request failed: {e}")

    def _call_ollama(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
        messages = [{"role": "system", "content": system_prompt}] if system_prompt else []
        messages.append({"role": "user", "content": prompt})
        payload = {
            "model": self.model, 
            "messages": messages, 
            "stream": False, 
            "format": "json" if json_mode else ""
        }
        try:
            resp = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload, timeout=60.0)
            resp.raise_for_status()
            return resp.json()["message"]["content"].strip()
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"Ollama API request failed: {e}")


    def _call_deepseek(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
        """Calls the DeepSeek API."""
        # if not DEEPSEEK_API_KEY:
        #     raise ValueError("DEEPSEEK_API_KEY environment variable is not set.")

        if not self.api_key:
            raise ValueError("DeepSeek API Key is not set.")

        
        headers = {
            "Content-Type": "application/json",
            # "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
            "Authorization": f"Bearer {self.api_key}"

        }

        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        # Always add user message
        messages.append({"role": "user", "content": prompt})

        payload = {
            "model": self.model,
            "messages": messages,
            "stream": False
        }

        try:
            response = requests.post(
                # "https://api.deepseek.com/chat/completions",
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=60
            )
            response.raise_for_status()
            data = response.json()

            if data and "choices" in data and len(data["choices"]) > 0:
                return data["choices"][0]["message"]["content"].strip()
            else:
                raise RuntimeError(f"DeepSeek API returned unexpected response format: {data}")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"DeepSeek API request failed: {e}")



    def parse_instruction(self, instruction: str) -> Dict[str, Any]:
        """Parses a natural language instruction into a structured JSON action."""
        try:
            response_text = self.generate_response(prompt=instruction, system_prompt=SYSTEM_PROMPT_PARSER, json_mode=True)
            match = re.search(r'```(json)?\s*({.*})\s*```', response_text, re.DOTALL)
            if match: response_text = match.group(2)
            parsed_json = json.loads(response_text)
            if "action" not in parsed_json: raise ValueError("Response missing 'action' key")
            return parsed_json
        except (json.JSONDecodeError, ValueError, RuntimeError) as e:
            print(f"Warning: LLM JSON parsing failed ({e}). Falling back to basic keywords.")
            lower_instruction = instruction.lower()
            if "analyze" in lower_instruction or "analysis" in lower_instruction: return {"action": "analyse_data", "params": {}}
            if "report" in lower_instruction: return {"action": "create_report", "params": {"topic": instruction}}
            if "article" in lower_instruction or "blog" in lower_instruction: return {"action": "write_article", "params": {"topic": instruction}}
            if "document" in lower_instruction or "memo" in lower_instruction: return {"action": "create_document", "params": {"topic": instruction}}

            return {"action": "unknown", "params": {}}



    # def parse_instruction(self, instruction: str) -> Dict[str, Any]:
    #     """Parses a natural language instruction into a structured JSON action."""
    #     try:
    #         # SYSTEM_PROMPT_PARSER is the system prompt for intent parsing
    #         response_text = self.generate_response(prompt=instruction, system_prompt=SYSTEM_PROMPT_PARSER, json_mode=True)
            
    #         # Clean JSON response by removing markdown fences (```json or ```)
    #         match = re.search(r'```(?:json)?\s*({.*})\s*```', response_text, re.DOTALL)
    #         if match: 
    #             response_text = match.group(1).strip() # Extract only the JSON content
    #         else:
    #             # If no markdown fences, try to find a JSON object directly
    #             match = re.search(r'({.*})', response_text.strip(), re.DOTALL)
    #             if match:
    #                 response_text = match.group(1).strip()
            
    #         parsed_json = json.loads(response_text)
    #         if "action" not in parsed_json: 
    #             raise ValueError("LLM response missing 'action' key.")
    #         return parsed_json
    #     except (json.JSONDecodeError, ValueError, RuntimeError) as e:
    #         print(f"Warning: LLM JSON parsing failed ({e}). Falling back to basic keywords.")
    #         # Fallback logic in case of JSON parsing failure
    #         lower_instruction = instruction.lower()
    #         action = "unknown"
    #         params = {}

    #         if "analyze" in lower_instruction or "analysis" in lower_instruction: 
    #             action = "analyse_data"
    #             file_matches = re.findall(r'(\S+\.(?:csv|xlsx|xls|json))', lower_instruction)
    #             if file_matches:
    #                 params['file'] = file_matches[0]
    #         elif "report" in lower_instruction: 
    #             action = "create_report"
    #             match = re.search(r'report about (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback to full instruction
    #         elif "article" in lower_instruction or "blog" in lower_instruction: 
    #             action = "write_article"
    #             match = re.search(r'(?:article|blog post) about (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback
    #         elif "document" in lower_instruction or "memo" in lower_instruction or "letter" in lower_instruction or "minutes" in lower_instruction: 
    #             action = "create_document"
    #             if "cover letter" in lower_instruction: params["doc_type"] = "cover_letter"
    #             elif "meeting minutes" in lower_instruction: params["doc_type"] = "minutes"
    #             elif "memo" in lower_instruction: params["doc_type"] = "memo"
    #             match = re.search(r'(?:document|memo|letter|minutes)(?: about)? (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback

    #         return {"action": action, "params": params}
    

# import os
# import re
# import json
# import sys
# from typing import Dict, Any, Optional
# import requests
# import google.generativeai as genai
# from openai import OpenAI
# from dotenv import load_dotenv

# load_dotenv()

# # Load API Keys from environment
# # OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
# # GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
# # OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://localhost:11434")
# # DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY") # DeepSeek API Key


# # System Prompt for Intent Parsing 
# SYSTEM_PROMPT_PARSER = """
# You are a master task-routing assistant. Your job is to analyze a user's request and convert it into a structured JSON command.
# The JSON object must have an "action" key and a "params" key.
# - action: Must be one of ["analyse_data", "create_report", "write_article", create_document', "unknown"].
# - params: 
#     -A dictionary of details.
#     - doc_type (string, required): One of 'cover_letter', 'minutes', 'memo'.
#     - topic (string, required): The subject of the document.
#     - audience (string, required): The recipient of the document.
#     - tone (string, optional): e.g., 'formal', 'casual'.
#     - length (string, optional): e.g., 'short', 'medium'.
#     - data_sources (list[string], optional): Key points to include.
        

# Example 1: "Can you analyze the sales_data.csv file for me?"
# Output:
# {
#     "action": "analyse_data",
#     "params": {
#         "file": "sales_data.csv"
#     }
# }

# Example 2: "draft a formal quarterly report about Q3 sales performance"
# Output:
# {
#     "action": "create_report",
#     "params": {
#         "topic": "Q3 sales performance",
#         "tone": "formal",
#         "type": "quarterly"
#     }
# }

# Example 3: "write me a blog post about the future of AI in healthcare"
# Output:
# {
#     "action": "write_article",
#     "params": {
#         "topic": "The Future of AI in Healthcare",
#         "style": "blog post"
#     }
# }

# Example 4: "Summarize our performance"
# Output:
# {
#     "action": "analyse_data",
#     "params": {}
# }

# Example 5: "Draft meeting minutes for the Q3 Strategy Meeting with the Executive Team."
# Output: {
#     "action": "create_document",
#     "params": {
#         "doc_type": "minutes",
#         "topic": "Q3 Strategy Meeting",
#         "audience": "Executive Team"
#     }
# }

# Output ONLY the JSON object and nothing else.
# """
# # ---------UPDATED CODE----------------------------------------------------------------------------------
# def get_api_key(provider: str) -> Optional[str]:
#     """
#     Safely retrieves an API key, first from a config.json file (for bundled apps)
#     and then falling back to environment variables.
#     """
#     # 1. Determine the base path for finding files, works for both bundled and normal modes.
#     if getattr(sys, 'frozen', False):
#         # We are running in a PyInstaller bundle.
#         base_path = sys._MEIPASS
#     else:
#         # We are running in a normal Python environment.
#         # The project root is two directories up from this file (app/agents/llm_client.py)
#         base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#     # 2. Construct the full path to the configuration file.
#     config_path = os.path.join(base_path, 'config.json')

#     # 3. Try to load the key from the config.json file.
#     try:
#         if os.path.exists(config_path):
#             with open(config_path, 'r') as f:
#                 config = json.load(f)
            
#             # Assuming config format is {"DEEPSEEK_API_KEY": "...", "OPENAI_API_KEY": "..."}
#             key = config.get(f"{provider.upper()}_API_KEY")
#             if key:
#                 return key
#     except Exception as e:
#         print(f"Warning: Could not read or parse config.json. Error: {e}")

#     # 4. If not found in config.json, fall back to environment variables.
#     # This also handles the original load_dotenv() approach.
#     env_key = os.getenv(f"{provider.upper()}_API_KEY")
#     if env_key:
#         return env_key

#     # 5. If no key is found anywhere, return None.
#     print(f"Error: API key for {provider.upper()} not found in config.json or environment variables.")
#     return None
# # # --------------------------------------------------------------------------------------------------------

# class LLMClient:
#     """A unified client to interact with multiple LLM providers."""
#     # def __init__(self, provider: str = "openai", model: str = None):
#     def __init__(self, provider: str = "deepseek", model: str = None):
#         """
#         provider: 'ollama', 'gemini', 'openai', 'deepseek'
#         model: Model name which will depend on the provider.
#         """
#         self.provider = provider.lower()
#         self._validate_provider()
#         # self.model = model or self._default_model()
#         # self.client = self._setup_client()
#         self.api_key = get_api_key(self.provider)
#         if not self.api_key:
#             raise ValueError(f"API Key for provider '{self.provider}' is missing.")
#         self.model = model or self._default_model()
#         self.client = self._setup_client()


#     def _default_model(self) -> str:
#         defaults = {
#             "openai": "gpt-4o-mini",
#             "gemini": "gemini-1.5-flash",
#             "ollama": "mistral",
#             # "deepseek": "deepseek-chat"  # DeepSeek default model
#             "deepseek": "deepseek/deepseek-r1-0528-qwen3-8b:free"
#         }

#         # return {"openai": "gpt-4o-mini", "gemini": "gemini-1.5-flash", "ollama": "mistral"}.get(self.provider)
#         return defaults.get(self.provider, "deepseek-chat")  # Default fallback if provider is unknown


#     def _validate_provider(self):
#         supported = ["ollama", "openai", "gemini", "deepseek"]
#         if self.provider not in supported:
#             raise ValueError(f"Provider '{self.provider}' is not supported. Choose from {supported}.")

#     # def _setup_client(self) -> Any:
#     #     if self.provider == "openai":
#     #         if not OPENAI_API_KEY: raise ValueError("OPENAI_API_KEY not set.")
#     #         return OpenAI(api_key=OPENAI_API_KEY)
#     #     elif self.provider == "gemini":
#     #         if not GEMINI_API_KEY: raise ValueError("GEMINI_API_KEY not set.")
#     #         genai.configure(api_key=GEMINI_API_KEY)
#     #         return genai.GenerativeModel(self.model)
#     #     # Ollama and DeepSeek typically use direct requests, so no client object is returned here.
#     #     return 
#     def _setup_client(self) -> Any:
#         if self.provider == "openai":
#             return OpenAI(api_key=self.api_key)
#         elif self.provider == "gemini":
#             genai.configure(api_key=self.api_key)
#             return genai.GenerativeModel(self.model)
#         # Ollama and DeepSeek use direct requests, so no client object is returned here.
#         return None

#     def generate_response(self, prompt: str, system_prompt: Optional[str] = None, json_mode: bool = False) -> str:
#         dispatch = {
#             "openai": self._call_openai, 
#             "gemini": self._call_gemini, 
#             "ollama": self._call_ollama,
#             "deepseek": self._call_deepseek # DeepSeek dispatch
            
#             }
#         return dispatch[self.provider](prompt, system_prompt, json_mode)

#     def _call_openai(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
#         messages = [{"role": "system", "content": system_prompt}] if system_prompt else []
#         messages.append({"role": "user", "content": prompt})
#         response_format = {"type": "json_object"} if json_mode else {"type": "text"}
#         try:
#             resp = self.client.chat.completions.create(
#                 model=self.model, 
#                 messages=messages, 
#                 temperature=0.1,
#                 response_format=response_format
#             )
#             return resp.choices[0].message.content.strip()
#         except Exception as e:
#             raise RuntimeError(f"OpenAI API request failed: {e}")

#     def _call_gemini(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
#         full_prompt = f"{system_prompt}\n\n{prompt}" if system_prompt else prompt
#         config = {"response_mime_type": "application/json"} if json_mode else {}
#         try:
#             resp = self.client.generate_content(
#                 full_prompt, 
#                 generation_config=genai.types.GenerationConfig(**config))
#             return resp.text.strip()
#         except Exception as e:
#             raise RuntimeError(f"Gemini API request failed: {e}")

#     def _call_ollama(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
#         messages = [{"role": "system", "content": system_prompt}] if system_prompt else []
#         messages.append({"role": "user", "content": prompt})
#         payload = {
#             "model": self.model, 
#             "messages": messages, 
#             "stream": False, 
#             "format": "json" if json_mode else ""
#         }
#         try:
#             resp = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload, timeout=60.0)
#             resp.raise_for_status()
#             return resp.json()["message"]["content"].strip()
#         except requests.exceptions.RequestException as e:
#             raise RuntimeError(f"Ollama API request failed: {e}")


#     def _call_deepseek(self, prompt: str, system_prompt: Optional[str], json_mode: bool) -> str:
#         """Calls the DeepSeek API."""
#         # if not DEEPSEEK_API_KEY:
#         #     raise ValueError("DEEPSEEK_API_KEY environment variable is not set.")

#         if not self.api_key:
#             raise ValueError("DeepSeek API Key is not set.")

        
#         headers = {
#             "Content-Type": "application/json",
#             # "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
#             "Authorization": f"Bearer {self.api_key}"

#         }

#         messages = []
#         if system_prompt:
#             messages.append({"role": "system", "content": system_prompt})
#         # Always add user message
#         messages.append({"role": "user", "content": prompt})

#         payload = {
#             "model": self.model,
#             "messages": messages,
#             "stream": False
#         }

#         try:
#             response = requests.post(
#                 # "https://api.deepseek.com/chat/completions",
#                 "https://openrouter.ai/api/v1/chat/completions",
#                 headers=headers,
#                 json=payload,
#                 timeout=60
#             )
#             response.raise_for_status()
#             data = response.json()

#             if data and "choices" in data and len(data["choices"]) > 0:
#                 return data["choices"][0]["message"]["content"].strip()
#             else:
#                 raise RuntimeError(f"DeepSeek API returned unexpected response format: {data}")
#         except requests.exceptions.RequestException as e:
#             raise RuntimeError(f"DeepSeek API request failed: {e}")



#     def parse_instruction(self, instruction: str) -> Dict[str, Any]:
#         """Parses a natural language instruction into a structured JSON action."""
#         try:
#             response_text = self.generate_response(prompt=instruction, system_prompt=SYSTEM_PROMPT_PARSER, json_mode=True)
#             match = re.search(r'```(json)?\s*({.*})\s*```', response_text, re.DOTALL)
#             if match: response_text = match.group(2)
#             parsed_json = json.loads(response_text)
#             if "action" not in parsed_json: raise ValueError("Response missing 'action' key")
#             return parsed_json
#         except (json.JSONDecodeError, ValueError, RuntimeError) as e:
#             print(f"Warning: LLM JSON parsing failed ({e}). Falling back to basic keywords.")
#             lower_instruction = instruction.lower()
#             if "analyze" in lower_instruction or "analysis" in lower_instruction: return {"action": "analyse_data", "params": {}}
#             if "report" in lower_instruction: return {"action": "create_report", "params": {"topic": instruction}}
#             if "article" in lower_instruction or "blog" in lower_instruction: return {"action": "write_article", "params": {"topic": instruction}}
#             if "document" in lower_instruction or "memo" in lower_instruction: return {"action": "create_document", "params": {"topic": instruction}}

#             return {"action": "unknown", "params": {}}



    # def parse_instruction(self, instruction: str) -> Dict[str, Any]:
    #     """Parses a natural language instruction into a structured JSON action."""
    #     try:
    #         # SYSTEM_PROMPT_PARSER is the system prompt for intent parsing
    #         response_text = self.generate_response(prompt=instruction, system_prompt=SYSTEM_PROMPT_PARSER, json_mode=True)
            
    #         # Clean JSON response by removing markdown fences (```json or ```)
    #         match = re.search(r'```(?:json)?\s*({.*})\s*```', response_text, re.DOTALL)
    #         if match: 
    #             response_text = match.group(1).strip() # Extract only the JSON content
    #         else:
    #             # If no markdown fences, try to find a JSON object directly
    #             match = re.search(r'({.*})', response_text.strip(), re.DOTALL)
    #             if match:
    #                 response_text = match.group(1).strip()
            
    #         parsed_json = json.loads(response_text)
    #         if "action" not in parsed_json: 
    #             raise ValueError("LLM response missing 'action' key.")
    #         return parsed_json
    #     except (json.JSONDecodeError, ValueError, RuntimeError) as e:
    #         print(f"Warning: LLM JSON parsing failed ({e}). Falling back to basic keywords.")
    #         # Fallback logic in case of JSON parsing failure
    #         lower_instruction = instruction.lower()
    #         action = "unknown"
    #         params = {}

    #         if "analyze" in lower_instruction or "analysis" in lower_instruction: 
    #             action = "analyse_data"
    #             file_matches = re.findall(r'(\S+\.(?:csv|xlsx|xls|json))', lower_instruction)
    #             if file_matches:
    #                 params['file'] = file_matches[0]
    #         elif "report" in lower_instruction: 
    #             action = "create_report"
    #             match = re.search(r'report about (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback to full instruction
    #         elif "article" in lower_instruction or "blog" in lower_instruction: 
    #             action = "write_article"
    #             match = re.search(r'(?:article|blog post) about (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback
    #         elif "document" in lower_instruction or "memo" in lower_instruction or "letter" in lower_instruction or "minutes" in lower_instruction: 
    #             action = "create_document"
    #             if "cover letter" in lower_instruction: params["doc_type"] = "cover_letter"
    #             elif "meeting minutes" in lower_instruction: params["doc_type"] = "minutes"
    #             elif "memo" in lower_instruction: params["doc_type"] = "memo"
    #             match = re.search(r'(?:document|memo|letter|minutes)(?: about)? (.+)', lower_instruction)
    #             if match: params["topic"] = match.group(1).strip()
    #             else: params["topic"] = instruction # Fallback

    #         return {"action": action, "params": params}