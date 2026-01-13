import logging
import traceback
import warnings
import csv
import shutil
import zipfile
from pathlib import Path
from urllib3.exceptions import InsecureRequestWarning

import PyPDF2
import pandas as pd
import requests
from docx2pdf import convert
from fpdf import FPDF
from PyPDF2 import PdfReader

# Note: The following imports are commented out or require additional installation
# import base64
# import img2pdf
# import pythoncom
# import win32com.client
# from src.docling.docling_custom import md_file_create

# Suppress insecure request warnings
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

# Configure logging
import sys
sys.path.append("..")
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
from logger.logger import get_logger

logger = get_logger()
print(f"#### Current working directory: {os.getcwd()}")

JSON_BLOCK_RE = re.compile(r"\{[\s\S]*\}", re.MULTILINE)

# Accepted file extensions for processing
ACCEPTED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.pdf', '.docx', '.zip'}
ACCEPTED_NONZIP = ACCEPTED_EXTENSIONS - {'.zip'}


def extract_json_block(text: str) -> str | None:
    """Extract the first JSON object from a string (best-effort approach)"""
    match = JSON_BLOCK_RE.search(text.strip())
    return match.group(0) if match else None


def coerce_to_json(text: str) -> dict:
    """Attempt to parse JSON directly; if it fails, try extracting a JSON block first"""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        block = extract_json_block(text)
        if not block:
            raise ValueError("No JSON block found in LLM output")
        return json.loads(block)


def to_pascal_case_with_underscore(s: str) -> str:
    """Convert string to PascalCase with underscores"""
    s = re.sub(r'[^a-zA-Z0-9 ]', '', s)
    words = s.strip().split()
    return '_'.join(word.capitalize() for word in words)


def clean_column_name(name: str) -> str:
    """Clean and standardize column names"""
    return to_pascal_case_with_underscore(name)


def is_protected_or_corrupted(data: dict) -> bool:
    """Check if document is protected or corrupted"""
    return (
        list(data.keys()) == ["DOCUMENT CATEGORY"] and
        data["DOCUMENT CATEGORY"] in ["Document Protected", "Corrupted"]
    )


def process_json_file(json_path: str, contract_id: str, file_name: str) -> dict | None:
    """Process a single JSON file and extract relevant data"""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    if is_protected_or_corrupted(data):
        return None
    
    row = {}
    for key, value in data.items():
        col_name = clean_column_name(key)
        if col_name == 'Document_Details':
            for section, section_val in value.items():
                if isinstance(section_val, dict):
                    for subkey, subval in section_val.items():
                        sub_col = clean_column_name(subkey)
                        norm_val = subval.get('normalised value', '') if isinstance(subval, dict) else ''
                        row[sub_col] = norm_val
        else:
            if col_name in ['Document_Reference_Number', 'Parent_Reference_Number']:
                row[col_name] = re.sub(r'\s+', '', value)
            else:
                row[col_name] = value
    
    row['Contract_Reference_Number'] = contract_id
    row['File_Name'] = file_name
    return row


def get_existing_contract_ids_and_rows(csv_path: str) -> tuple[set, list, list]:
    """Get existing contract IDs and rows from CSV file"""
    contract_ids = set()
    existing_rows = []
    columns = []
    
    if os.path.exists(csv_path):
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            columns = reader.fieldnames or []
            for row in reader:
                contract_ids.add(row.get('Contract_Reference_Number', ''))
                existing_rows.append(row)
    
    return contract_ids, existing_rows, columns


def build_csv_from_scratch(inp_json_dir: str, csv_path: str):
    """Build CSV file from all JSON files in directory"""
    rows = []
    all_columns = set()
    
    for contract_id in os.listdir(inp_json_dir):
        contract_path = os.path.join(inp_json_dir, contract_id)
        if not os.path.isdir(contract_path):
            continue
        
        for file in os.listdir(contract_path):
            if not file.endswith('.json'):
                continue
            
            json_path = os.path.join(contract_path, file)
            row = process_json_file(json_path, contract_id, file)
            if row:
                rows.append(row)
                all_columns.update(row.keys())
    
    if not rows:
        logger.warning("No valid JSON files found.")
        return
    
    all_columns = list(all_columns)
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=all_columns)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def prepend_new_contracts_to_csv(inp_json_dir: str, csv_path: str):
    """Prepend new contracts to existing CSV file"""
    if not os.path.exists(csv_path):
        build_csv_from_scratch(inp_json_dir, csv_path)
        logger.info("CSV created from scratch.")
        return
    
    existing_contract_ids, existing_rows, existing_columns = get_existing_contract_ids_and_rows(csv_path)
    new_rows = []
    all_columns = set(existing_columns)
    
    for contract_id in os.listdir(inp_json_dir):
        contract_path = os.path.join(inp_json_dir, contract_id)
        if not os.path.isdir(contract_path):
            continue
        
        if contract_id in existing_contract_ids:
            continue
        
        for file in os.listdir(contract_path):
            if not file.endswith('.json'):
                continue
            
            json_path = os.path.join(contract_path, file)
            row = process_json_file(json_path, contract_id, file)
            if row:
                new_rows.append(row)
                all_columns.update(row.keys())
    
    if not new_rows:
        logger.info("No new contracts to prepend.")
        return
    
    all_columns = list(all_columns)
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=all_columns)
        writer.writeheader()
        for row in new_rows:
            writer.writerow(row)
        for row in existing_rows:
            writer.writerow(row)


def get_next_contract_folder(output_dir: str) -> str:
    """Get the next available contract folder name"""
    pattern = re.compile(r'^User_Contract_(\d+)$')
    existing = []
    
    for name in os.listdir(output_dir):
        if pattern.match(name) and os.path.isdir(os.path.join(output_dir, name)):
            match = pattern.match(name)
            if match:
                existing.append(int(match.group(1)))
    
    next_num = max(existing, default=0) + 1
    folder_name = f'User_Contract_{next_num}'
    folder_path = os.path.join(output_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def handle_upload(filepath: str, output_dir: str) -> tuple[str, str | None]:
    """Handle file upload and extraction"""
    filename = os.path.basename(filepath)
    _, ext = os.path.splitext(filename.lower())
    
    if ext not in ACCEPTED_EXTENSIONS:
        msg = f"Accepted formats are: {', '.join(sorted(ACCEPTED_EXTENSIONS))}"
        return msg, None
    
    os.makedirs(output_dir, exist_ok=True)
    contract_folder = get_next_contract_folder(output_dir)
    
    if ext != '.zip':
        dest_path = os.path.join(contract_folder, filename)
        if not os.path.exists(dest_path):
            shutil.copy2(filepath, dest_path)
            msg = f"{filename} copied to {contract_folder}."
        else:
            msg = f"{filename} already exists in {contract_folder}. Skipped."
        return msg, contract_folder
    
    # Handle ZIP files
    with zipfile.ZipFile(filepath, 'r') as zip_ref:
        extracted = []
        skipped = []
        
        for member in zip_ref.namelist():
            member_ext = os.path.splitext(member.lower())[1]
            if member_ext in ACCEPTED_NONZIP:
                dest_path = os.path.join(contract_folder, os.path.basename(member))
                if not os.path.exists(dest_path):
                    with zip_ref.open(member) as source, open(dest_path, 'wb') as target:
                        shutil.copyfileobj(source, target)
                    extracted.append(os.path.basename(member))
                else:
                    skipped.append(os.path.basename(member))
            else:
                skipped.append(os.path.basename(member))
        
        msg = f"Extracted: {extracted}\nSkipped: {skipped}\nIn folder: {contract_folder}"
        return msg, contract_folder


def get_token() -> str:
    """Retrieve API authentication token"""
    while True:
        try:
            token_response = requests.get(
                'https://swoosh-process-api-stg.51433.app.standardchartered.com/v1/swoosh/authn/sc-idp/token',
                verify=False
            )
            token = token_response.json().get('access_token')
            if token:
                return token
        except Exception as e:
            logger.warning(f"Failed to get token: {e}")
        
        logger.warning("Retrying token in 120 seconds...")
        time.sleep(120)


def llm_callback(prompt, max_retries: int = 3):
    """Call LLM API with retry mechanism"""
    token = get_token()
    headers = {'Authorization': f"Bearer {token}", 'Content-Type': 'application/json'}
    payload = {
        "model": "claude-sonnet-4-5",
        "messages": [
            {
                "role": "system",
                "content": [
                    {
                        "type": "text",
                        "text": "\nYou are an expert in document processing and meta data extraction"
                    }
                ]
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.0,
        "seed": 42
    }

    for attempt in range(max_retries):
        try:
            response = requests.post(
                'https://gateway.scaifactory.dev.azure.scbdev.net/v1/chat/completions',
                json=payload,
                headers=headers,
                verify=False
            )
            resp_json = response.json()
            error = resp_json.get('error', {})
            error_message = error.get('message', '')
            
            # Parse error message
            parsed_error = {}
            if isinstance(error_message, str):
                try:
                    parsed_error = json.loads(error_message)
                except (json.JSONDecodeError, TypeError):
                    parsed_error = {}
            
            # Handle rate limit error
            if (parsed_error.get('error_code') == 'REQUEST_LIMIT_EXCEEDED' or
                'REQUEST_LIMIT_EXCEEDED' in error_message):
                logger.warning("Rate limit exceeded, waiting 60 seconds before retrying...")
                time.sleep(60)
                continue
            
            # Handle input too long error
            if (parsed_error.get('message', '') == "Input is too long for requested model." or
                "Input is too long for requested model." in error_message):
                logger.error("Input too long for model.")
                return None
            
            # Successfully get response
            llm_output = resp_json.get('choices', [{}])[0].get('message', {}).get('content')
            return llm_output
            
        except Exception as e:
            logger.error(f"Error in processing the LLM: {e}")
            logger.error(traceback.format_exc())
    
    logger.error(f"Failed after {max_retries} retries")
    return None


class FileConverter:
    """File converter for transforming various formats to PDF"""
    
    @staticmethod
    def convert_xlsx_to_pdf(xlsx_path: str, pdf_path: str) -> bool:
        """Convert Excel file to PDF format"""
        try:
            df = pd.read_excel(xlsx_path)
            csv_temp_path = pdf_path.replace('.pdf', '.csv')
            df.to_csv(csv_temp_path, index=False)
            return FileConverter.convert_csv_to_pdf(csv_temp_path, pdf_path)
        except Exception as e:
            logger.error(f"Failed to convert Excel to PDF: {e}")
            return False
    
    @staticmethod
    def convert_csv_to_pdf(csv_path: str, pdf_path: str) -> bool:
        """Convert CSV file to PDF format"""
        try:
            df = pd.read_csv(csv_path)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            
            pdf.cell(200, 10, txt=f"CSV File: {Path(csv_path).name}", ln=True, align='C')
            pdf.ln(10)
            
            for index, row in df.iterrows():
                line = " | ".join([str(x) for x in row.values])
                pdf.multi_cell(0, 10, txt=line)
                if index > 50:  # Limit number of rows
                    pdf.cell(0, 10, txt="... (truncated)", ln=True)
                    break
            
            pdf.output(pdf_path)
            return True
        except Exception as e:
            logger.error(f"Failed to convert CSV to PDF: {e}")
            return False
    
    @staticmethod
    def file_to_base64(file_path: str) -> str | None:
        """Convert file content to Base64 encoding"""
        try:
            with open(file_path, "rb") as file:
                encoded_content = base64.b64encode(file.read()).decode('utf-8')
            return encoded_content
        except Exception as e:
            logger.error(f"Error converting file {file_path} to Base64: {e}")
            return None
    
    @staticmethod
    def extract_text_from_pdf(pdf_path: str) -> str | None:
        """Extract text from PDF file"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            logger.error(f"Failed to extract text from PDF {pdf_path}: {e}")
            return None


class ContractMetadataExtractor:
    """Contract metadata extractor for processing contract documents"""
    
    def __init__(self, base_folder: str, user_contract_folder: str, 
                 prompt_template_path: str, llm_callback: callable,
                 output_folder: str):
        
        self.base_folder = base_folder
        self.llm_callback = llm_callback
        self.output_folder = output_folder
        self.user_contract_folder = user_contract_folder
        
        # Load prompt template
        with open(prompt_template_path, 'r', encoding='utf-8') as f:
            self.prompt_template = f.read()
    
    def _convert_file_to_pdf(self, file_path: str, contract_name: str) -> tuple[str, bool]:
        """Convert various file formats to PDF for processing"""
        file_extension = Path(file_path).suffix.lower()
        temp_pdf_path = os.path.join("doc_to_pdf", f"{contract_name}_{Path(file_path).stem}.pdf")
        
        # If file is already PDF, use it directly
        if file_extension == '.pdf':
            return file_path, True
        
        # Check if already converted
        if os.path.exists(temp_pdf_path):
            logger.info(f"File already converted and present in temp: {temp_pdf_path}")
            return temp_pdf_path, True
        
        # Convert based on file type
        conversion_success = False
        try:
            if file_extension == '.docx':
                convert(file_path, temp_pdf_path)
                conversion_success = True
            elif file_extension in ['.xlsx', '.xls']:
                conversion_success = FileConverter.convert_xlsx_to_pdf(file_path, temp_pdf_path)
            elif file_extension == '.csv':
                conversion_success = FileConverter.convert_csv_to_pdf(file_path, temp_pdf_path)
            # Note: The following converters require additional libraries
            # elif file_extension in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            #     conversion_success = self.convert_image_to_pdf(file_path, temp_pdf_path)
            # elif file_extension == '.msg':
            #     conversion_success = self.convert_msg_to_pdf(file_path, temp_pdf_path)
        except Exception as e:
            logger.error(f"Failed to convert file {file_path} to PDF: {e}")
        
        return temp_pdf_path, conversion_success
    
    def _build_prompt(self, contract_number: str, pdf_base64: str, 
                     base_prompt: str, md_text: str = None) -> list:
        """Build LLM prompt with document content"""
        prompt_replace = base_prompt.replace("{{contract_number}}", contract_number)
        
        pdf_content = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": pdf_base64
            }
        }
        
        text_content = {
            "type": "text",
            "text": md_text
        }
        
        prompt = [
            {
                "type": "text",
                "text": f"{prompt_replace}"
            },
            text_content if md_text is not None else pdf_content
        ]
        
        return prompt
    
    def process_single_contract(self, contract_folder: str) -> dict:
        """Process all files in a single contract folder"""
        contract_name = Path(contract_folder).name
        logger.info(f"Processing contract: {contract_name} ############################################")

        contract_files = glob.glob(os.path.join(contract_folder, "*"))
        if not contract_files:
            logger.warning(f"No files in folder: {contract_folder}")
            return {}

        results = {}
        output_folder = os.path.join(self.output_folder, contract_name)
        os.makedirs(output_folder, exist_ok=True)
        
        for contract_file in contract_files:
            file_extension = Path(contract_file).suffix.lower()
            file_name = Path(contract_file).name
            
            # Skip unsupported file formats
            if file_extension in [".msg", ".doc", ".xls"]:
                logger.warning(f"Skipping unsupported format: {file_extension}")
                results[file_name] = "Unsupported format"
                continue
            
            # Prepare output path
            output_filename = Path(contract_file).stem + '.json'
            output_path = os.path.join(output_folder, output_filename)
            
            # Check if already processed
            if os.path.exists(output_path):
                logger.info(f"File already processed: {output_path}")
                results[file_name] = "Processed Already!"
                continue
            
            # Handle file conversion
            logger.info(f"Processing file: {file_name}")
            pdf_path, conversion_success = self._convert_file_to_pdf(contract_file, contract_name)
            
            encoded_content = None
            md_text = None
            
            if conversion_success and len(PdfReader(pdf_path).pages) < 100:
                encoded_content = FileConverter.file_to_base64(pdf_path)
            else:
                # For large files or failed conversions, skip processing
                logger.warning(f"Skipping file (conversion failed or too large): {file_name}")
                results[file_name] = "File too large or conversion failed"
                continue
            
            # Build prompt and call LLM
            prompt = self._build_prompt(contract_name, encoded_content, 
                                       self.prompt_template, md_text)
            
            try:
                llm_result = self.llm_callback(prompt)
                if not llm_result:
                    raise ValueError("LLM returned no result")
                
                output_result = coerce_to_json(llm_result)
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(output_result, f, ensure_ascii=False, indent=2)
                
                logger.info(f"Saved results: {output_path}")
                results[file_name] = "Successfully processed"
                
            except Exception as e:
                logger.error(f"Process failed for {contract_file}: {e}")
                results[file_name] = f"Failed: {str(e)}"
        
        return results
    
    def process_all_contracts(self) -> dict:
        """Process all contract folders in the base directory"""
        all_results = {}
        subfolders = [self.user_contract_folder]
        
        if not subfolders:
            logger.warning(f"No subfolders in: {self.base_folder}")
            return all_results
        
        for folder in subfolders:
            if folder is None:
                continue
            try:
                results = self.process_single_contract(folder)
                if results:
                    all_results[Path(folder).name] = results
            except Exception as e:
                logger.error(f"Folder process failed: {folder}: {e}")
                all_results[Path(folder).name] = f"Failed: {str(e)}"
        
        return all_results
