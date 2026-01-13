import os
import json
import glob
from pathlib import Path
import PyPDF2
import base64
import requests
from docx2pdf import convert
import re
import csv
from PyPDF2 import PdfReader
import zipfile
from fpdf import FPDF
import pandas as pd
import img2pdf
# import win32com.client
import pythoncom
import urllib3
import traceback
import time
import sys
sys.path.append("..")
import warnings
from urllib3.exceptions import InsecureRequestWarning

import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
# from src.docling.docling_custom import md_file_create

# Suppress only InsecureRequestWarning
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

print("####cwd",os.getcwd())
from pipelines.extraction import extraction_pipeline_clauses as LLM_extraction_pipeline_clauses
from logger.logger import get_logger
logger = get_logger()

JSON_BLOCK_RE = re.compile(r"\{[\s\S]*\}", re.MULTILINE)

def extract_json_block(text):
    """Best-effort: extract the first JSON object from a string."""
    match = JSON_BLOCK_RE.search(text.strip())
    return match.group(0) if match else None

def coerce_to_json(text):
    """Try to parse JSON directly; if it fails, try extracting a JSON block first."""
    try:
        return json.loads(text)
    except Exception:
        block = extract_json_block(text)
        if not block:
            raise ValueError("No JSON block found in LLM output")
        return json.loads(block)

class ContractMetadataExtractor:
    def __init__(self, base_folder, user_contract_folder, prompt_template_path, llm_callback, output_folder):

        self.base_folder = base_folder
        self.llm_callback = llm_callback
        self.output_folder = output_folder
        self.user_contract_folder = user_contract_folder

        with open(prompt_template_path, 'r', encoding='utf-8') as f:
            self.prompt_template = f.read()
    
    def extract_text_from_pdf(self, pdf_path):
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            logger.error(f"提取PDF文本失败: {pdf_path}, 错误: {e}")
            return None
        
    # Function to convert file content to Base64
    def file_to_base64(self,file_path):
        try:
            with open(file_path, "rb") as file:
                # Read the file content as bytes and encode it to Base64
                encoded_content = base64.b64encode(file.read()).decode('utf-8')
            return encoded_content
        except Exception as e:
            print(f"Error converting file {file_path} to Base64: {e}")
            return None
        
    def convert_xlsx_to_pdf(self, xlsx_path, pdf_path):
        """convert Excel to PDF"""
        try:
            
            df = pd.read_excel(xlsx_path, sheet_name=None)
            
            
            html_content = "<html><head><style>table {border-collapse: collapse; width: 100%;} th, td {border: 1px solid black; padding: 8px; text-align: left;}</style></head><body>"
            
            for sheet_name, data in df.items():
                html_content += f"<h2>Sheet: {sheet_name}</h2>"
                html_content += data.to_html(index=False, classes='table table-striped')
                html_content += "<br>"
            
            html_content += "</body></html>"
                
            return self._convert_xlsx_to_pdf_fallback(xlsx_path, pdf_path)
                
        except Exception as e:
            logger.error(f"Failed to convert Excel to PDF: {e}")
            return False
    
    def _convert_xlsx_to_pdf_fallback(self, xlsx_path, pdf_path):
        
        try:
            df = pd.read_excel(xlsx_path)
            csv_temp_path = pdf_path.replace('.pdf', '.csv')
            df.to_csv(csv_temp_path, index=False)
            return self.convert_csv_to_pdf(csv_temp_path, pdf_path)
        except Exception as e:
            logger.error(f"Fallback conversion failed: {e}")
            return False
    
    def convert_csv_to_pdf(self, csv_path, pdf_path):
        """convert CSV to PDF"""
        try:
            df = pd.read_csv(csv_path)
            html_content = f"<html><head><style>table {{border-collapse: collapse; width: 100%;}} th, td {{border: 1px solid black; padding: 8px; text-align: left;}}</style></head><body>"
            html_content += df.to_html(index=False, classes='table table-striped')
            html_content += "</body></html>"
            
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)
        
            pdf.cell(200, 10, txt=f"CSV File: {Path(csv_path).name}", ln=True, align='C')
            pdf.ln(10)
            
            for index, row in df.iterrows():
                line = " | ".join([str(x) for x in row.values])
                pdf.multi_cell(0, 10, txt=line)
                if index > 50:
                    pdf.cell(0, 10, txt="... (truncated)", ln=True)
                    break
            
            pdf.output(pdf_path)
            return True
                
        except Exception as e:
            logger.error(f"Failed to convert CSV to PDF: {e}")
            return False
    
    def convert_image_to_pdf(self, image_path, pdf_path):
        """convert image to PDF"""
        try:
            with open(pdf_path, "wb") as f:
                f.write(img2pdf.convert(image_path))
            return True
        except Exception as e:
            logger.error(f"Failed to convert image to PDF: {e}")
            return False
    
    def convert_msg_to_pdf(self, msg_path, pdf_path):
        """Convert Outlook emails to 为PDF"""
        try:
            pythoncom.CoInitialize()
            
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            msg = outlook.OpenSharedItem(msg_path)
            content = f"""
            Subject: {msg.Subject}
            From: {msg.SenderName} ({msg.SenderEmailAddress})
            To: {msg.To}
            Date: {msg.SentOn}
            Body:
            {msg.Body}
            """
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, txt=content)
            pdf.output(pdf_path)
            return True
            
        except Exception as e:
            logger.error(f"Failed to convert MSG to PDF: {e}")
            return False
        finally:
            pythoncom.CoUninitialize()
    
    def process_single_contract(self, contract_folder):
        contract_name = Path(contract_folder).name
        logger.info(f"Porcess contract: {contract_name} ############################################")

        contract_files = glob.glob(os.path.join(contract_folder, "*"))
        
        if not contract_files:
            logger.warning(f"No files in the folder: {contract_folder} ")
            return
        
        ###########################################################################
        # if len(contract_files)>10:
        #     logger.warning(f"More than 10 files in the contract: {contract_name}")
        #     return
        # elif not contract_name in self.collected_contract_list:
        #     logger.warning(f"The contract is Not there in the list: {contract_name}")
        #     return


        results = {}
        
        for contract_file in contract_files:
            file_extension = Path(contract_file).suffix.lower()

            if file_extension in [".msg",".doc",".xls"]:
                logger.warning(f"Skipping unsupported format: {file_extension} ")
                continue
            
            output_filename = Path(contract_file).stem + '.json'
            output_folder = os.path.join(self.output_folder, contract_name)
            
            # if os.path.isdir(output_folder):
            #     continue
            
            os.makedirs(output_folder,exist_ok=True)
            output_path = os.path.join(output_folder, output_filename)

            if os.path.exists(output_path):
                logger.info(f"File processed already: {output_path}")
                results[Path(contract_file).name] = "Processed Already!"
                continue

            file_name = Path(contract_file).name
            logger.info(f"Process file: {file_name}")
            
            md_text = None; encoded_content = None; conversion_success = False; prompt = None

            temp_pdf_path = os.path.join("doc_to_pdf",contract_name+"_"+Path(contract_file).stem+'.pdf')

            if file_extension == '.pdf':
                temp_pdf_path = contract_file
                conversion_success = True
                
            elif os.path.exists(temp_pdf_path):
                logger.info(f"File convereted already and present in temp: {temp_pdf_path}")
                conversion_success = True
                    
            elif file_extension == '.docx':
                try:
                    convert(contract_file, temp_pdf_path)
                    conversion_success = True
                except Exception as e:
                    logger.error(f"Failed to convert DOCX to PDF: {e}")
                
            elif file_extension in ['.xlsx', '.xls']:
                conversion_success = self.convert_xlsx_to_pdf(contract_file, temp_pdf_path)
                
            elif file_extension == '.csv':
                conversion_success = self.convert_csv_to_pdf(contract_file, temp_pdf_path)
                
            elif file_extension in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
                conversion_success = self.convert_image_to_pdf(contract_file, temp_pdf_path)
                
            elif file_extension == '.msg':
                conversion_success = self.convert_msg_to_pdf(contract_file, temp_pdf_path)
            
            if conversion_success == True and len(PdfReader(temp_pdf_path).pages) < 100:
                encoded_content = self.file_to_base64(temp_pdf_path)
            else:
                continue
                md_pth = os.path.join("doc_to_pdf",contract_name+"_"+Path(contract_file).stem+'.md')
                if os.path.exists(md_pth):
                    logger.info(f"DOCLING file exists: {md_pth}")
                    with open(md_pth, 'r', encoding='utf-8') as file:
                        md_text = file.read()
                else:
                    logger.info("DOCLING File processing started...............")
                    md_text = md_file_create(contract_file)
                    logger.info("DOCLING File processed succesfully...............")
                    with open(md_pth, "w", encoding="utf-8") as f:
                        f.write(md_text)
            
            prompt = self.build_prompt(contract_name, encoded_content, base_prompt=self.prompt_template, md_text = md_text)
            
            try:
                llm_result = self.llm_callback(prompt)

                output_result = coerce_to_json(llm_result)
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(output_result, f, ensure_ascii=False, indent=2)
                
                logger.info(f"Saved results: {output_path}")
                results[file_name] = "Successfully!"
                
            except Exception as e:
                logger.error(f"Process failed for {contract_file}: {e}")
                results[file_name] = f"Failed: {str(e)}"
        
        return results
    
    def process_all_contracts(self):
        all_results = {}

        #subfolders = [f.path for f in os.scandir(self.base_folder) if f.is_dir()]

        subfolders = [self.user_contract_folder]
        
        if not subfolders:
            logger.warning(f"No sub folder: {self.base_folder}")
            return all_results
        
        for folder in subfolders:
            try:
                results = self.process_single_contract(folder)
                if results:
                    all_results[Path(folder).name] = results
            except Exception as e:
                #traceback.print_exc()
                logger.error(f"Folder process failed: {folder}: {e}")
                all_results[Path(folder).name] = f"Failed: {str(e)}"
        
        return all_results
    
    def build_prompt(self,contract_number,pdf_base64,base_prompt, md_text=None):
        prompt_replace = base_prompt.replace("{{contract_number}}",contract_number)
        
        pdf = {         "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_base64
            }
                    }
                    
        text = {         "type": "text",
                     "text" : md_text
            }

            
        prompt = [
            {
                "type": "text",
                "text": f"{prompt_replace}"
            },
            text if md_text is not None else pdf
        ]
        return prompt
    
def get_token():
    while True:
        try:
            token_response = requests.get(
                'https://swoosh-process-api-stg.51433.app.standardchartered.com/v1/swoosh/authn/sc-idp/token',
                verify=False
            )
            token = token_response.json().get('access_token')
        except Exception as e:
            logger.warning(f"Exception occurred: {e}")
            token = None

        if token:
            return token
        else:
            logger.warning(f"Retrying token ...........")
            time.sleep(120)  # Wait for 2 minutes


def llm_callback(prompt, max_retries=3):
    TOKEN = get_token()
    headers = {'Authorization': f"Bearer {TOKEN}", 'Content-Type': 'application/json'}
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

    llm_response = None
    for attempt in range(max_retries):
        try:
            llm_response = requests.post(
                'https://gateway.scaifactory.dev.azure.scbdev.net/v1/chat/completions',
                json=payload,
                headers=headers,
                verify=False
            )
            resp_json = llm_response.json()
            error = resp_json.get('error', {})
            error_message = error.get('message', '')
            parsed_error = {}
            if isinstance(error_message, str):
                try:
                    parsed_error = json.loads(error_message)
                except Exception:
                    parsed_error = {}
            # Check for rate limit error
            if (parsed_error.get('error_code') == 'REQUEST_LIMIT_EXCEEDED' or
                'REQUEST_LIMIT_EXCEEDED' in error_message):
                logger.warning("Rate limit exceeded, waiting 60 seconds before retrying...")
                time.sleep(60)
                continue
            # Check for input too long error
            if (parsed_error.get('message', '') == "Input is too long for requested model." or
                "Input is too long for requested model." in error_message):
                logger.error("Input too long for model.")
                break
            # Success
            llm_output = resp_json.get('choices', [{}])[0].get('message', {}).get('content')
            return llm_output
        except Exception as e:
            logger.error(f"Error in processing the LLM: {TOKEN}")
            logger.error(traceback.format_exc())
    logger.error(f"Failed after {max_retries} retries: {llm_response.json() if llm_response else 'No response'}")
    return None

import os
import shutil
import zipfile
import re

ACCEPTED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.pdf', '.docx', '.zip'}
ACCEPTED_NONZIP = ACCEPTED_EXTENSIONS - {'.zip'}

def get_next_contract_folder(output_dir):
    pattern = re.compile(r'^User_Contract_(\d+)$')
    existing = [
        int(pattern.match(name).group(1))
        for name in os.listdir(output_dir)
        if pattern.match(name) and os.path.isdir(os.path.join(output_dir, name))
    ]
    next_num = max(existing, default=0) + 1
    folder_name = f'User_Contract_{next_num}'
    folder_path = os.path.join(output_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def handle_upload(filepath, output_dir):
    filename = os.path.basename(filepath)
    _, ext = os.path.splitext(filename.lower())
    if ext not in ACCEPTED_EXTENSIONS:
        return f"Accepted formats are: {', '.join(sorted(ACCEPTED_EXTENSIONS))}", None

    os.makedirs(output_dir, exist_ok=True)
    contract_folder = get_next_contract_folder(output_dir)

    if ext != '.zip':
        dest_path = os.path.join(contract_folder, filename)
        if not os.path.exists(dest_path):
            shutil.copy2(filepath, dest_path)
            return f"{filename} copied to {contract_folder}.", contract_folder
        else:
            return f"{filename} already exists in {contract_folder}. Skipped.", contract_folder
    else:
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
        
import os
import json
import csv
import re

def to_pascal_case_with_underscore(s):
    s = re.sub(r'[^a-zA-Z0-9 ]', '', s)
    words = s.strip().split()
    return '_'.join(word.capitalize() for word in words)

def clean_column_name(name):
    return to_pascal_case_with_underscore(name)

def is_protected_or_corrupted(data):
    return (
        list(data.keys()) == ["DOCUMENT CATEGORY"] and
        data["DOCUMENT CATEGORY"] in ["Document Protected", "Corrupted"]
    )

def process_json_file(json_path, contract_id, file_name):
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

def get_existing_contract_ids_and_rows(csv_path):
    contract_ids = set()
    existing_rows = []
    columns = []
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        columns = reader.fieldnames
        for row in reader:
            contract_ids.add(row.get('Contract_Reference_Number', ''))
            existing_rows.append(row)
    return contract_ids, existing_rows, columns

def build_csv_from_scratch(inp_json_dir, csv_path):
    rows = []
    all_columns = set()
    for contract_id in os.listdir(inp_json_dir):
        contract_path = os.path.join(inp_json_dir, contract_id)
        if not os.path.isdir(contract_path):
            continue
        for file in os.listdir(contract_path):
            if file.endswith('.json'):
                json_path = os.path.join(contract_path, file)
                row = process_json_file(json_path, contract_id, file)
                if row:
                    rows.append(row)
                    all_columns.update(row.keys())
    if not rows:
        print("No valid JSONs found.")
        return
    all_columns = list(all_columns)
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=all_columns)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

def prepend_new_contracts_to_csv(inp_json_dir, csv_path):
    if not os.path.exists(csv_path):
        build_csv_from_scratch(inp_json_dir, csv_path)
        print("CSV created from scratch.")
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
            if file.endswith('.json'):
                json_path = os.path.join(contract_path, file)
                row = process_json_file(json_path, contract_id, file)
                if row:
                    new_rows.append(row)
                    all_columns.update(row.keys())
    if not new_rows:
        print("No new contracts to prepend.")
        return
    all_columns = list(all_columns)
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=all_columns)
        writer.writeheader()
        for row in new_rows:
            writer.writerow(row)
        for row in existing_rows:
            writer.writerow(row)
