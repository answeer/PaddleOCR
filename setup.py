import os
import json
import glob
from pathlib import Path
import PyPDF2
import base64
import logging
import requests
from docx2pdf import convert
import re
import csv
from PyPDF2 import PdfReader
import zipfile
# import py7zr
from fpdf import FPDF
import pandas as pd
import img2pdf
# import win32com.client
import pythoncom
import urllib3
import traceback
import time

import warnings
from urllib3.exceptions import InsecureRequestWarning

import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
# from src.docling.docling_custom import md_file_create

# Suppress only InsecureRequestWarning
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
    def __init__(self, base_folder, 
                 user_prompt_path, clause_details_path,
                 llm_callback, output_folder, meta_output_folder):

        self.base_folder = base_folder
        self.llm_callback = llm_callback
        self.output_folder = output_folder
        self.meta_output_folder = meta_output_folder

        with open(user_prompt_path, 'r', encoding='utf-8') as f:
            self.user_prompt = f.read()
            
        with open(clause_details_path, 'r', encoding='utf-8') as f:
            self.clause_details = json.load(f)

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
      
    def clause_parser(self):
        
        data = self.clause_details
        selected_services = self.services
        
        standard_general_clause_details = []
        for clause in data.get("standard_general_clause_details", [])[:2]:
            standard_general_clause_details.append({
                "clause_key": clause.get("clause_key"),
                "is_present": None,
                "clause_coverage_score": None,
                "confidence": None,
                "coordinates": None
            })
     
        service_specific_clauses = {}
        for service in selected_services:
            clauses = data.get("service_specific_clauses", {}).get(service, [])[:2]
            service_specific_clauses[service] = []
            for clause in clauses:
                service_specific_clauses[service].append({
                    "clause_key": clause.get("clause_key"),
                    "is_present": None,
                    "clause_coverage_score": None,
                    "confidence": None,
                    "coordinates": None
                })
     
        clause_template = {
            "standard_general_clause_details": standard_general_clause_details,
            "service_specific_clauses": service_specific_clauses
        }
     
        # Filtered JSON (all items for selected services)
        filtered_json = {
            "standard_general_clause_details": data.get("standard_general_clause_details", []),
            "service_specific_clauses": {
                service: data.get("service_specific_clauses", {}).get(service, [])
                for service in selected_services
            }
        }
     
        return clause_template, filtered_json
    
    def get_normalised_services(self, data):
        details = data.get("DOCUMENT DETAILS", {})
        services = details.get("SERVICES/SUPPLIES DETAILS", {})
        mentioned = services.get("Services Mentioned", {})
        normalised = mentioned.get("normalised value", [])
        return normalised
            
    def convert_clause_details(self, meta_output_path):

        try:
            with open(meta_output_path, 'r', encoding = 'utf-8') as f:
                meta_data = json.load(f)

            services =  self.get_normalised_services(meta_data)
            self.services = services
            clause_details_converted, clause_details_filtered = self.clause_parser()
            return json.dumps(clause_details_converted), json.dumps(clause_details_filtered)
            
        except Exception as e:
            logger.error(f"Could not read the meta json: {e}")
            return None, None
     
    def process_single_contract(self, contract_folder):
        contract_name = Path(contract_folder).name
        logger.info(f"Porcess contract: {contract_name}")

        contract_files = glob.glob(os.path.join(contract_folder, "*"))

        if not contract_files:
            logger.warning(f"No files in the folder: {contract_folder} ")
            return

        results = {}
        
        for contract_file in contract_files:
            file_extension = Path(contract_file).suffix.lower()
            
            if file_extension in [".msg",".doc",".xls"]:
                logger.warning(f"Skipping unsupported format: {file_extension} ")
                continue
            
            output_filename = Path(contract_file).stem + '.json'
            output_folder = os.path.join(self.output_folder, contract_name)
            meta_output_folder = os.path.join(self.meta_output_folder, contract_name)

            # if os.path.isdir(output_folder):
            #     continue

            os.makedirs(output_folder,exist_ok=True)
            output_path = os.path.join(output_folder, output_filename)
            meta_output_path = os.path.join(meta_output_folder, output_filename)
            
            if not os.path.exists(meta_output_path):
                logger.warning(f"There is no meta data json exists: {meta_output_path}")
                continue
            else:
                clause_details_converted, clause_details_filtered = self.convert_clause_details(meta_output_path)
                
                if clause_details_converted == None:
                    continue
                
                else:
                    user_prompt_converted = self.user_prompt + clause_details_converted

            if os.path.exists(output_path):
                logger.info(f"File processed already: {output_path}")
                results[Path(contract_file).name] = "Processed Already!"
                continue

            file_name = Path(contract_file).name
            logger.info(f"Process file: {file_name}")
            
            md_text = None; encoded_content = None; conversion_success = False

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
                print("SKIPPING DOCKLING...............")
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
                        
            prompt = self.build_prompt(contract_name, encoded_content, user_prompt_converted, 
                                       clause_details_filtered, md_text = md_text)
            
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

    def get_greatest_user_contract_dir(self):
        pattern = re.compile(r'^User_Contract_(\d+)$')
        max_num = -1
        max_dir = None

        for entry in os.listdir(self.base_folder):
            full_path = os.path.join(self.base_folder, entry)
            if os.path.isdir(full_path):
                match = pattern.match(entry)
                if match:
                    num = int(match.group(1))
                    if num > max_num:
                        max_num = num
                        max_dir = full_path

        return max_dir
    
    def process_all_contracts(self):
        all_results = {}

        #subfolders = [f.path for f in os.scandir(self.base_folder) if f.is_dir()]
        subfolders = [self.get_greatest_user_contract_dir()]
        
        if not subfolders:
            logger.warning(f"No sub folder: {self.base_folder}")
            return all_results
        
        for folder in subfolders:
            try:
                results = self.process_single_contract(folder)
                if results:
                    all_results[Path(folder).name] = results
            except Exception as e:
                logger.error(f"Folder process failed: {folder}: {e}")
                all_results[Path(folder).name] = f"Failed: {str(e)}"


        
        
        return all_results
    
    def build_prompt(self,contract_number,pdf_base64,user_prompt_converted, clause_details_filtered, md_text=None):
        
        pdf = { "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": pdf_base64}
                    }
                    
        text = { "type": "text",
                 "text" : md_text
                    }
        
        prompt = [
                {
                    "type": "text",
                    "text": f"{user_prompt_converted}"
                },
                {
                    "type": "text",
                    "text": f"{clause_details_filtered}"
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

def extract_page_numbers(coords):
    page_numbers = []
    for page in coords.keys():
        match = re.search(r'\d+', str(page))
        if match:
            page_numbers.append(int(match.group()))
    page_numbers = sorted(set(page_numbers))
    return ",".join(str(num) for num in page_numbers) if page_numbers else "0"

def get_existing_contracts(output_csv):
    contracts = set()
    if os.path.exists(output_csv):
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                contracts.add(row["Contract_Reference_Number"])
    return contracts

def json_to_csv_prepend_new(input_folder, output_csv):
    header = [
        "Contract_Reference_Number", "File_Name", "clause_type", "clause_key",
        "is_present", "clause_coverage_score", "confidence", "page_number"
    ]

    # 1. Get existing contracts from CSV
    existing_contracts = get_existing_contracts(output_csv)

    # 2. List all subdirectories in input folder
    all_contracts = [
        d for d in os.listdir(input_folder)
        if os.path.isdir(os.path.join(input_folder, d))
    ]

    # 3. Find new contracts
    new_contracts = [c for c in all_contracts if c not in existing_contracts]

    # 4. Collect new rows
    new_rows = []
    for contract_reference_number in new_contracts:
        contract_dir = os.path.join(input_folder, contract_reference_number)
        for file in os.listdir(contract_dir):
            if file.endswith(".json"):
                file_path = os.path.join(contract_dir, file)
                file_name = file
                with open(file_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except Exception:
                        continue

                # Standard clauses
                for clause in data.get("standard_general_clause_details", []):
                    if not clause.get("is_present"):
                        continue
                    clause_key = clause.get("clause_key")
                    coords = clause.get("coordinates") or {}
                    page_numbers = extract_page_numbers(coords)
                    new_rows.append([
                        contract_reference_number,
                        file_name,
                        "standard and general",
                        clause_key,
                        clause.get("is_present"),
                        clause.get("clause_coverage_score"),
                        clause.get("confidence"),
                        page_numbers
                    ])

                # Service-specific clauses
                for service_name, clauses in data.get("service_specific_clauses", {}).items():
                    for clause in clauses:
                        if not clause.get("is_present"):
                            continue
                        clause_key = clause.get("clause_key")
                        coords = clause.get("coordinates") or {}
                        page_numbers = extract_page_numbers(coords)
                        new_rows.append([
                            contract_reference_number,
                            file_name,
                            service_name,
                            clause_key,
                            clause.get("is_present"),
                            clause.get("clause_coverage_score"),
                            clause.get("confidence"),
                            page_numbers
                        ])

    # 5. Read old rows (if any)
    old_rows = []
    if os.path.exists(output_csv):
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            old_rows = list(reader)
        # Remove header from old_rows if present
        if old_rows and old_rows[0][0] == header[0]:
            old_rows = old_rows[1:]

    # 6. Write new rows at the top, then old rows
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(new_rows)
        writer.writerows(old_rows)
