import os
import json
import glob
import re
import time
import logging
import traceback
import warnings
from pathlib import Path
from urllib3.exceptions import InsecureRequestWarning

import PyPDF2
import pandas as pd
import requests
from docx2pdf import convert
from fpdf import FPDF
from PyPDF2 import PdfReader

# 注意：以下导入已被注释掉或可能需要额外安装
# import base64
# import img2pdf
# import pythoncom
# import win32com.client
# import py7zr
# from src.docling.docling_custom import md_file_create

# 抑制不安全请求警告
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

# 日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

JSON_BLOCK_RE = re.compile(r"\{[\s\S]*\}", re.MULTILINE)


def extract_json_block(text: str) -> str | None:
    """从字符串中提取第一个JSON对象（尽力而为）"""
    match = JSON_BLOCK_RE.search(text.strip())
    return match.group(0) if match else None


def coerce_to_json(text: str) -> dict:
    """尝试解析JSON，失败则尝试提取JSON块"""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        block = extract_json_block(text)
        if not block:
            raise ValueError("No JSON block found in LLM output")
        return json.loads(block)


def extract_page_numbers(coords: dict) -> str:
    """从坐标字典中提取页码"""
    page_numbers = []
    for page in coords.keys():
        match = re.search(r'\d+', str(page))
        if match:
            page_numbers.append(int(match.group()))
    page_numbers = sorted(set(page_numbers))
    return ",".join(str(num) for num in page_numbers) if page_numbers else "0"


def get_existing_contracts(output_csv: str) -> set:
    """从CSV文件中获取已存在的合同引用号"""
    contracts = set()
    if os.path.exists(output_csv):
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                contracts.add(row["Contract_Reference_Number"])
    return contracts


def json_to_csv_prepend_new(input_folder: str, output_csv: str):
    """将JSON文件转换为CSV，新数据放在顶部"""
    header = [
        "Contract_Reference_Number", "File_Name", "clause_type", "clause_key",
        "is_present", "clause_coverage_score", "confidence", "page_number"
    ]

    # 获取现有合同
    existing_contracts = get_existing_contracts(output_csv)

    # 列出输入文件夹中的所有子目录
    all_contracts = [
        d for d in os.listdir(input_folder)
        if os.path.isdir(os.path.join(input_folder, d))
    ]

    # 找出新合同
    new_contracts = [c for c in all_contracts if c not in existing_contracts]

    # 收集新行数据
    new_rows = []
    for contract_reference_number in new_contracts:
        contract_dir = os.path.join(input_folder, contract_reference_number)
        for file in os.listdir(contract_dir):
            if not file.endswith(".json"):
                continue
                
            file_path = os.path.join(contract_dir, file)
            with open(file_path, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError as e:
                    logger.error(f"Failed to parse JSON {file_path}: {e}")
                    continue

            # 处理标准条款
            for clause in data.get("standard_general_clause_details", []):
                if not clause.get("is_present"):
                    continue
                coords = clause.get("coordinates") or {}
                page_numbers = extract_page_numbers(coords)
                new_rows.append([
                    contract_reference_number,
                    file,
                    "standard and general",
                    clause.get("clause_key"),
                    clause.get("is_present"),
                    clause.get("clause_coverage_score"),
                    clause.get("confidence"),
                    page_numbers
                ])

            # 处理服务特定条款
            for service_name, clauses in data.get("service_specific_clauses", {}).items():
                for clause in clauses:
                    if not clause.get("is_present"):
                        continue
                    coords = clause.get("coordinates") or {}
                    page_numbers = extract_page_numbers(coords)
                    new_rows.append([
                        contract_reference_number,
                        file,
                        service_name,
                        clause.get("clause_key"),
                        clause.get("is_present"),
                        clause.get("clause_coverage_score"),
                        clause.get("confidence"),
                        page_numbers
                    ])

    # 读取旧行数据
    old_rows = []
    if os.path.exists(output_csv):
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            old_rows = list(reader)
        # 移除标题行
        if old_rows and old_rows[0][0] == header[0]:
            old_rows = old_rows[1:]

    # 写入新数据在顶部，然后是旧数据
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(new_rows)
        writer.writerows(old_rows)


def get_token() -> str:
    """获取API令牌"""
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


def llm_callback(prompt, max_retries=3):
    """调用LLM API"""
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
            
            # 解析错误信息
            parsed_error = {}
            if isinstance(error_message, str):
                try:
                    parsed_error = json.loads(error_message)
                except (json.JSONDecodeError, TypeError):
                    parsed_error = {}
            
            # 处理速率限制错误
            if (parsed_error.get('error_code') == 'REQUEST_LIMIT_EXCEEDED' or
                'REQUEST_LIMIT_EXCEEDED' in error_message):
                logger.warning("Rate limit exceeded, waiting 60 seconds before retrying...")
                time.sleep(60)
                continue
            
            # 处理输入过长错误
            if (parsed_error.get('message', '') == "Input is too long for requested model." or
                "Input is too long for requested model." in error_message):
                logger.error("Input too long for model.")
                return None
            
            # 成功获取响应
            llm_output = resp_json.get('choices', [{}])[0].get('message', {}).get('content')
            return llm_output
            
        except Exception as e:
            logger.error(f"Error in processing the LLM: {e}")
            logger.error(traceback.format_exc())
    
    logger.error(f"Failed after {max_retries} retries")
    return None


class FileConverter:
    """文件转换器，处理各种格式到PDF的转换"""
    
    @staticmethod
    def convert_xlsx_to_pdf(xlsx_path: str, pdf_path: str) -> bool:
        """转换Excel文件到PDF"""
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
        """转换CSV文件到PDF"""
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
                if index > 50:  # 限制行数
                    pdf.cell(0, 10, txt="... (truncated)", ln=True)
                    break
            
            pdf.output(pdf_path)
            return True
        except Exception as e:
            logger.error(f"Failed to convert CSV to PDF: {e}")
            return False
    
    @staticmethod
    def file_to_base64(file_path: str) -> str | None:
        """将文件转换为Base64编码"""
        try:
            with open(file_path, "rb") as file:
                encoded_content = base64.b64encode(file.read()).decode('utf-8')
            return encoded_content
        except Exception as e:
            logger.error(f"Error converting file {file_path} to Base64: {e}")
            return None


class ContractMetadataExtractor:
    """合同元数据提取器"""
    
    def __init__(self, base_folder: str, user_prompt_path: str, 
                 clause_details_path: str, llm_callback: callable,
                 output_folder: str, meta_output_folder: str):
        
        self.base_folder = base_folder
        self.llm_callback = llm_callback
        self.output_folder = output_folder
        self.meta_output_folder = meta_output_folder
        self.services = []
        
        # 加载提示词和条款详情
        with open(user_prompt_path, 'r', encoding='utf-8') as f:
            self.user_prompt = f.read()
            
        with open(clause_details_path, 'r', encoding='utf-8') as f:
            self.clause_details = json.load(f)
    
    def _get_normalised_services(self, data: dict) -> list:
        """从元数据中获取标准化服务列表"""
        details = data.get("DOCUMENT DETAILS", {})
        services = details.get("SERVICES/SUPPLIES DETAILS", {})
        mentioned = services.get("Services Mentioned", {})
        return mentioned.get("normalised value", [])
    
    def _create_clause_template(self, services: list) -> tuple[dict, dict]:
        """创建条款模板"""
        standard_general_clause_details = []
        for clause in self.clause_details.get("standard_general_clause_details", [])[:2]:
            standard_general_clause_details.append({
                "clause_key": clause.get("clause_key"),
                "is_present": None,
                "clause_coverage_score": None,
                "confidence": None,
                "coordinates": None
            })
        
        service_specific_clauses = {}
        for service in services:
            clauses = self.clause_details.get("service_specific_clauses", {}).get(service, [])[:2]
            service_specific_clauses[service] = []
            for clause in clauses:
                service_specific_clauses[service].append({
                    "clause_key": clause.get("clause_key"),
                    "is_present": None,
                    "clause_coverage_score": None,
                    "confidence": None,
                    "coordinates": None
                })
        
        # 创建模板
        clause_template = {
            "standard_general_clause_details": standard_general_clause_details,
            "service_specific_clauses": service_specific_clauses
        }
        
        # 创建过滤后的完整JSON
        filtered_json = {
            "standard_general_clause_details": self.clause_details.get("standard_general_clause_details", []),
            "service_specific_clauses": {
                service: self.clause_details.get("service_specific_clauses", {}).get(service, [])
                for service in services
            }
        }
        
        return clause_template, filtered_json
    
    def _convert_clause_details(self, meta_output_path: str) -> tuple[str, str]:
        """转换条款详情为JSON字符串"""
        try:
            with open(meta_output_path, 'r', encoding='utf-8') as f:
                meta_data = json.load(f)

            services = self._get_normalised_services(meta_data)
            self.services = services
            clause_details_converted, clause_details_filtered = self._create_clause_template(services)
            
            return json.dumps(clause_details_converted), json.dumps(clause_details_filtered)
            
        except Exception as e:
            logger.error(f"Could not read the meta json: {e}")
            return None, None
    
    def _build_prompt(self, contract_number: str, pdf_base64: str, 
                     user_prompt_converted: str, clause_details_filtered: str,
                     md_text: str = None) -> list:
        """构建LLM提示词"""
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
                "text": f"{user_prompt_converted}"
            },
            {
                "type": "text",
                "text": f"{clause_details_filtered}"
            },
            text_content if md_text is not None else pdf_content
        ]
        
        return prompt
    
    def _convert_file_to_pdf(self, file_path: str, contract_name: str) -> tuple[str, bool]:
        """将文件转换为PDF，返回PDF路径和是否成功"""
        file_extension = Path(file_path).suffix.lower()
        temp_pdf_path = os.path.join("doc_to_pdf", f"{contract_name}_{Path(file_path).stem}.pdf")
        
        # 如果已经是PDF，直接使用
        if file_extension == '.pdf':
            return file_path, True
        
        # 检查是否已转换
        if os.path.exists(temp_pdf_path):
            logger.info(f"File converted already and present in temp: {temp_pdf_path}")
            return temp_pdf_path, True
        
        # 根据文件类型进行转换
        conversion_success = False
        try:
            if file_extension == '.docx':
                convert(file_path, temp_pdf_path)
                conversion_success = True
            elif file_extension in ['.xlsx', '.xls']:
                conversion_success = FileConverter.convert_xlsx_to_pdf(file_path, temp_pdf_path)
            elif file_extension == '.csv':
                conversion_success = FileConverter.convert_csv_to_pdf(file_path, temp_pdf_path)
            # 注意：以下转换器需要额外库支持
            # elif file_extension in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            #     conversion_success = self.convert_image_to_pdf(file_path, temp_pdf_path)
            # elif file_extension == '.msg':
            #     conversion_success = self.convert_msg_to_pdf(file_path, temp_pdf_path)
        except Exception as e:
            logger.error(f"Failed to convert file {file_path} to PDF: {e}")
        
        return temp_pdf_path, conversion_success
    
    def _get_greatest_user_contract_dir(self) -> str | None:
        """获取最大的用户合同目录"""
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
    
    def process_single_contract(self, contract_folder: str) -> dict:
        """处理单个合同文件夹"""
        contract_name = Path(contract_folder).name
        logger.info(f"Processing contract: {contract_name}")

        contract_files = glob.glob(os.path.join(contract_folder, "*"))
        if not contract_files:
            logger.warning(f"No files in folder: {contract_folder}")
            return {}

        results = {}
        output_folder = os.path.join(self.output_folder, contract_name)
        meta_output_folder = os.path.join(self.meta_output_folder, contract_name)
        os.makedirs(output_folder, exist_ok=True)
        
        for contract_file in contract_files:
            file_extension = Path(contract_file).suffix.lower()
            file_name = Path(contract_file).name
            
            # 跳过不支持的文件格式
            if file_extension in [".msg", ".doc", ".xls"]:
                logger.warning(f"Skipping unsupported format: {file_extension}")
                results[file_name] = "Unsupported format"
                continue
            
            # 准备输出路径
            output_filename = Path(contract_file).stem + '.json'
            output_path = os.path.join(output_folder, output_filename)
            meta_output_path = os.path.join(meta_output_folder, output_filename)
            
            # 检查元数据是否存在
            if not os.path.exists(meta_output_path):
                logger.warning(f"No meta data json exists: {meta_output_path}")
                results[file_name] = "Missing meta data"
                continue
            
            # 检查是否已处理
            if os.path.exists(output_path):
                logger.info(f"File already processed: {output_path}")
                results[file_name] = "Processed Already!"
                continue
            
            # 转换条款详情
            clause_details_converted, clause_details_filtered = self._convert_clause_details(meta_output_path)
            if clause_details_converted is None:
                results[file_name] = "Failed to convert clause details"
                continue
            
            user_prompt_converted = self.user_prompt + clause_details_converted
            
            # 处理文件转换
            logger.info(f"Processing file: {file_name}")
            pdf_path, conversion_success = self._convert_file_to_pdf(contract_file, contract_name)
            
            encoded_content = None
            md_text = None
            
            if conversion_success and len(PdfReader(pdf_path).pages) < 100:
                encoded_content = FileConverter.file_to_base64(pdf_path)
            else:
                # 对于大文件或转换失败的情况，可以尝试其他方法
                # 注意：这里原来的docling代码被注释掉了
                logger.warning(f"Skipping large file or conversion failed: {file_name}")
                results[file_name] = "File too large or conversion failed"
                continue
            
            # 构建提示词并调用LLM
            prompt = self._build_prompt(contract_name, encoded_content, 
                                       user_prompt_converted, clause_details_filtered, md_text)
            
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
        """处理所有合同"""
        all_results = {}
        subfolders = [self._get_greatest_user_contract_dir()]
        
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
