import os
import json

from PIL import Image
import base64
from io import BytesIO


import requests

model_selection = {
    "chatgpt": "swo-gpt4o",
    "llm":"swo-llama-3-3-70b-instruct",
    "claud": "claude-sonnet-4-5",
    "qwen":"databricks-qwen3-next-80b-a3b-instruct"
}

def get_token():
    token_response = requests.get('/token', verify=False)
    token = token_response.json().get('access_token')
    return token

def base_64_conv(file_pth):

    with open(file_pth,"rb") as pdf_file:
            pdf_bytes = pdf_file.read()
    
        # Encode the bytes to Base64
    pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")
    
    return pdf_base64


    
token = get_token()

with open(r"test_prompt.txt", 'r', encoding = 'utf-8') as file:
    dat_txt = file.read()
    
PDF_path1 = r"services.pdf"

prompt = [
    {
            "type": "text",
            "text": dat_txt},
    
        {         "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base_64_conv(PDF_path1)
         }
    }     
        
    ]
         
headers = {
           'Authorization': f"Bearer {token}", 
           'Content-Type': 'application/json'
           }

payload = {
        "model":model_selection["claud"],
        "messages":[
             {
            "role": "system",
            "content": [
                {
                    "type": "text",
                    'text': "You are an expert in the contract analysis, please extract the requested clause from the given contract document"
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


llm_response = requests.post('v1/chat/completions', json=payload, headers=headers, verify=False)

llm_output = llm_response.json().get('choices')[0]['message']['content']

print(llm_output)

llm_response.json()
 
