def process_single_pdf(pdf_path, template_json, token, output_path=None):
    print(f"Processing: {pdf_path}")
    pdf_base64 = pdf_to_base64(pdf_path)
    llm_response = call_llm(pdf_base64, template_json, token)
    
    # 记录原始响应以备调试
    raw_response_file = output_path.with_suffix('.raw.txt') if output_path else None
    if raw_response_file:
        with open(raw_response_file, 'w', encoding='utf-8') as f:
            f.write(llm_response)
    
    try:
        filled_json = extract_json_from_response(llm_response)
    except Exception as e:
        print(f"JSON parsing failed for {pdf_path}: {e}")
        print(f"Raw response saved to {raw_response_file}")
        raise
    
    if output_path is None:
        output_path = pdf_path.with_suffix('.json')
    else:
        output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(filled_json, f, indent=2, ensure_ascii=False)
    print(f"Saved to: {output_path}")
    return filled_json
