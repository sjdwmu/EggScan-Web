# -*- coding: utf-8 -*-
# =============================================================================
# --- 导入核心库 ---
# =============================================================================
import os
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file, jsonify

# =============================================================================
# --- 全局变量与辅助函数 ---
# =============================================================================
# (注释) 动态导入重量级库，加快程序启动速度
def import_heavy_libraries():
    global fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    if 'fitz' not in globals():
        print("正在加载分析库...")
        import fitz
        import requests
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        print("✓ 分析库加载完成")

LLM_URL = "https://api.deepseek.com/v1/chat/completions"

def smart_extract_text(pdf_path):
    """(注释) 从PDF中提取纯文本。"""
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        return text
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

# =============================================================================
# --- 模式一：泛读模式分析逻辑 ---
# =============================================================================
def call_llm_for_broad_read(pdf_text, api_key):
    prompt_text = """
    请你扮演一位专业的生物医学研究员，仔细阅读以下英文文献内容。
    你的任务是：
    1. 用中文精炼地总结出最重要的核心观点和发现，每一点作为一段。
    2. 在每一段中文总结下方，附上该总结所依据的最核心的1-2句英文原文。
    3. 最后，将你附上的那句“英文原文”翻译成中文。
    请严格按照以下格式返回，每个观点之间用 '---' 分隔：
    [中文提炼]: 中文总结内容。
    [核心原文]: Original English quote.
    [原文翻译]: 核心原文的中文翻译。
    """
    system_prompt = "你是一个擅长快速抓取论文核心亮点的学术助手。"
    user_content = f"{prompt_text}\n\n---\n以下是需要分析的文本内容:\n\n{pdf_text[:40000]}"
    
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}], "temperature": 0.5, "max_tokens": 4096}
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: 请求失败 - {e}"

def parse_broad_read_output(llm_text):
    if llm_text.startswith("API_ERROR:"):
        return [{'中文提炼': llm_text, '核心原文': '', '原文翻译': ''}]
    results = []
    sections = llm_text.strip().split('---')
    for section in sections:
        if not section.strip(): continue
        summary = re.search(r"\[中文提炼\]:\s*(.*)", section, re.DOTALL)
        quote = re.search(r"\[核心原文\]:\s*(.*)", section, re.DOTALL)
        translation = re.search(r"\[原文翻译\]:\s*(.*)", section, re.DOTALL)
        results.append({
            '中文提炼': summary.group(1).strip() if summary else "N/A",
            '核心原文': quote.group(1).strip() if quote else "N/A",
            '原文翻译': translation.group(1).strip() if translation else "N/A",
        })
    return results

# =============================================================================
# --- 模式二：精读模式分析逻辑 ---
# =============================================================================
def call_llm_for_deep_read(pdf_text, api_key, fields, language):
    fields_str = ", ".join(fields)
    if language == '中文':
        instruction = f"作为一名顶尖科研分析师，请对以下论文进行深入的“精读”分析。\n第一步（内心思考）：针对用户指定的每一个分析维度【{fields_str}】，首先在论文全文中定位所有相关信息。\n第二步（输出结果）：基于第一步定位到的信息，对每个维度进行精准、全面且高度浓缩的总结，并用【中文】呈现。\n请严格按照“【字段名】:【总结内容】”的格式输出，不要输出思考过程。"
    else: # English
        instruction = f"As a top-tier research analyst, conduct a deep analysis of the following paper.\nStep 1 (Internal Thought): For each user-specified field [{fields_str}], first locate all relevant information.\nStep 2 (Final Output): Based on Step 1, provide a precise, comprehensive, and condensed summary for each field in **English**.\nStrictly adhere to the format `**[Field Name]**: **[Summary Content]**`. Do not output your thought process."
    
    prompt = f"{instruction}\n\n---\n论文内容如下:\n{pdf_text[:40000]}"
    system_prompt = "你是一个能够执行多步推理和深度分析的学术专家。"
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": prompt}], "temperature": 0.3, "max_tokens": 4096}
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: 请求失败 - {e}"

def parse_deep_read_output(llm_text, fields):
    if llm_text.startswith("API_ERROR:"):
        return {field: (llm_text if i == 0 else "") for i, field in enumerate(fields)}
    
    result_dict = {field: "未提取到" for field in fields}
    field_pattern = "|".join([re.escape(f) for f in fields])
    matches = re.findall(r"【?(" + field_pattern + r")】?:\s*(.*?)(?=\n【?(" + field_pattern + r")】?:|\Z)", llm_text, re.DOTALL | re.IGNORECASE)
    
    for match in matches:
        field_name, content = match[0].strip(), match[1].strip()
        for f in fields:
            if f.lower() in field_name.lower():
                result_dict[f] = content
                break
    return result_dict

# =============================================================================
# --- 模式三：自定义模式分析逻辑 ---
# =============================================================================
def call_llm_for_custom_mode(pdf_text, api_key, custom_prompt):
    system_prompt = "你是一个强大的、通用的文本分析助手。请严格遵循用户提供的指令来分析给定的文本内容。"
    user_content = f"{custom_prompt}\n\n---\n以下是需要分析的文本内容:\n\n{pdf_text[:40000]}"
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}], "temperature": 0.3, "max_tokens": 4096}
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: 请求失败 - {e}"

def parse_custom_output(llm_text):
    if llm_text.startswith("API_ERROR:"):
        return {'错误': llm_text}
    
    result_dict = {}
    matches = re.findall(r"(?:【(.+?)】|(?<=\n)\*\*(.+?)\*\*):\s*(.*)", llm_text)
    if matches:
        for match in matches:
            key, value = (match[0] or match[1]).strip(), match[2].strip()
            result_dict[key] = value
        return result_dict
    
    lines = llm_text.strip().split('\n')
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            result_dict[key.strip()] = value.strip()
    return result_dict if result_dict else {'分析结果': llm_text}

# =============================================================================
# --- 主处理流程 ---
# =============================================================================
def process_single_pdf(pdf_file, api_key, mode, fields, language, custom_prompt):
    """(注释) 根据模式，分发任务给不同的处理函数。"""
    filename = pdf_file.filename
    print(f"📄 开始处理: {filename} (模式: {mode})")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_file.save(tmp.name)
        pdf_path = tmp.name
        
    try:
        text = smart_extract_text(pdf_path)
        if len(text.strip()) < 200:
            print(f"    ⚠️ 文本太少，跳过")
            return []

        if mode == '泛读模式':
            llm_output = call_llm_for_broad_read(text, api_key)
            structured_data_list = parse_broad_read_output(llm_output)
            for item in structured_data_list: item['文件名'] = filename
            return structured_data_list
        
        elif mode == '精读模式':
            llm_output = call_llm_for_deep_read(text, api_key, fields, language)
            structured_data = parse_deep_read_output(llm_output, fields)
            structured_data['文件名'] = filename
            return [structured_data]
            
        elif mode == '自定义模式':
            llm_output = call_llm_for_custom_mode(text, api_key, custom_prompt)
            structured_data = parse_custom_output(llm_output)
            structured_data['文件名'] = filename
            return [structured_data]
            
    except Exception as e:
        print(f"    ❌ 处理时出错: {e}")
        return [{'文件名': filename, '错误': f'处理失败: {e}'}]
    finally:
        os.unlink(pdf_path)

def process_pdfs(pdf_files, api_key, mode, fields, language, custom_prompt):
    """(注释) 并行处理所有上传的PDF文件。"""
    import_heavy_libraries()
    all_results = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(process_single_pdf, pdf, api_key, mode, fields, language, custom_prompt) for pdf in pdf_files]
        for future in as_completed(futures):
            try:
                result_list = future.result()
                if result_list:
                    all_results.extend(result_list)
            except Exception as exc:
                print(f'❌ 执行时产生异常: {exc}')
    return all_results

# =============================================================================
# --- Excel 生成与美化 ---
# =============================================================================
def generate_excel(results):
    """(注释) 一个更通用的Excel生成函数，能处理任意列。"""
    if not results: return None
    df = pd.DataFrame(results)
    
    if '文件名' in df.columns:
        cols = df.columns.tolist()
        cols.insert(0, cols.pop(cols.index('文件名')))
        df = df[cols]
        
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        tmp.close()
        beautify_excel(tmp.name)
        return tmp.name

def beautify_excel(filepath):
    """(注释) 通用的Excel美化函数。"""
    wb = load_workbook(filepath)
    ws = wb.active
    header_fill = PatternFill(fill_type="solid", fgColor="BDD7EE")
    header_font = Font(bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    for col_cells in ws.columns:
        length = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(length + 5, 60)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin_border
    
    ws.freeze_panes = 'B1'
    wb.save(filepath)

# =============================================================================
# --- Flask 应用路由 ---
# =============================================================================
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs_route():
    # (注释) 从表单中获取所有参数
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('api_key')
    mode = request.form.get('mode', '泛读模式')
    fields = request.form.getlist('fields')
    language = request.form.get('language', '中文')
    custom_prompt = request.form.get('custom_prompt', '')
    
    if not api_key or not api_key.startswith("sk-"):
        return jsonify({"error": "API密钥为空或格式不正确"}), 400
    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400

    print(f"收到请求: {len(pdf_files)}个文件 | 模式: {mode}")
    results = process_pdfs(pdf_files, api_key, mode, fields, language, custom_prompt)
    
    if not results:
        return jsonify({"error": "未能成功处理任何文件"}), 500

    output_file_path = generate_excel(results)
    if not output_file_path:
        return jsonify({"error": "生成Excel文件失败"}), 500
        
    response = send_file(output_file_path, as_attachment=True, download_name='EggScan_Result.xlsx')
    
    @response.call_on_close
    def remove_file():
        os.unlink(output_file_path)
    
    print("✅ 所有任务完成，Excel文件已发送。")
    return response

# --- 程序主入口 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

