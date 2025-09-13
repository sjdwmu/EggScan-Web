# -*- coding: utf-8 -*-
# =============================================================================
# --- EggScan 云端分析工具 (最终修正版) ---
# =============================================================================
# (注释) 此版本已根据您的要求，将“泛读模式”恢复为提取固定结构化字段的
# (注释) 原始功能，并包含了所有必要的函数，可以直接部署。
# =============================================================================

import os
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file, jsonify

# (注释) 动态导入重量级库，加快程序启动速度
def import_heavy_libraries():
    """(注释) 动态导入重量级库，只在运行时加载"""
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

# =============================================================================
# --- 核心辅助函数 ---
# =============================================================================
def smart_extract_text(pdf_path):
    """(注释) 从PDF中提取纯文本。"""
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        return text
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

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
# --- 模式一：泛读模式 (恢复为您要的原始功能) ---
# =============================================================================
BROAD_READ_FIELDS = ["研究背景", "研究方法", "实验设计", "结果分析", "讨论"]

def call_llm_for_broad_read(pdf_text, api_key):
    """(注释) 调用LLM，提取固定的结构化字段。"""
    fields_str = "\n".join([f"- {field}" for field in BROAD_READ_FIELDS])
    prompt = f"请从以下论文内容中，分别提取如下结构化信息，每个要点请分行总结：\n{fields_str}\n\n请严格按照“【字段名】:【总结内容】”的格式输出。\n\n---\n论文内容如下:\n{pdf_text[:40000]}"
    system_prompt = "你是一个擅长将论文内容进行结构化总结的学术助手。"
    
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": prompt}], "temperature": 0.3, "max_tokens": 4096}
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: {e}"

def parse_broad_read_output(llm_text):
    """(注释) 解析泛读模式的输出。"""
    if llm_text.startswith("API_ERROR:"):
        return {field: (llm_text if i == 0 else "") for i, field in enumerate(BROAD_READ_FIELDS)}
    
    result_dict = {field: "未提取到" for field in BROAD_READ_FIELDS}
    for field in BROAD_READ_FIELDS:
        # (注释) 使用更灵活的正则表达式来匹配字段
        match = re.search(f"【{re.escape(field)}】:\s*(.*?)(?=\n【|\Z)", llm_text, re.DOTALL)
        if match:
            result_dict[field] = match.group(1).strip()
    return result_dict

# =============================================================================
# --- 模式二：精读模式 ---
# =============================================================================
def call_llm_for_deep_read(pdf_text, api_key, fields, language):
    fields_str = ", ".join(fields)
    instruction = f"作为一名顶尖科研分析师，请对以下论文进行深入的“精读”分析，并针对用户指定的每一个分析维度【{fields_str}】，进行精准、全面且高度浓缩的总结，并用【{language}】呈现。\n请严格按照“【字段名】:【总结内容】”的格式输出。"
    prompt = f"{instruction}\n\n---\n论文内容如下:\n{pdf_text[:40000]}"
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "temperature": 0.3, "max_tokens": 4096}
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e: return f"API_ERROR: {e}"

def parse_deep_read_output(llm_text, fields):
    if llm_text.startswith("API_ERROR:"): return {field: (llm_text if i == 0 else "") for i, field in enumerate(fields)}
    result_dict = {field: "未提取到" for field in fields}
    for field in fields:
        match = re.search(f"【{re.escape(field)}】:\s*(.*?)(?=\n【|\Z)", llm_text, re.DOTALL)
        if match: result_dict[field] = match.group(1).strip()
    return result_dict

# =============================================================================
# --- 模式三：自定义模式 ---
# =============================================================================
def call_llm_for_custom_mode(pdf_text, api_key, custom_prompt):
    user_content = f"{custom_prompt}\n\n---\n以下是需要分析的文本内容:\n\n{pdf_text[:40000]}"
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "user", "content": user_content}], "temperature": 0.3, "max_tokens": 4096}
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e: return f"API_ERROR: {e}"

def parse_custom_output(llm_text):
    if llm_text.startswith("API_ERROR:"): return {'错误': llm_text}
    result_dict = {}
    # (注释) 增强的解析，能匹配更多格式
    matches = re.findall(r"(?:【(.+?)】|(?<=\n)\*\*(.+?)\*\*):\s*(.*)", llm_text)
    if not matches: matches = re.findall(r"(.+?):\s*(.*)", llm_text)
    
    for match in matches:
        key = (match[0] or next((m for m in match if m), None)).strip()
        value = match[-1].strip()
        if key: result_dict[key] = value

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
            structured_data = parse_broad_read_output(llm_output)
            structured_data['文件名'] = filename
            return [structured_data]
        
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
# --- Excel 生成与 Flask 应用路由 ---
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


