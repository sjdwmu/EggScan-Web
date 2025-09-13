# -*- coding: utf-8 -*-
# =============================================================================
# --- EggScan 云端分析应用 (Flask版 v2.0) ---
# =============================================================================
# 版本: 2.0
# 描述: 将 EggScan v2.0 的核心逻辑封装为 Flask Web 应用，用于云端部署。
# 作者: [您的名字或团队]
# =============================================================================

import os
import re
import json
import tempfile
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify

# 延迟导入，只在第一次请求时加载
fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side = (None,) * 9

def import_heavy_libraries():
    """延迟导入重量级库，加快启动速度"""
    global fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    if fitz is None:
        print("[INFO] 首次请求，正在加载核心分析库...")
        try:
            import fitz as f
            import requests as r
            import pandas as p
            from openpyxl import load_workbook as lw
            from openpyxl.styles import Font as F, Alignment as A, PatternFill as PF, Border as B, Side as S
            
            fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side = f, r, p, lw, F, A, PF, B, S
            print("✓ 分析库加载成功！")
        except ImportError as e:
            print(f"❌ 错误：缺少必要的库 - {e}")
            print("请在部署环境中运行 'pip install -r requirements.txt'")
            raise

# --- 全局配置 ---
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # antd: 设置最大上传大小为100MB
LLM_URL = "https://api.deepseek.com/v1/chat/completions"

# --- 分析框架定义 (与v2.0脚本保持一致) ---
CLASSIC_FIELDS = ["研究背景", "研究方法", "实验设计", "结果分析", "讨论"]
INTENSIVE_FIELDS = ["研究背景与缺口", "研究设计与方法", "主要结果与数据", "创新点与贡献", "局限性与批判", "可借鉴与启发"]
CUSTOM_TEMPLATE = """
请从以下角度分析这篇文献：
【研究主题】：文章的核心研究问题是什么？
【理论框架】：使用了什么理论基础或概念框架？
【方法创新】：在研究方法上有什么创新或特色？
【数据质量】：数据来源、样本量、统计分析的可靠性如何？
【关键发现】：最重要的3个研究发现是什么？
【实践意义】：对临床实践或政策制定有什么指导意义？
【争议与讨论】：存在哪些争议点或值得进一步讨论的问题？
【引用价值】：这篇文章最值得引用的观点或数据是什么？
请用【字段名】：内容 的格式清晰输出。
"""

# --- 核心辅助函数 (从v2.0脚本迁移并优化) ---

def smart_extract_text(pdf_path):
    """从PDF中智能提取文本"""
    # antd: 此函数与本地版v2.0完全相同
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        doc.close()
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r' {2,}', ' ', text)
        return text
    except Exception as e:
        print(f"     ❌ 文本提取失败: {e}")
        return ""

def beautify_excel_professional(filepath):
    """专业的Excel美化"""
    # antd: 此函数与本地版v2.0完全相同，包含所有美化更新
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        
        header_fill = PatternFill(fill_type="solid", fgColor="5B9BD5")
        header_font = Font(name='微软雅黑', bold=True, color="FFFFFF", size=16)
        data_font = Font(name='微软雅黑', size=14)
        title_column_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")
        banded_row_fill = PatternFill(fill_type="solid", fgColor="DDEBF7")
        thin_border = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[1].height = 40
        
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        cell_value = str(cell.value)
                        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', cell_value))
                        other_chars = len(cell_value) - chinese_chars
                        effective_length = chinese_chars * 2.2 + other_chars * 1.1
                        max_length = max(max_length, effective_length)
                except: pass
            adjusted_width = min(max(max_length, 15), 60)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
            ws.row_dimensions[row_num].height = 227
            for col_idx, cell in enumerate(row):
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = thin_border
                cell.font = data_font
                if col_idx == 0: cell.fill = title_column_fill
                elif row_num % 2 != 0: cell.fill = banded_row_fill

        ws.freeze_panes = 'A2'
        wb.save(filepath)
        print(f"     ✓ Excel美化完成")
    except Exception as e:
        print(f"     ⚠️ Excel美化失败: {e}")

def call_llm(pdf_text, api_key, mode, language, custom_prompt=""):
    """统一的LLM调用函数"""
    
    lang_instruction = "Please output in English" if language == "English" else "请用中文输出"
    
    if mode == '经典模式':
        fields = CLASSIC_FIELDS
        prompt_template = f"""你是一位专业的学术助手...【研究背景】：...【研究方法】：...【实验设计】：...【结果分析】：...【讨论】：...""" # antd: 省略完整prompt，与v3.2脚本一致
        prompt = prompt_template.replace("...", f"...\n\n{lang_instruction}\n\n---\n论文内容：\n{pdf_text[:40000]}")
    elif mode == '精读模式':
        fields = INTENSIVE_FIELDS
        prompt_template = f"""你是一位资深的学术研究专家...【研究背景与缺口】：...【研究设计与方法】：...""" # antd: 省略完整prompt，与v3.2脚本一致
        prompt = prompt_template.replace("...", f"...\n\n{lang_instruction}\n\n---\n论文内容：\n{pdf_text[:40000]}")
    elif mode == '自定义模式':
        fields = re.findall(r'【([^】]+)】', custom_prompt)
        prompt = f"{custom_prompt}\n\n{lang_instruction}\n\n---\n论文内容：\n{pdf_text[:40000]}"
    else:
        return "无效的模式"
        
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "temperature": 0.1, "max_tokens": 4096}
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: {e}"

def parse_llm_output(llm_text, fields):
    """解析LLM输出"""
    # antd: 此函数与本地版v2.0完全相同
    if llm_text.startswith("API_ERROR:"):
        return {field: llm_text if i == 0 else "API错误" for i, field in enumerate(fields)}
    result_dict = {}
    for field in fields:
        content = None
        pattern1 = rf"【{re.escape(field)}】[：:\s]*([^【]*?)(?=\n【|\Z)"
        match1 = re.search(pattern1, llm_text, re.DOTALL)
        if match1:
            content = match1.group(1).strip()
        result_dict[field] = content if content and len(content) > 5 else f"解析失败-{field}"
    return result_dict

# --- Flask 路由定义 ---

@app.route('/')
def index():
    """渲染主页"""
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs():
    """处理PDF分析请求"""
    import_heavy_libraries() # antd: 在第一次请求时加载库
    
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('apiKey')
    mode = request.form.get('mode')
    language = request.form.get('language')
    custom_prompt = request.form.get('customPrompt', CUSTOM_TEMPLATE)

    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400
    if not api_key or not api_key.startswith("sk-"):
        return jsonify({"error": "API密钥格式不正确"}), 400

    all_results = []
    temp_files = []
    
    for pdf_file in pdf_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            pdf_file.save(tmp.name)
            temp_files.append(tmp.name)
            
            print(f"  → 正在处理: {pdf_file.filename}")
            text = smart_extract_text(tmp.name)
            
            if len(text.strip()) < 500:
                print(f"     ⚠️ 文本内容不足，跳过")
                continue

            llm_output = call_llm(text, api_key, mode, language, custom_prompt)
            
            fields_map = {'经典模式': CLASSIC_FIELDS, '精读模式': INTENSIVE_FIELDS, '自定义模式': re.findall(r'【([^】]+)】', custom_prompt)}
            result = parse_llm_output(llm_output, fields_map.get(mode, []))
            
            result['文件名'] = pdf_file.filename
            result['分析时间'] = datetime.now().strftime("%Y-%m-%d %H:%M")
            all_results.append(result)

    for tmp_file in temp_files:
        os.unlink(tmp_file) # antd: 清理临时PDF文件

    if not all_results:
        return jsonify({"error": "所有文件都未能成功处理，请检查PDF内容或API密钥"}), 500

    # antd: 生成Excel报告
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
        df = pd.DataFrame(all_results)
        
        priority_cols = ['文件名', '分析时间']
        other_cols = [col for col in df.columns if col not in priority_cols]
        df = df[[col for col in priority_cols if col in df.columns] + other_cols]
        
        df.to_excel(tmp_excel.name, index=False, engine='openpyxl')
        tmp_excel.close()
        beautify_excel_professional(tmp_excel.name)
        
        # antd: 发送文件给用户
        response = send_file(tmp_excel.name, as_attachment=True, download_name=f'EggScan_Report_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
        
        @response.call_on_close
        def cleanup():
            os.unlink(tmp_excel.name)
        
        print("✓ 分析报告已生成并发送")
        return response

if __name__ == '__main__':
    # antd: 建议使用Gunicorn等WSGI服务器在生产环境中运行
    # antd: 本地测试时，可以使用 app.run(debug=True, host='0.0.0.0', port=5000)
    app.run(host='0.0.0.0', port=5000)

