# -*- coding: utf-8 -*-
# =============================================================================
# --- EggScan 云端分析工具 (专业逻辑版) ---
# =============================================================================
# (注释) 本版本基于您本地验证成功的强大分析逻辑进行改造，
# (注释) 适配云端部署，并包含三种核心分析模式。
# =============================================================================

import os
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file, jsonify

# (注释) 动态导入重量级库，只在应用启动后，第一次请求时加载
def import_heavy_libraries():
    """动态导入重量级库"""
    global fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    if 'fitz' not in globals():
        print("正在加载核心分析库...")
        import fitz
        import requests
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        print("✓ 分析库加载成功！")

# (注释) API和常量定义
LLM_URL = "https://api.deepseek.com/v1/chat/completions"
CLASSIC_FIELDS = ["研究背景", "研究方法", "实验设计", "结果分析", "讨论"]
INTENSIVE_FIELDS = ["研究背景与缺口", "研究设计与方法", "主要结果与数据", "创新点与贡献", "局限性与批判", "可借鉴与启发"]

# =============================================================================
# --- 核心辅助函数 ---
# =============================================================================

def smart_extract_text(pdf_path):
    """(注释) 从PDF中智能提取文本"""
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        doc.close()
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

def beautify_excel_professional(filepath):
    """(注释) 这是您本地版本中使用的专业Excel美化函数"""
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        header_fill = PatternFill(fill_type="solid", fgColor="5B9BD5")
        header_font = Font(name='微软雅黑', bold=True, color="FFFFFF", size=14)
        data_font = Font(name='微软雅黑', size=12)
        title_column_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")
        banded_row_fill = PatternFill(fill_type="solid", fgColor="DDEBF7")
        thin_border = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[1].height = 30
        
        for column in ws.columns:
            max_length = 0
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max(max_length + 2, 15), 50)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
            # (注释) 网页版建议不要设置过高的固定行高，以适应不同内容长度
            # ws.row_dimensions[row_num].height = 227 
            for col_idx, cell in enumerate(row):
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = thin_border
                cell.font = data_font
                if col_idx == 0:
                    cell.fill = title_column_fill
                elif row_num % 2 != 0:
                    cell.fill = banded_row_fill
        
        ws.freeze_panes = 'A2'
        wb.save(filepath)
        print("    ✓ Excel美化完成")
    except Exception as e:
        print(f"    ⚠️ Excel美化失败: {e}")

# =============================================================================
# --- LLM 调用与解析 (基于您本地的成功代码) ---
# =============================================================================

def call_llm(api_key, system_prompt, user_prompt):
    """(注释) 统一的LLM API调用函数"""
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.1,
        "max_tokens": 4096
    }
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"API_ERROR: {e}"

def parse_llm_output(llm_text, fields):
    """(注释) 统一的LLM输出解析函数"""
    if llm_text.startswith("API_ERROR:"):
        return {field: llm_text if i == 0 else "API错误" for i, field in enumerate(fields)}
    result_dict = {field: "未提取到" for field in fields}
    for field in fields:
        pattern = rf"【{re.escape(field)}】[：:\s]*([^【]*?)(?=\n【|\Z)"
        match = re.search(pattern, llm_text, re.DOTALL)
        if match:
            result_dict[field] = match.group(1).strip()
    return result_dict

# =============================================================================
# --- 主处理流程 ---
# =============================================================================

def process_single_pdf(pdf_file, api_key, mode, language, custom_prompt):
    """(注释) 根据模式，处理单个PDF文件"""
    filename = pdf_file.filename
    print(f"📄 开始处理: {filename} (模式: {mode})")
    
    # (注释) 将上传的文件保存到临时文件以便处理
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_file.save(tmp.name)
        pdf_path = tmp.name
        
    try:
        text = smart_extract_text(pdf_path)
        if len(text.strip()) < 500:
            print(f"    ⚠️ 文本太少，跳过")
            return None

        # (注释) 根据模式选择不同的Prompt
        system_prompt = "你是一个擅长论文分析的学术助手，请准确提取论文中的关键信息。"
        lang_instruction = "请用英文输出" if language == "English" else "请用中文输出"
        
        if mode == '经典模式':
            fields = CLASSIC_FIELDS
            user_prompt = f"目标：从提供的论文内容中，提取核心的五个结构化信息。\n{lang_instruction}\n请严格按照以下格式提取关键信息（每个字段都必须填写）：\n\n【研究背景】：\n【研究方法】：\n【实验设计】：\n【结果分析】：\n【讨论】：\n\n---\n论文内容：\n{text[:40000]}"
        
        elif mode == '精读模式':
            fields = INTENSIVE_FIELDS
            system_prompt = "你是资深的学术研究专家，擅长批判性地深度解析学术论文。"
            user_prompt = f"目标：完全理解文献的来龙去脉，批判性评估其价值。\n{lang_instruction}\n请严格按照以下六个维度进行详细分析：\n\n【研究背景与缺口】：\n【研究设计与方法】：\n【主要结果与数据】：\n【创新点与贡献】：\n【局限性与批判】：\n【可借鉴与启发】：\n\n---\n论文内容：\n{text[:40000]}"

        elif mode == '自定义模式':
            system_prompt = "你是专业的学术分析助手，请根据用户要求分析文献。"
            user_prompt = f"{custom_prompt}\n\n{lang_instruction}\n\n---\n论文内容：\n{text[:40000]}"
            fields = re.findall(r'【([^】]+)】', custom_prompt)
            if not fields: # 如果用户没用括号，就尝试解析所有内容
                fields = None
        else:
            return None

        llm_output = call_llm(api_key, system_prompt, user_prompt)
        
        if fields:
            result = parse_llm_output(llm_output, fields)
        else: # (注释) 为没有预设字段的自定义模式做特殊解析
            result = {'分析结果': llm_output}

        result['文件名'] = filename
        return result
            
    except Exception as e:
        print(f"    ❌ 处理时出错: {e}")
        return {'文件名': filename, '错误': f'处理失败: {e}'}
    finally:
        os.unlink(pdf_path) # (注释) 确保删除临时文件

# =============================================================================
# --- Flask 应用与路由 ---
# =============================================================================

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs_route():
    import_heavy_libraries() # (注释) 确保在第一次请求时加载库

    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('api_key')
    mode = request.form.get('mode', '经典模式')
    language = request.form.get('language', '中文')
    custom_prompt = request.form.get('custom_prompt', '')
    
    if not api_key or not api_key.startswith("sk-"):
        return jsonify({"error": "API密钥为空或格式不正确"}), 400
    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400

    print(f"收到请求: {len(pdf_files)}个文件 | 模式: {mode}")
    all_results = []
    # (注释) 使用并行处理来加速
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(process_single_pdf, pdf, api_key, mode, language, custom_prompt) for pdf in pdf_files]
        for future in as_completed(futures):
            result = future.result()
            if result:
                all_results.append(result)
    
    if not all_results:
        return jsonify({"error": "未能成功处理任何文件"}), 500

    # (注释) 使用临时文件生成Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df = pd.DataFrame(all_results)
        # (注释) 智能排序，确保文件名总在第一列
        if '文件名' in df.columns:
            cols = df.columns.tolist()
            cols.insert(0, cols.pop(cols.index('文件名')))
            df = df[cols]
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        beautify_excel_professional(tmp.name)
        output_file_path = tmp.name

    response = send_file(output_file_path, as_attachment=True, download_name='EggScan_Analysis_Result.xlsx')
    
    # (注释) 请求结束后自动删除服务器上的临时Excel文件
    @response.call_on_close
    def remove_file():
        os.unlink(output_file_path)
    
    print("✅ 所有任务完成，Excel文件已发送。")
    return response

# (注释) 这是云端部署的入口
if __name__ == '__main__':
    # (注释) 在Render等平台上，会使用Gunicorn启动，不会直接运行这里
    # (注释) 但为了本地测试方便，保留app.run
    app.run(host='0.0.0.0', port=5000)

