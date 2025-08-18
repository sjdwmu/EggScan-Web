# -*- coding: utf-8 -*-
# =============================================================================
# --- 导入核心库 ---
# =============================================================================
import os
import sys
import subprocess
import importlib.util
import re
import configparser
import tempfile
from pathlib import Path

# --- Flask Web 框架库 ---
from flask import Flask, render_template, request, send_file, jsonify

# --- API库（Flask输入需要，优先导入）
import tkinter as tk
from tkinter import simpledialog, messagebox

# --- 核心功能函数 ---
# 延迟导入重量级库（加快启动速度）
def import_heavy_libraries():
    """延迟导入重量级库，只在需要时才导入"""
    global fitz, convert_from_path, pytesseract, Image, requests
    global pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    
    # 打印加载信息
    print("正在加载分析库...")
    
    # 这里我们只导入 Python 库，不依赖本地可执行文件
    import fitz
    from pdf2image import convert_from_path
    import pytesseract
    from PIL import Image
    import requests
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    print("✓ 分析库加载完成")

# LLM相关常量
LLM_URL = "https://api.deepseek.com/v1/chat/completions"
FIELDS = ["研究背景", "研究方法", "实验设计", "结果分析", "讨论"]

# 核心逻辑函数
def clean_bullet(text):
    """清理文本中的多余符号"""
    text = re.sub(r'^[\s*\-*•·#]+', '', text, flags=re.MULTILINE)
    text = text.replace('**', '').replace('*', '').replace('###', '').replace('##', '').replace('#', '')
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def smart_extract_text(pdf_path, min_chars=1000):
    """智能提取PDF文本，不足时尝试OCR"""
    print(f"    尝试直接提取文本...")
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        effective_chars = len(''.join(text.split()))
        zh_count = sum('\u4e00' <= c <= '\u9fff' for c in text)
        en_count = sum(c.isalpha() for c in text)
        print(f"    提取到 {len(text)} 个字符（有效字符: {effective_chars}，中文: {zh_count}，英文: {en_count}）")
        
        if effective_chars >= min_chars or (effective_chars > 500 and (zh_count > 100 or en_count > 300)):
            print(f"    ✅ 文本提取成功")
            return text
            
        pages = len(doc)
        avg_chars_per_page = effective_chars / pages if pages > 0 else 0
        if avg_chars_per_page > 200:
            print(f"    ✅ 文本提取成功（每页平均 {avg_chars_per_page:.0f} 个字符）")
            return text
            
        print(f"    ⚠️ 文本过少，切换到OCR模式...")
        print(f"    开始OCR识别（这可能需要较长时间）...")
        return ocr_from_pdf(pdf_path)
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

def ocr_from_pdf(pdf_path):
    """使用OCR识别PDF中的文本"""
    try:
        # 在Web应用中，我们不依赖本地可执行文件，
        # 而是假设服务器环境已经安装了tesseract和poppler
        images = convert_from_path(pdf_path, dpi=200)
        text_all = ""
        total_pages = len(images)
        for idx, img in enumerate(images, 1):
            print(f"    正在识别第 {idx}/{total_pages} 页...")
            t = pytesseract.image_to_string(img, lang='chi_sim+eng')
            text_all += f"\n---- 第{idx}页 ----\n{t}\n"
        print(f"    ✅ OCR识别完成")
        return text_all
    except Exception as e:
        print(f"    ❌ OCR错误: {e}")
        # 这里不抛出异常，而是返回空字符串，确保程序继续运行
        return ""

def extract_key_sections(pdf_text, max_length=30000):
    """智能提取文本关键部分以节省LLM token"""
    if len(pdf_text) <= max_length:
        return pdf_text
    print(f"    文本过长（{len(pdf_text)}字符），智能提取关键内容...")
    
    key_sections = {
        '摘要': ['摘要', 'abstract', 'summary', '概要'],
        '引言': ['引言', 'introduction', '前言', '背景', 'background'],
        '方法': ['方法', 'method', 'methodology', '材料与方法', 'materials and methods', '实验方法'],
        '结果': ['结果', 'result', 'findings', '实验结果', '研究结果'],
        '讨论': ['讨论', 'discussion', '分析', 'analysis'],
        '结论': ['结论', 'conclusion', '总结', 'summary and conclusion']
    }
    
    extracted_parts = []
    used_length = 0
    header_length = min(5000, len(pdf_text))
    extracted_parts.append(("开头部分", pdf_text[:header_length]))
    used_length += header_length
    
    text_lower = pdf_text.lower()
    found_sections = []
    
    for section_name, keywords in key_sections.items():
        for keyword in keywords:
            positions = []
            start = 0
            while True:
                pos = text_lower.find(keyword.lower(), start)
                if pos == -1:  
                    break
                if pos == 0 or pdf_text[pos-1] in '\n\r':
                    positions.append(pos)
                start = pos + 1
            
            if positions:
                section_start = positions[0]
                section_end = min(section_start + 5000, len(pdf_text))
                section_content = pdf_text[section_start:section_end]
                found_sections.append((section_name, section_start, section_content))
                break
    
    found_sections.sort(key=lambda x: x[1])
    
    for section_name, _, content in found_sections:
        if used_length + len(content) <= max_length:
            extracted_parts.append((section_name, content))
            used_length += len(content)
            print(f"    ✓ 提取了 {section_name} 部分")
    
    if used_length < max_length - 3000:
        tail_length = min(3000, max_length - used_length)
        extracted_parts.append(("结尾部分", pdf_text[-tail_length:]))
        print(f"    ✓ 提取了结尾部分")
    
    result = []
    for name, content in extracted_parts:
        result.append(f"\n\n===== {name} =====\n")
        result.append(content)
    
    final_text = "".join(result)
    print(f"    智能提取完成，保留了 {len(final_text)} 字符")
    return final_text

def build_prompt(pdf_text):
    """构建发送给LLM的Prompt"""
    max_length = 30000
    pdf_text = extract_key_sections(pdf_text, max_length)
    instruction = "请从以下论文内容中分别提取如下结构化信息：\n"
    instruction += "".join([f"- {field}\n" for field in FIELDS])
    instruction += "\n每个要点请分行，正文如下：\n\n"
    return instruction + pdf_text

def extract_fields_with_llm(pdf_text, api_key):
    """调用LLM API进行分析"""
    prompt = build_prompt(pdf_text)
    HEADERS = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "你是一个擅长论文分析的学术助手，请准确提取论文中的关键信息"},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7,
        "max_tokens": 4000
    }
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=60)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.Timeout:
        print("    ⚠️ API请求超时，尝试使用更短的文本...")
        shorter_text = extract_key_sections(pdf_text, 15000)
        prompt = build_prompt(shorter_text)
        payload["messages"][1]["content"] = prompt
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=60)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"    ⚠️ API请求错误: {e}")
        raise

def parse_llm_output(text):
    """解析LLM返回的文本"""
    # 修复：移除LLM输出中可能出现的开头和结尾的额外方括号
    text = text.strip()
    if text.startswith('【') and text.endswith('】'):
        text = text[1:-1]
    
    result = {field: "" for field in FIELDS}
    text = text.replace('###', '').replace('##', '')
    for field in FIELDS:
        if field in text:
            start = text.find(field)
            next_field_pos = [text.find(f) for f in FIELDS if f != field and text.find(f) > start]
            end = min(next_field_pos) if next_field_pos else len(text)
            content = text[start + len(field):end].strip()
            if content.startswith(':') or content.startswith('：'):
                content = content[1:].strip()
            result[field] = clean_bullet(content)
    return result

def beautify_excel(filepath):
    """美化Excel文件格式"""
    wb = load_workbook(filepath)
    ws = wb.active
    
    header_fill = PatternFill(fill_type="solid", fgColor="BDD7EE")
    header_font = Font(bold=True, color="000000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='808080'),
        right=Side(style='thin', color='808080'),
        top=Side(style='thin', color='808080'),
        bottom=Side(style='thin', color='808080')
    )
    medium_border = Border(
        left=Side(style='medium', color='000000'),
        right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'),
        bottom=Side(style='medium', color='000000')
    )
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = medium_border
    
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin_border
    
    for column_cells in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        col_letter = column_cells[0].column_letter
        ws.column_dimensions[col_letter].width = min(max_length + 6, 55)
    
    ws.row_dimensions[1].height = 30
    for row_num in range(2, ws.max_row + 1):
        ws.row_dimensions[row_num].height = 100
    
    wb.save(filepath)
    print("    ✅ Excel格式美化完成")

def process_pdfs(pdf_files, api_key):
    """处理上传的PDF文件列表"""
    import_heavy_libraries() # 确保库已加载
    
    if not api_key.strip().startswith("sk-"):
        print("❌ API密钥格式不正确")
        return [{"文件名": "API密钥格式不正确"}]
        
    HEADERS = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    RESULTS = []
    
    for idx, pdf_file in enumerate(pdf_files, 1):
        filename = pdf_file.filename
        print(f"📄 [{idx}/{len(pdf_files)}] 正在处理: {filename}")
        
        # 将文件保存到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            pdf_file.save(tmp.name)
            pdf_path = tmp.name
            
        try:
            text = smart_extract_text(pdf_path)
            if len(text.strip()) < 100:
                print("    ⚠️ 提取的文本太少，跳过此文件")
                continue
            
            print("    正在调用LLM分析...")
            llm_output = extract_fields_with_llm(text, api_key)
            structured_data = parse_llm_output(llm_output)
            structured_data["文件名"] = filename
            RESULTS.append(structured_data)
            print("    ✅ 处理成功\n")
            
        except Exception as e:
            print(f"    ❌ 处理文件时出错: {e}\n")
            error_data = {field: "处理失败" for field in FIELDS}
            error_data["文件名"] = filename
            error_data["错误信息"] = str(e)
            RESULTS.append(error_data)
            
        finally:
            os.unlink(pdf_path) # 删除临时文件
            
    return RESULTS

def generate_excel(results):
    """将结果生成Excel文件"""
    if not results:
        return None
    
    # 使用临时文件保存Excel，然后返回
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df = pd.DataFrame(results)
        df.to_excel(tmp.name, index=False)
        tmp.close()
        beautify_excel(tmp.name)
        return tmp.name
        
# =============================================================================
# --- Flask 应用初始化 ---
# =============================================================================
app = Flask(__name__)

# --- 路由定义 ---
@app.route('/')
def index():
    """主页，返回上传文件界面"""
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs():
    """处理PDF分析请求"""
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('api_key')
    
    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400
        
    print(f"收到请求：处理 {len(pdf_files)} 个文件")
    results = process_pdfs(pdf_files, api_key)
    
    if not results:
        return jsonify({"error": "处理失败，请检查API密钥或文件内容"}), 500

    output_file_path = generate_excel(results)
    if not output_file_path:
        return jsonify({"error": "生成Excel文件失败"}), 500
        
    response = send_file(output_file_path, as_attachment=True, download_name='EggScan_Result.xlsx')
    
    # 删除Excel临时文件
    @response.call_on_close
    def remove_file():
        os.unlink(output_file_path)
    
    print("✅ Excel文件已发送，任务完成。")
    return response

# --- 程序主入口 ---
if __name__ == '__main__':
    # 延迟导入重量级库
    import_heavy_libraries()
    
    # 运行Flask应用，host='0.0.0.0'让局域网内其他设备可访问
    app.run(host='0.0.0.0', port=5000)
