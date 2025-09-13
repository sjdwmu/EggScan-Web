# -*- coding: utf-8 -*-
# =============================================================================
# --- 导入核心库 ---
# =============================================================================
import os
import re
import tempfile

# --- Flask Web 框架库 ---
from flask import Flask, render_template, request, send_file, jsonify

# (注释) 移除了 tkinter，因为它用于桌面GUI，在Web服务器上不适用且会引发错误。

# --- 核心功能函数 ---
# (注释) 将所有重量级库的导入都放在这个函数里，在程序启动时不加载，在用户请求时才加载。
def import_heavy_libraries():
    """延迟导入重量级库，只在需要时才导入"""
    global fitz, convert_from_path, pytesseract, Image, requests
    global pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    
    print("正在加载分析库...")
    
    import fitz
    # (注释) pdf2image 在非Windows服务器上部署可能需要额外配置poppler路径，这里假设环境已配置好
    from pdf2image import convert_from_path
    import pytesseract
    from PIL import Image
    import requests
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    print("✓ 分析库加载完成")

# (注释) LLM相关常量保持不变
LLM_URL = "https://api.deepseek.com/v1/chat/completions"

# =============================================================================
# --- (改动部分 1): 更新核心分析函数 ---
# =============================================================================

def call_llm_for_analysis(pdf_text, api_key):
    """
    (注释)
    这是本次升级的核心函数。
    它取代了旧的 build_prompt 和 extract_fields_with_llm 函数。
    功能：构建新的Prompt，调用LLM API，并返回原始的、未经解析的分析结果。
    """
    # (注释) 首先，对过长的文本进行智能截取，这部分逻辑保留
    max_length = 30000
    truncated_text = extract_key_sections(pdf_text, max_length)

    # (注释) 这是全新的Prompt，指导LLM按“中文提炼-核心原文-原文翻译”的格式输出
    prompt = f"""
    请你扮演一位专业的生物医学研究员，仔细阅读以下英文文献内容。
    你的任务是：
    1. 用中文精炼地总结出最重要的核心观点和发现，每一点作为一段。
    2. 在每一段中文总结下方，附上该总结所依据的最核心的1-2句英文原文。
    3. 最后，将你附上的那句“英文原文”翻译成中文。

    请严格按照以下格式返回，每个观点之间用 '---' 分隔，不要有任何多余的解释：
    [中文提炼]: 这里是你的中文总结内容。
    [核心原文]: Here is the original English quote.
    [原文翻译]: 这里是对上面那句核心原文的中文翻译。
    ---
    [中文提炼]: 这里是第二段中文总结。
    [核心原文]: Here is another key English sentence.
    [原文翻译]: 这是对第二句原文的翻译。
    """

    # (注释) System arole 保持不变，指导AI的角色
    system_prompt = "你是一个擅长论文分析的学术助手，请准确、精炼地提取论文中的关键信息，并严格按照用户要求的格式输出。"
    
    # (注释) 构造完整的请求内容
    user_content = f"请分析以下论文内容：\n\n{truncated_text}\n\n{prompt}"
    
    HEADERS = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "temperature": 0.5, # (注释) 稍微降低温度，让输出更稳定、聚焦
        "max_tokens": 4096  # (注释) 保持足够的输出空间
    }
    
    try:
        # (注释) API请求逻辑基本不变，但增加了对重试逻辑的简化
        print("    正在调用LLM API...")
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=120) # (注释) 延长超时时间
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        print(f"    ⚠️ API请求错误: {e}")
        # (注释) 如果请求失败，返回一个特殊的错误字符串，方便后续处理
        return f"ERROR: API请求失败 - {e}"

def parse_llm_output_new(llm_text):
    """
    (注释)
    这是新增的解析函数，用于处理新格式的LLM返回结果。
    它取代了旧的 parse_llm_output 函数。
    """
    # (注释) 检查返回的是否是错误信息
    if llm_text.startswith("ERROR:"):
        # (注释) 如果是错误，返回一个包含错误信息的列表
        return [{'chinese_summary': llm_text, 'original_quote': '', 'quote_translation': ''}]

    results = []
    # (注释) 使用 '---' 作为分隔符，将LLM返回的多个要点分割成列表
    sections = llm_text.strip().split('---')
    
    print(f"    解析LLM输出，找到 {len(sections)} 个要点...")
    for section in sections:
        if not section.strip():
            continue
        
        # (注释) 使用正则表达式安全地提取每个部分的内容
        summary_match = re.search(r"\[中文提炼\]:\s*(.*)", section, re.DOTALL)
        quote_match = re.search(r"\[核心原文\]:\s*(.*)", section, re.DOTALL)
        translation_match = re.search(r"\[原文翻译\]:\s*(.*)", section, re.DOTALL)
        
        # (注释) .strip() 用于去除可能存在的前后多余空格或换行符
        chinese_summary = summary_match.group(1).strip() if summary_match else "未提取到"
        original_quote = quote_match.group(1).strip() if quote_match else "未提取到"
        quote_translation = translation_match.group(1).strip() if translation_match else "未提取到"
        
        results.append({
            'chinese_summary': chinese_summary,
            'original_quote': original_quote,
            'quote_translation': quote_translation
        })
        
    return results

# =============================================================================
# --- (改动部分 2): 更新主处理流程和Excel生成 ---
# =============================================================================

def process_pdfs(pdf_files, api_key):
    """
    (注释)
    更新主处理函数，调用新的分析和解析逻辑。
    """
    import_heavy_libraries()
    
    if not api_key or not api_key.strip().startswith("sk-"):
        print("❌ API密钥格式不正确")
        # (注释) 直接返回一个包含错误信息的字典，让前端知道问题
        return {'error': 'API密钥为空或格式不正确'}
        
    all_results = []
    
    for idx, pdf_file in enumerate(pdf_files, 1):
        filename = pdf_file.filename
        print(f"📄 [{idx}/{len(pdf_files)}] 正在处理: {filename}")
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            pdf_file.save(tmp.name)
            pdf_path = tmp.name
            
        try:
            text = smart_extract_text(pdf_path)
            if len(text.strip()) < 200:
                print("    ⚠️ 提取的文本太少，跳过此文件")
                continue
            
            # (注释) 调用新的分析函数
            llm_output = call_llm_for_analysis(text, api_key)
            # (注释) 调用新的解析函数
            structured_data_list = parse_llm_output_new(llm_output)
            
            # (注释) 将解析出的每个要点与文件名关联，并添加到总结果中
            for item in structured_data_list:
                item['文件名'] = filename
                all_results.append(item)
            
            print("    ✅ 处理成功\n")
            
        except Exception as e:
            print(f"    ❌ 处理文件时出错: {e}\n")
            # (注释) 如果处理过程中出现意外错误，也记录下来
            all_results.append({
                '文件名': filename,
                '中文提炼': f'处理失败: {e}',
                '核心原文': '',
                '原文翻译': ''
            })
        finally:
            os.unlink(pdf_path)
            
    return all_results

def generate_excel(results):
    """
    (注释)
    更新Excel生成函数，以适应新的数据结构和列名。
    """
    if not results:
        return None
    
    df = pd.DataFrame(results)
    
    # (注释) 定义新的列名和顺序
    column_order = ['文件名', '中文提炼', '核心原文', '原文翻译']
    # (注释) 筛选数据，确保即使有错误列也能正常生成
    df = df.reindex(columns=column_order)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        tmp.close()
        beautify_excel_new(tmp.name) # (注释) 调用新的美化函数
        return tmp.name

def beautify_excel_new(filepath):
    """
    (注释)
    新的Excel美化函数，根据新的列宽进行调整。
    """
    wb = load_workbook(filepath)
    ws = wb.active
    
    header_fill = PatternFill(fill_type="solid", fgColor="BDD7EE")
    header_font = Font(bold=True, color="000000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # (注释) 设置列宽
    ws.column_dimensions['A'].width = 30  # 文件名
    ws.column_dimensions['B'].width = 55  # 中文提炼
    ws.column_dimensions['C'].width = 55  # 核心原文
    ws.column_dimensions['D'].width = 55  # 原文翻译
    ws.row_dimensions[1].height = 30
    
    for row in ws.iter_rows(min_row=2):
        ws.row_dimensions[row[0].row].height = 120 # (注释) 增加行高以容纳更多内容
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin_border
            
    wb.save(filepath)
    print("    ✅ Excel格式美化完成")


# =============================================================================
# --- (未改动部分): 保留了大部分的PDF文本提取和辅助函数 ---
# =============================================================================
def clean_bullet(text):
    text = re.sub(r'^[\s*\-*•·#]+', '', text, flags=re.MULTILINE)
    text = text.replace('**', '').replace('*', '').replace('###', '').replace('##', '').replace('#', '')
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def smart_extract_text(pdf_path, min_chars=1000):
    print(f"    尝试直接提取文本...")
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        effective_chars = len(''.join(text.split()))
        
        if effective_chars >= min_chars:
            print(f"    ✅ 文本提取成功 ({len(text)} 字符)")
            return text
            
        print(f"    ⚠️ 文本过少 ({effective_chars} 有效字符)，可能需要OCR...")
        # (注释) 在Web服务器环境下，OCR依赖复杂且耗时，暂时简化逻辑，优先使用文本提取
        # (注释) 如果文本提取效果不佳，可以考虑后续为OCR功能增加专门的配置
        # return ocr_from_pdf(pdf_path) 
        return text # 即使文本少，也先返回
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

# (注释) ocr_from_pdf 函数暂时保留，但在 smart_extract_text 中被注释掉了，以简化服务器部署
def ocr_from_pdf(pdf_path):
    try:
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
        return ""

def extract_key_sections(pdf_text, max_length=30000):
    if len(pdf_text) <= max_length:
        return pdf_text
    print(f"    文本过长（{len(pdf_text)}字符），智能提取关键内容...")
    
    key_sections = {
        '摘要': ['abstract', 'summary'],
        '引言': ['introduction', 'background'],
        '方法': ['method', 'materials and methods'],
        '结果': ['result', 'findings'],
        '讨论': ['discussion', 'analysis'],
        '结论': ['conclusion']
    }
    
    extracted_content = []
    # (注释) 简化逻辑：优先提取摘要、结论、引言和讨论
    for section_name, keywords in key_sections.items():
        for keyword in keywords:
            try:
                # (注释) 使用正则表达式查找以换行符开头的关键词，更准确
                match = re.search(r'\n\s*' + keyword + r'\s*\n', pdf_text, re.IGNORECASE)
                if match:
                    start_pos = match.start()
                    # (注释) 寻找下一个章节标题作为结束位置
                    next_section_pos = len(pdf_text)
                    for next_kw_list in key_sections.values():
                        for next_kw in next_kw_list:
                            pos = pdf_text.lower().find(f'\n{next_kw}\n', start_pos + 1)
                            if pos != -1:
                                next_section_pos = min(next_section_pos, pos)
                    
                    content = pdf_text[start_pos:next_section_pos]
                    extracted_content.append(content)
                    break # 找到一个关键词就跳出
            except Exception:
                continue

    final_text = "\n\n".join(extracted_content)
    if len(final_text) < 5000: # (注释) 如果提取的部分太少，就用截断的方式
        final_text = pdf_text[:max_length]

    print(f"    智能提取完成，保留了 {len(final_text)} 字符")
    return final_text


# =============================================================================
# --- Flask 应用初始化与路由 (这部分基本不变) ---
# =============================================================================
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs():
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('api_key')
    
    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400
        
    print(f"收到请求：处理 {len(pdf_files)} 个文件")
    results = process_pdfs(pdf_files, api_key)
    
    # (注释) 检查是否是API Key错误
    if isinstance(results, dict) and 'error' in results:
        return jsonify(results), 400

    if not results:
        return jsonify({"error": "未能成功处理任何文件，请检查文件内容"}), 500

    output_file_path = generate_excel(results)
    if not output_file_path:
        return jsonify({"error": "生成Excel文件失败"}), 500
        
    response = send_file(output_file_path, as_attachment=True, download_name='EggScan_Result_Updated.xlsx')
    
    @response.call_on_close
    def remove_file():
        try:
            os.unlink(output_file_path)
        except Exception as e:
            print(f"删除临时文件失败: {e}")
    
    print("✅ Excel文件已发送，任务完成。")
    return response

# --- 程序主入口 ---
if __name__ == '__main__':
    # (注释) 移除 `import_heavy_libraries()` 调用，因为它应该在请求时被调用，而不是启动时
    app.run(host='0.0.0.0', port=5000, debug=True) # (注释) 建议在开发时开启debug模式

