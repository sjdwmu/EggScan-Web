# -*- coding: utf-8 -*-
# =============================================================================
# --- 导入核心库 ---
# =============================================================================
import os
import re
import tempfile
# (注释) 导入并发库，用于并行处理任务
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Flask Web 框架库 ---
from flask import Flask, render_template, request, send_file, jsonify

# --- 核心功能函数 ---
def import_heavy_libraries():
    """延迟导入重量级库，只在需要时才导入"""
    global fitz, convert_from_path, pytesseract, Image, requests
    global pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    
    print("正在加载分析库...")
    
    import fitz
    from pdf2image import convert_from_path
    import pytesseract
    from PIL import Image
    import requests
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    print("✓ 分析库加载完成")

LLM_URL = "https://api.deepseek.com/v1/chat/completions"

# =============================================================================
# --- (改动部分 1): 核心分析逻辑拆分与并行化 ---
# =============================================================================

def call_llm_for_analysis(pdf_text, api_key):
    """
    (注释)
    调用LLM API的核心函数，保持不变。
    """
    max_length = 30000
    truncated_text = extract_key_sections(pdf_text, max_length)

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
    system_prompt = "你是一个擅长论文分析的学术助手，请准确、精炼地提取论文中的关键信息，并严格按照用户要求的格式输出。"
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
        "temperature": 0.5,
        "max_tokens": 4096
    }
    
    try:
        print(f"    正在为片段调用LLM API...")
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=120)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        print(f"    ⚠️ API请求错误: {e}")
        return f"ERROR: API请求失败 - {e}"

def parse_llm_output_new(llm_text):
    """
    (注释)
    解析函数，保持不变。
    """
    if llm_text.startswith("ERROR:"):
        return [{'chinese_summary': llm_text, 'original_quote': '', 'quote_translation': ''}]

    results = []
    sections = llm_text.strip().split('---')
    
    print(f"    解析LLM输出，找到 {len(sections)} 个要点...")
    for section in sections:
        if not section.strip():
            continue
        
        summary_match = re.search(r"\[中文提炼\]:\s*(.*)", section, re.DOTALL)
        quote_match = re.search(r"\[核心原文\]:\s*(.*)", section, re.DOTALL)
        translation_match = re.search(r"\[原文翻译\]:\s*(.*)", section, re.DOTALL)
        
        chinese_summary = summary_match.group(1).strip() if summary_match else "未提取到"
        original_quote = quote_match.group(1).strip() if quote_match else "未提取到"
        quote_translation = translation_match.group(1).strip() if translation_match else "未提取到"
        
        results.append({
            'chinese_summary': chinese_summary,
            'original_quote': original_quote,
            'quote_translation': quote_translation
        })
        
    return results

def process_single_pdf(pdf_file, api_key):
    """
    (注释) 
    新增的函数，封装了处理单个PDF文件的所有逻辑。
    这个函数将在一个独立的线程中被执行。
    """
    filename = pdf_file.filename
    print(f"📄 开始处理: {filename}")
    
    # 将文件保存到临时文件以便处理
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_file.save(tmp.name)
        pdf_path = tmp.name
        
    try:
        text = smart_extract_text(pdf_path)
        if len(text.strip()) < 200:
            print(f"    ⚠️ {filename}: 文本太少，跳过")
            return []  # 返回空列表代表此文件无结果

        llm_output = call_llm_for_analysis(text, api_key)
        structured_data_list = parse_llm_output_new(llm_output)
        
        # 将解析出的每个要点与文件名关联
        for item in structured_data_list:
            item['文件名'] = filename
        
        print(f"    ✅ {filename}: 处理成功")
        return structured_data_list
        
    except Exception as e:
        print(f"    ❌ {filename}: 处理时出错: {e}")
        return [{'文件名': filename, '中文提炼': f'处理失败: {e}', '核心原文': '', '原文翻译': ''}]
    finally:
        # 确保临时文件被删除
        os.unlink(pdf_path)

def process_pdfs(pdf_files, api_key):
    """
    (注释)
    这是改动最大的地方：主处理函数。
    它不再是逐个处理文件，而是创建一个线程池，将所有文件的处理任务并发执行。
    """
    import_heavy_libraries()
    
    if not api_key or not api_key.strip().startswith("sk-"):
        return {'error': 'API密钥为空或格式不正确'}
        
    all_results = []
    # (注释) 创建一个最多5个线程的线程池。这意味着最多可以同时处理5个PDF文件。
    with ThreadPoolExecutor(max_workers=5) as executor:
        # (注释) 将所有文件的处理任务提交到线程池
        future_to_pdf = {executor.submit(process_single_pdf, pdf, api_key): pdf.filename for pdf in pdf_files}
        
        # (注释) as_completed会等待任何一个任务完成，然后立即处理它的结果
        for future in as_completed(future_to_pdf):
            pdf_name = future_to_pdf[future]
            try:
                result_list = future.result()
                all_results.extend(result_list)
            except Exception as exc:
                print(f'❌ 文件 {pdf_name} 在执行时产生了异常: {exc}')
                all_results.append({'文件名': pdf_name, '中文提炼': f'执行异常: {exc}', '核心原文': '', '原文翻译': ''})
                
    return all_results

# =============================================================================
# --- (改动部分 2): 更新Excel生成与美化函数 ---
# =============================================================================

def generate_excel(results):
    """
    (注释)
    Excel生成函数保持不变。
    """
    if not results:
        return None
    
    df = pd.DataFrame(results)
    column_order = ['文件名', '中文提炼', '核心原文', '原文翻译']
    df = df.reindex(columns=column_order)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        tmp.close()
        beautify_excel_new(tmp.name)
        return tmp.name

def beautify_excel_new(filepath):
    """
    (注释)
    Excel美化函数保持不变。
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

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 55
    ws.row_dimensions[1].height = 30
    
    for row in ws.iter_rows(min_row=2):
        ws.row_dimensions[row[0].row].height = 120
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
            
        print(f"    ⚠️ 文本过少 ({effective_chars} 有效字符)")
        return text
    except Exception as e:
        print(f"    ❌ 文本提取失败: {e}")
        return ""

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
    for section_name, keywords in key_sections.items():
        for keyword in keywords:
            try:
                match = re.search(r'\n\s*' + keyword + r'\s*\n', pdf_text, re.IGNORECASE)
                if match:
                    start_pos = match.start()
                    next_section_pos = len(pdf_text)
                    for next_kw_list in key_sections.values():
                        for next_kw in next_kw_list:
                            pos = pdf_text.lower().find(f'\n{next_kw}\n', start_pos + 1)
                            if pos != -1:
                                next_section_pos = min(next_section_pos, pos)
                    
                    content = pdf_text[start_pos:next_section_pos]
                    extracted_content.append(content)
                    break
            except Exception:
                continue

    final_text = "\n\n".join(extracted_content)
    if len(final_text) < 5000:
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
        
    print(f"收到请求：处理 {len(pdf_files)} 个文件，开始并行分析...")
    results = process_pdfs(pdf_files, api_key)
    
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
    
    print("✅ 所有并行任务完成，Excel文件已发送。")
    return response

# --- 程序主入口 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

