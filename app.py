# -*- coding: utf-8 -*-
# =============================================================================
# --- EggScan 云端分析工具 (修正版) ---
# =============================================================================

import os
import re
import tempfile
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file, jsonify

# 全局变量声明
fitz = None
requests = None
pd = None
load_workbook = None
Font = None
Alignment = None
PatternFill = None
Border = None
Side = None

def import_heavy_libraries():
    """动态导入重量级库"""
    global fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side
    if fitz is None:
        print("正在加载核心分析库...")
        import fitz as _fitz
        import requests as _requests
        import pandas as _pd
        from openpyxl import load_workbook as _load_workbook
        from openpyxl.styles import Font as _Font, Alignment as _Alignment, PatternFill as _PatternFill, Border as _Border, Side as _Side
        
        fitz = _fitz
        requests = _requests
        pd = _pd
        load_workbook = _load_workbook
        Font = _Font
        Alignment = _Alignment
        PatternFill = _PatternFill
        Border = _Border
        Side = _Side
        print("✓ 分析库加载成功！")

# API和常量定义
LLM_URL = "https://api.deepseek.com/v1/chat/completions"

# 泛读框架（快速筛选）
SKIMMING_FIELDS = ["研究问题", "核心论点", "研究方法", "关键结论", "相关性评估"]

# 精读框架（深度分析）
INTENSIVE_FIELDS = ["研究背景与缺口", "研究设计与方法", "主要结果与数据", "创新点与贡献", "局限性与批判", "可借鉴与启发"]

# 自定义模板
CUSTOM_TEMPLATE = """
请从以下角度分析这篇文献：
【研究主题】：文章的核心研究问题是什么？
【理论框架】：使用了什么理论基础？
【方法创新】：研究方法上有什么创新？
【数据质量】：数据来源和统计分析的可靠性如何？
【关键发现】：最重要的3个研究发现是什么？
【实践意义】：对实践有什么指导意义？

请用【字段名】：内容 的格式输出。
"""

# =============================================================================
# --- 核心辅助函数 ---
# =============================================================================

def smart_extract_text(pdf_path):
    """从PDF中智能提取文本"""
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text() for page in doc])
        doc.close()
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r' {2,}', ' ', text)
        return text
    except Exception as e:
        print(f"❌ 文本提取失败: {e}")
        return ""

def beautify_excel_professional(filepath):
    """专业的Excel美化"""
    try:
        wb = load_workbook(filepath)
        ws = wb.active
        
        header_fill = PatternFill(fill_type="solid", fgColor="5B9BD5")
        header_font = Font(name='微软雅黑', bold=True, color="FFFFFF", size=11)
        data_font = Font(name='微软雅黑', size=10)
        
        thin_border = Border(
            left=Side(style='thin', color='B4C6E7'),
            right=Side(style='thin', color='B4C6E7'),
            top=Side(style='thin', color='B4C6E7'),
            bottom=Side(style='thin', color='B4C6E7')
        )
        
        # 设置表头样式
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[1].height = 30
        
        # 自动调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        cell_value = str(cell.value)
                        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', cell_value))
                        other_chars = len(cell_value) - chinese_chars
                        effective_length = chinese_chars * 2 + other_chars
                        max_length = max(max_length, effective_length)
                except:
                    pass
            adjusted_width = min(max(max_length * 0.8, 10), 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 设置数据区域样式
        for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
            ws.row_dimensions[row_num].height = 60  # 适中的行高
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = thin_border
                cell.font = data_font
                if row_num % 2 == 0:
                    cell.fill = PatternFill(fill_type="solid", fgColor="F2F2F2")
        
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        wb.save(filepath)
        print("✓ Excel美化完成")
        
    except Exception as e:
        print(f"⚠️ Excel美化失败: {e}")

# =============================================================================
# --- LLM调用函数 ---
# =============================================================================

def call_llm_for_mode(pdf_text, api_key, mode, language):
    """根据模式调用LLM"""
    
    # 确保requests已导入
    if requests is None:
        import_heavy_libraries()
    
    lang_instruction = "Please output in English" if language == "English" else "请用中文输出"
    
    # 根据模式构建prompt
    if mode == '泛读模式' or mode == '经典五段式':
        prompt = f"""
你是一位专业的文献筛选专家，请对这篇论文进行快速泛读分析。
目标：快速判断文献的相关性和核心价值。

{lang_instruction}

请严格按照以下格式提取关键信息（每个字段必须填写）：

【研究问题】：这篇文章具体想回答什么问题？
【核心论点】：作者最核心的观点是什么？（一句话总结）
【研究方法】：这是什么类型的研究？
【关键结论】：最重要的研究结论是什么？
【相关性评估】：评估其研究价值（高相关/中相关/低相关）

---
论文内容：
{pdf_text[:25000]}
"""
        fields = SKIMMING_FIELDS
        
    elif mode == '精读模式':
        prompt = f"""
你是一位资深的学术研究专家，请对这篇论文进行深度精读分析。

{lang_instruction}

请按照以下六个维度进行详细分析（每个维度至少3-5句话）：

【研究背景与缺口】：详细阐述研究背景和空白
【研究设计与方法】：包括样本量、分组、统计方法等
【主要结果与数据】：关键数据和图表引用
【创新点与贡献】：理论/方法/实践创新
【局限性与批判】：作者承认的+你发现的问题
【可借鉴与启发】：可直接借鉴的方法和研究思路

---
论文内容：
{pdf_text[:35000]}
"""
        fields = INTENSIVE_FIELDS
        
    elif mode == '自定义模式':
        prompt = f"""
{CUSTOM_TEMPLATE}

{lang_instruction}

---
论文内容：
{pdf_text[:30000]}
"""
        fields = re.findall(r'【([^】]+)】', CUSTOM_TEMPLATE)
    else:
        return None, None
    
    # 构建请求
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "你是专业的学术分析助手。"},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,
        "max_tokens": 3000  # 减少token数量加快响应
    }
    
    try:
        print(f"  → 正在调用DeepSeek API...")
        response = requests.post(
            LLM_URL,
            headers=headers,
            json=payload,
            timeout=60  # 60秒超时
        )
        response.raise_for_status()
        result = response.json()["choices"][0]["message"]["content"]
        print(f"  ✓ API调用成功")
        return result, fields
    except Exception as e:
        print(f"  ❌ API调用失败: {e}")
        return f"API_ERROR: {e}", fields

def parse_llm_output(llm_text, fields):
    """解析LLM输出"""
    if llm_text.startswith("API_ERROR:"):
        return {field: llm_text if i == 0 else "API错误" for i, field in enumerate(fields)}
    
    result_dict = {}
    for field in fields:
        pattern = rf"【{re.escape(field)}】[：:\s]*([^【]*?)(?=\n【|\Z)"
        match = re.search(pattern, llm_text, re.DOTALL)
        if match:
            content = match.group(1).strip()
            result_dict[field] = content if content else "未提取到"
        else:
            result_dict[field] = "未提取到"
    
    return result_dict

# =============================================================================
# --- 处理单个PDF ---
# =============================================================================

def process_single_pdf(pdf_file, api_key, mode, language):
    """处理单个PDF文件"""
    filename = pdf_file.filename
    print(f"📄 处理文件: {filename}")
    
    # 保存临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_file.save(tmp.name)
        pdf_path = tmp.name
    
    try:
        # 提取文本
        text = smart_extract_text(pdf_path)
        if len(text.strip()) < 500:
            print(f"  ⚠️ 文本内容太少，跳过")
            return None
        
        # 调用LLM
        llm_output, fields = call_llm_for_mode(text, api_key, mode, language)
        
        if fields:
            result = parse_llm_output(llm_output, fields)
        else:
            result = {'分析结果': llm_output}
        
        result['文件名'] = filename
        result['分析时间'] = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        return result
        
    except Exception as e:
        print(f"  ❌ 处理失败: {e}")
        return {'文件名': filename, '错误': str(e)}
    finally:
        # 清理临时文件
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)

# =============================================================================
# --- Flask应用 ---
# =============================================================================

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

@app.route('/')
def index():
    """渲染主页"""
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze_pdfs():
    """处理PDF分析请求"""
    
    # 首次请求时加载库
    import_heavy_libraries()
    
    # 获取表单数据（注意字段名要匹配前端）
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('apiKey')  # 注意：前端发送的是 apiKey
    mode = request.form.get('mode', '泛读模式')
    language = request.form.get('language', '中文')
    custom_prompt = request.form.get('customPrompt', CUSTOM_TEMPLATE)
    
    # 调试输出
    print("\n" + "="*50)
    print("收到分析请求：")
    print(f"  文件数量: {len(pdf_files)}")
    print(f"  分析模式: {mode}")
    print(f"  输出语言: {language}")
    if api_key:
        print(f"  API密钥: {api_key[:8]}...{api_key[-4:]}")
    else:
        print("  ⚠️ API密钥为空！")
    print("="*50 + "\n")
    
    # 验证输入
    if not api_key:
        return jsonify({"error": "API密钥不能为空"}), 400
    
    if not api_key.startswith("sk-"):
        return jsonify({"error": "API密钥格式不正确（应以sk-开头）"}), 400
    
    if not pdf_files or len(pdf_files) == 0:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400
    
    # 限制文件数量（避免超时）
    if len(pdf_files) > 5:
        return jsonify({"error": "为避免超时，每次最多处理5个文件"}), 400
    
    # 处理文件
    all_results = []
    success_count = 0
    
    # 使用线程池并行处理（最多3个并发）
    max_workers = min(3, len(pdf_files))
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for pdf_file in pdf_files:
            future = executor.submit(process_single_pdf, pdf_file, api_key, mode, language)
            futures.append(future)
        
        # 收集结果
        for future in as_completed(futures):
            try:
                result = future.result(timeout=90)  # 单文件最多90秒
                if result and '错误' not in result:
                    all_results.append(result)
                    success_count += 1
            except Exception as e:
                print(f"  ❌ 处理异常: {e}")
    
    # 检查结果
    if not all_results:
        return jsonify({"error": "所有文件都处理失败，请检查API密钥或PDF内容"}), 500
    
    print(f"\n✓ 成功处理 {success_count}/{len(pdf_files)} 个文件")
    
    # 生成Excel报告
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            df = pd.DataFrame(all_results)
            
            # 调整列顺序
            if '文件名' in df.columns:
                cols = df.columns.tolist()
                cols.remove('文件名')
                cols.insert(0, '文件名')
                if '分析时间' in df.columns:
                    cols.remove('分析时间')
                    cols.insert(1, '分析时间')
                df = df[cols]
            
            # 保存Excel
            df.to_excel(tmp.name, index=False, engine='openpyxl')
            beautify_excel_professional(tmp.name)
            
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"EggScan_{mode}_{timestamp}.xlsx"
            
            # 发送文件
            response = send_file(
                tmp.name,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # 清理临时文件
            @response.call_on_close
            def cleanup():
                if os.path.exists(tmp.name):
                    os.unlink(tmp.name)
            
            print(f"✓ 报告已生成: {filename}")
            return response
            
    except Exception as e:
        print(f"❌ 生成报告失败: {e}")
        return jsonify({"error": f"生成报告失败: {str(e)}"}), 500

@app.route('/test', methods=['GET'])
def test():
    """测试接口"""
    return jsonify({
        "status": "ok",
        "message": "EggScan服务正在运行",
        "version": "3.0"
    })

# 错误处理
@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({"error": "文件太大，请确保总大小不超过100MB"}), 413

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "服务器内部错误，请稍后重试"}), 500

if __name__ == '__main__':
    # 本地测试
    app.run(host='0.0.0.0', port=5000, debug=True)
