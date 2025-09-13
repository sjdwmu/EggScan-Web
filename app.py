# -*- coding: utf-8 -*-
# =============================================================================
# --- EggScan 云端分析应用 (异步处理版 v3.0) ---
# =============================================================================

import os
import re
import json
import uuid
import tempfile
import threading
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from collections import defaultdict

# 延迟导入
fitz, requests, pd, load_workbook, Font, Alignment, PatternFill, Border, Side = (None,) * 9

def import_heavy_libraries():
    """延迟导入重量级库"""
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
            raise

# Flask配置
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
LLM_URL = "https://api.deepseek.com/v1/chat/completions"

# 任务存储（生产环境建议使用Redis）
TASKS = {}
TASK_RESULTS = {}

# 分析框架定义
SKIMMING_FIELDS = ["研究问题", "核心论点", "研究方法", "关键结论", "相关性评估"]
INTENSIVE_FIELDS = ["研究背景与缺口", "研究设计与方法", "主要结果与数据", "创新点与贡献", "局限性与批判", "可借鉴与启发"]
CUSTOM_TEMPLATE = """
请从以下角度分析这篇文献：
【研究主题】：文章的核心研究问题是什么？
【理论框架】：使用了什么理论基础或概念框架？
【方法创新】：在研究方法上有什么创新或特色？
【数据质量】：数据来源、样本量、统计分析的可靠性如何？
【关键发现】：最重要的3个研究发现是什么？
【实践意义】：对临床实践或政策制定有什么指导意义？
请用【字段名】：内容 的格式清晰输出。
"""

# =============================================================================
# --- 核心函数（与v3.0保持一致）---
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
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[1].height = 30
        
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
        
        for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
            ws.row_dimensions[row_num].height = 80
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = thin_border
                cell.font = data_font
                if row_num % 2 == 0:
                    cell.fill = PatternFill(fill_type="solid", fgColor="F2F2F2")
        
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        wb.save(filepath)
        
    except Exception as e:
        print(f"⚠️ Excel美化失败: {e}")

def call_llm_for_mode(pdf_text, api_key, mode, language):
    """根据模式调用LLM"""
    lang_instruction = "Please output in English" if language == "English" else "请用中文输出"
    
    if mode == '泛读模式':
        prompt = f"""
你是一位专业的文献筛选专家，请对这篇论文进行快速泛读分析（5-10分钟内完成）。
目标：快速判断文献的相关性和核心价值。

{lang_instruction}

请严格按照以下格式提取关键信息：
【研究问题】：这篇文章具体想回答什么问题？
【核心论点】：作者最核心的观点是什么？（一句话总结）
【研究方法】：这是什么类型的研究？（如：RCT/Meta分析/队列研究等）
【关键结论】：最重要的研究结论是什么？
【相关性评估】：评估其研究价值（高相关/中相关/低相关）

---
论文内容：
{pdf_text[:30000]}
"""
        fields = SKIMMING_FIELDS
        
    elif mode == '精读模式':
        prompt = f"""
你是一位资深的学术研究专家，请对这篇论文进行全面深入的精读分析。

{lang_instruction}

请按照以下六个维度进行详细分析：
【研究背景与缺口】：详细阐述研究背景和空白
【研究设计与方法】：包括样本量、分组、统计方法等
【主要结果与数据】：关键数据和图表引用
【创新点与贡献】：理论/方法/实践创新
【局限性与批判】：作者承认的+你发现的问题
【可借鉴与启发】：可直接借鉴的方法和研究思路

---
论文内容：
{pdf_text[:40000]}
"""
        fields = INTENSIVE_FIELDS
    else:
        return None, None
    
    HEADERS = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "max_tokens": 4096
    }
    
    try:
        response = requests.post(LLM_URL, headers=HEADERS, json=payload, timeout=180)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"], fields
    except Exception as e:
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
            result_dict[field] = content if content and len(content) > 5 else f"解析失败-{field}"
        else:
            result_dict[field] = f"未提取到-{field}"
    
    return result_dict

# =============================================================================
# --- 异步任务处理 ---
# =============================================================================

def process_pdfs_async(task_id, pdf_files_data, api_key, mode, language, custom_prompt):
    """异步处理PDF文件"""
    import_heavy_libraries()
    
    TASKS[task_id]['status'] = 'processing'
    TASKS[task_id]['total'] = len(pdf_files_data)
    TASKS[task_id]['processed'] = 0
    
    all_results = []
    
    for idx, (filename, file_content) in enumerate(pdf_files_data):
        try:
            # 保存临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(file_content)
                tmp_path = tmp.name
            
            # 提取文本
            text = smart_extract_text(tmp_path)
            os.unlink(tmp_path)  # 删除临时文件
            
            if len(text.strip()) < 500:
                TASKS[task_id]['messages'].append(f"⚠️ {filename}: 文本内容不足，跳过")
                continue
            
            # 调用LLM
            if mode == '自定义模式':
                # 自定义模式处理
                full_prompt = f"{custom_prompt}\n\n论文内容：\n{text[:40000]}"
                llm_output, _ = call_llm_for_mode(text, api_key, '泛读模式', 'Chinese')  # 临时使用
                fields = re.findall(r'【([^】]+)】', custom_prompt)
                result = parse_llm_output(llm_output, fields)
            else:
                llm_output, fields = call_llm_for_mode(text, api_key, mode, language)
                result = parse_llm_output(llm_output, fields)
            
            result['文件名'] = filename
            result['分析时间'] = datetime.now().strftime("%Y-%m-%d %H:%M")
            all_results.append(result)
            
            # 更新进度
            TASKS[task_id]['processed'] = idx + 1
            TASKS[task_id]['messages'].append(f"✓ {filename} 处理完成")
            
        except Exception as e:
            TASKS[task_id]['messages'].append(f"❌ {filename}: {str(e)}")
    
    # 生成Excel
    if all_results:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                df = pd.DataFrame(all_results)
                
                # 调整列顺序
                priority_cols = ['文件名', '分析时间']
                other_cols = [col for col in df.columns if col not in priority_cols]
                df = df[[col for col in priority_cols if col in df.columns] + other_cols]
                
                df.to_excel(tmp_excel.name, index=False, engine='openpyxl')
                beautify_excel_professional(tmp_excel.name)
                
                # 保存结果
                with open(tmp_excel.name, 'rb') as f:
                    TASK_RESULTS[task_id] = {
                        'filename': f'EggScan_{mode}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                        'data': f.read()
                    }
                
                os.unlink(tmp_excel.name)
                
            TASKS[task_id]['status'] = 'completed'
            TASKS[task_id]['messages'].append("🎉 分析完成！报告已生成")
        except Exception as e:
            TASKS[task_id]['status'] = 'failed'
            TASKS[task_id]['messages'].append(f"❌ 生成报告失败: {str(e)}")
    else:
        TASKS[task_id]['status'] = 'failed'
        TASKS[task_id]['messages'].append("❌ 没有成功处理任何文件")

# =============================================================================
# --- Flask路由 ---
# =============================================================================

@app.route('/')
def index():
    """渲染主页"""
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def start_analysis():
    """启动异步分析任务"""
    pdf_files = request.files.getlist('pdfs')
    api_key = request.form.get('apiKey')
    mode = request.form.get('mode')
    language = request.form.get('language', '中文')
    custom_prompt = request.form.get('customPrompt', CUSTOM_TEMPLATE)
    
    # 验证输入
    if not pdf_files:
        return jsonify({"error": "请至少上传一个PDF文件"}), 400
    if not api_key or not api_key.startswith("sk-"):
        return jsonify({"error": "API密钥格式不正确"}), 400
    
    # 创建任务ID
    task_id = str(uuid.uuid4())
    
    # 初始化任务状态
    TASKS[task_id] = {
        'status': 'pending',
        'total': 0,
        'processed': 0,
        'messages': [],
        'created_at': datetime.now().isoformat()
    }
    
    # 读取所有PDF文件内容
    pdf_files_data = []
    for pdf_file in pdf_files:
        pdf_files_data.append((pdf_file.filename, pdf_file.read()))
    
    # 启动异步处理线程
    thread = threading.Thread(
        target=process_pdfs_async,
        args=(task_id, pdf_files_data, api_key, mode, language, custom_prompt)
    )
    thread.daemon = True
    thread.start()
    
    return jsonify({
        "task_id": task_id,
        "message": "任务已创建，正在处理中..."
    })

@app.route('/status/<task_id>')
def get_status(task_id):
    """获取任务状态"""
    if task_id not in TASKS:
        return jsonify({"error": "任务不存在"}), 404
    
    task = TASKS[task_id]
    return jsonify({
        "status": task['status'],
        "total": task['total'],
        "processed": task['processed'],
        "messages": task['messages'][-10:],  # 只返回最近10条消息
        "progress": (task['processed'] / task['total'] * 100) if task['total'] > 0 else 0
    })

@app.route('/download/<task_id>')
def download_result(task_id):
    """下载分析结果"""
    if task_id not in TASK_RESULTS:
        return jsonify({"error": "结果不存在或任务未完成"}), 404
    
    result = TASK_RESULTS[task_id]
    
    # 创建响应
    from io import BytesIO
    return send_file(
        BytesIO(result['data']),
        as_attachment=True,
        download_name=result['filename'],
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/cleanup', methods=['POST'])
def cleanup_old_tasks():
    """清理旧任务（可选）"""
    # 清理超过1小时的任务
    from datetime import timedelta
    cutoff_time = datetime.now() - timedelta(hours=1)
    
    tasks_to_remove = []
    for task_id, task in TASKS.items():
        if datetime.fromisoformat(task['created_at']) < cutoff_time:
            tasks_to_remove.append(task_id)
    
    for task_id in tasks_to_remove:
        TASKS.pop(task_id, None)
        TASK_RESULTS.pop(task_id, None)
    
    return jsonify({"cleaned": len(tasks_to_remove)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

