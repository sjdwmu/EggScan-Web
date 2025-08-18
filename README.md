# EggScan-Web
EggScan 文献分析工具
这是一个基于Python的Web应用，旨在帮助用户快速批量分析PDF文献，并提取其中的结构化信息，最终生成格式美观的Excel报告。
✨ 功能特点
📚 PDF文本提取: 支持从PDF文件中提取文本，对于难以直接提取的PDF，会自动切换到OCR（光学字符识别）模式。
🧠 LLM 智能分析: 调用大语言模型（LLM），从论文中智能提取“研究背景”、“研究方法”、“实验设计”、“结果分析”和“讨论”等关键信息。
📊 Excel报告生成: 将分析结果以结构化的表格形式输出到Excel文件中，并进行自动化格式美化。
🌐 跨平台: 作为Web应用，用户无需安装任何本地依赖，只需通过浏览器即可访问和使用。
🛠️ 技术栈
后端: Python (Flask)
前端: HTML, Tailwind CSS, JavaScript
核心库:
Flask: 轻量级Web框架
PyMuPDF: 用于PDF文本提取
pdf2image: 用于PDF转图像（支持OCR）
pytesseract: OCR核心库
pandas: 数据处理和Excel导出
openpyxl: Excel格式美化
requests: 调用LLM API
🚀 本地部署和使用
本应用可以轻松部署到任何支持Python的环境中。
1. 安装依赖
首先，请确保你的系统中安装了Python 3.8+，然后安装项目依赖。
pip install -r requirements.txt
2. 运行应用
在项目根目录（与 app.py 同一级）下，运行以下命令启动Web服务。
python app.py
3. 访问应用
程序启动后，你将在终端看到一个本地地址，如 http://127.0.0.1:5000。
在浏览器中打开这个地址，即可访问应用。
4. 远程访问
如果你希望在局域网内让其他设备访问，请将地址中的 127.0.0.1 替换为你的电脑在局域网中的IP地址。

💡 使用说明
打开浏览器访问应用地址。
在主页上传一个或多个PDF文件。
输入你的 DeepSeek API 密钥。
点击“开始分析并下载Excel”，程序将自动处理并返回一个包含分析结果的 EggScan_Result.xlsx 文件。
注意:
请确保你的API密钥有效。
大文件或大量文件处理可能需要较长时间。
