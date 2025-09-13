🔬 EggScan 文献智能分析系统 v2.0 (云端版)
这是一个基于Flask的Web应用，旨在将EggScan的强大文献分析能力部署到云端，方便用户通过浏览器访问和使用。

✨ 功能特性
三种分析模式:

经典五段式: 快速提取论文的核心框架（背景、方法、设计、结果、讨论）。

精读模式: 提供深入的、批判性的学术分析，挖掘研究灵感。

自定义模式: 用户可以根据自己的特定需求，灵活定制分析的维度。

Web界面: 提供一个现代化、用户友好的界面，支持多文件上传、模式选择和API密钥管理。

专业级Excel报告: 自动生成经过精心美化的Excel分析报告，格式清晰，阅读体验佳。

API密钥本地保存: 方便用户在自己的浏览器中记住API密钥，无需重复输入。

🚀 本地运行指南
在部署到云服务器之前，请先在本地环境中成功运行。

1. 准备环境
确保您的电脑已安装 Python 3.8 或更高版本。

建议使用虚拟环境以隔离项目依赖。

# antd: 创建一个新的虚拟环境 (例如使用venv)
python -m venv venv

# antd: 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

2. 安装依赖
将本项目提供的 requirements.txt 文件放在项目根目录，然后运行：

pip install -r requirements.txt

3. 运行应用
在项目根目录下，直接运行 app.py 文件：

python app.py

终端会显示类似以下信息：

 * Running on [http://0.0.0.0:5000/](http://0.0.0.0:5000/)
Press CTRL+C to quit

现在，打开您的浏览器，访问 http://127.0.0.1:5000 或 http://localhost:5000，您应该就能看到EggScan的Web界面了。

☁️ 云端部署建议
将Flask应用部署到生产环境，不推荐直接使用 app.run()。您应该使用一个专业的 WSGI (Web Server Gateway Interface) 服务器，例如 Gunicorn 或 Waitress。

使用 Gunicorn (适用于Linux服务器)
安装 Gunicorn:

pip install gunicorn

运行应用:

# antd: -w 4 表示启动4个工作进程，-b 0.0.0.0:8000 表示绑定到8000端口
# antd: app:app 指的是运行 app.py 文件中的 app 实例
gunicorn -w 4 -b 0.0.0.0:8000 app:app

使用 Waitress (适用于Windows服务器)
安装 Waitress:

pip install waitress

运行应用:

waitress-serve --host 0.0.0.0 --port 8000 app:app

部署平台
您可以将此应用部署到任何支持Python的云平台，例如：

Heroku: 对小型应用非常友好，部署流程简单。

Vercel: 虽然以部署前端框架闻名，但也支持Python Serverless Functions。

传统云服务器 (VPS): 例如阿里云、腾讯云、AWS EC2等。您需要在服务器上配置好Python环境、安装依赖，并使用Gunicorn/Waitress配合Nginx等反向代理来运行此应用。

祝您部署顺利！
