# -- 文件名: Dockerfile (Python/Flask 版本) --
# -- 功能: 为你的 EggScan-Web Python 应用打包镜像 --

# 1. 选择一个官方的 Python 运行时作为基础镜像
# 我们选择一个轻量的 slim 版本
FROM python:3.9-slim

# 2. 设置工作目录
# 容器内所有操作都会在这个目录下进行
WORKDIR /app

# 3. 复制依赖文件并安装依赖
# 先只复制这一个文件，可以利用 Docker 的缓存机制，如果依赖不变，下次构建会更快
COPY requirements.txt .
RUN pip install -i https://mirrors.aliyun.com/pypi/simple/ --no-cache-dir -r requirements.txt
# 4. 复制项目所有文件
# 把当前目录 (.) 下的所有文件复制到容器的 /app 目录 (.)
COPY . .

# 5. 声明应用运行的端口
# 根据 Flask/Gunicorn 的常用实践，我们使用 5000 端口
EXPOSE 5000

# 6. 定义容器启动时执行的命令
# 我们使用 Gunicorn 来启动你的应用，它是一个专业的 WSGI 服务器，比 app.py 自带的开发服务器更稳定高效
# app:app 的意思是：运行 app.py 文件中的 app 这个 Flask 实例
CMD ["gunicorn", "--workers", "4", "--bind", "0.0.0.0:5000", "--timeout", "300", "app:app"]
