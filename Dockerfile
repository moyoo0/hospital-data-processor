# 使用轻量级 Python 3.9 镜像
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量 (防止 Python 生成 .pyc 文件，并让日志直接输出)
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# 复制依赖文件并安装
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制项目代码
COPY . .

# 创建必要的临时目录 (如果 app.py 没有自动创建的话)
RUN mkdir -p temp_uploads temp_downloads

# 暴露端口
EXPOSE 5000

# 启动命令
CMD ["python", "app.py"]
