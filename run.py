import os
from flask import Flask
from app.routes import register_routes

def create_app():
    app = Flask(__name__, 
                template_folder='app/templates', 
                static_folder='app/static')

    # 配置上传和下载目录
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'temp_uploads')
    app.config['DOWNLOAD_FOLDER'] = os.path.join(BASE_DIR, 'temp_downloads')
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

    # 确保目录存在
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

    # 注册路由
    register_routes(app)
    
    return app

app = create_app()

if __name__ == '__main__':
    # host='0.0.0.0' 使其可被外部访问 (Docker 需要)
    app.run(debug=False, host='0.0.0.0', port=5010)
