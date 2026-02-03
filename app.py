from flask import Flask, request, jsonify, render_template, send_from_directory, after_this_request
import sys
import os
import uuid
import shutil
from werkzeug.utils import secure_filename

# 将 scripts/src 加入路径以便导入
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts', 'src'))

from process_data import process_hospital_data
from batch_merge import merge_excel_files

app = Flask(__name__)

# 配置上传和下载目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp_uploads')
DOWNLOAD_FOLDER = os.path.join(BASE_DIR, 'temp_downloads')

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制最大上传 16MB

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')

@app.route('/api/download/<path:filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/api/process_data', methods=['POST'])
def api_process_data():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file:
        # 为每个请求创建唯一 ID，防止冲突
        task_id = str(uuid.uuid4())
        # 使用原始文件名作为基础，但要确保安全 (这里我们保留中文名，稍微简化处理)
        # 如果需要极度安全，可以使用 secure_filename，但会丢失中文
        original_filename = file.filename
        
        # 构建临时输入输出路径
        task_upload_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
        os.makedirs(task_upload_dir, exist_ok=True)
        
        src_path = os.path.join(task_upload_dir, original_filename)
        file.save(src_path)
        
        # 输出文件路径
        output_filename = f"{os.path.splitext(original_filename)[0]}_processed.xlsx"
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

        try:
            success = process_hospital_data(src_file=src_path, output_file=output_path)
            
            # 清理上传的临时文件
            shutil.rmtree(task_upload_dir, ignore_errors=True)
            
            if success:
                return jsonify({
                    "message": "处理成功",
                    "download_url": f"/api/download/{output_filename}",
                    "filename": output_filename
                }), 200
            else:
                return jsonify({"error": "数据处理失败，请检查文件格式。"}), 500
        except Exception as e:
            return jsonify({"error": str(e)}), 500

@app.route('/api/merge_files', methods=['POST'])
def api_merge_files():
    if 'files' not in request.files:
        return jsonify({"error": "No files part"}), 400
    
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return jsonify({"error": "No selected files"}), 400

    # 1. 创建本次任务的独立上传目录
    task_id = str(uuid.uuid4())
    task_input_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
    os.makedirs(task_input_dir, exist_ok=True)

    # 2. 保存所有上传的文件
    saved_count = 0
    for file in files:
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            file.save(os.path.join(task_input_dir, file.filename))
            saved_count += 1
    
    if saved_count == 0:
        shutil.rmtree(task_input_dir)
        return jsonify({"error": "未找到有效的 Excel 文件"}), 400

    # 3. 确定输出文件名
    custom_name = request.form.get('output_filename')
    output_filename = custom_name if custom_name else f"合并汇总_{task_id[:8]}.xlsx"
    if not output_filename.endswith(('.xlsx', '.xls')):
        output_filename += '.xlsx'
        
    # 我们不直接传绝对路径给 merge_excel_files 的 output_path，因为它内部会拼路径
    # 但由于我们的脚本设计是传入 output_dir，所以没问题
    
    try:
        # 调用核心脚本
        # 注意：这里我们利用 output_filename 参数来控制输出文件名
        # 脚本内部会把 output_dir 和 output_filename 拼起来
        result_path = merge_excel_files(
            input_dir=task_input_dir, 
            output_dir=app.config['DOWNLOAD_FOLDER'], 
            output_filename=output_filename
        )
        
        # 清理输入文件
        shutil.rmtree(task_input_dir, ignore_errors=True)

        if result_path and os.path.exists(result_path):
            # 获取实际生成的文件名 (因为脚本可能会自动改名，虽然我们传了 output_filename)
            final_filename = os.path.basename(result_path)
            return jsonify({
                "message": f"成功合并 {saved_count} 个文件",
                "download_url": f"/api/download/{final_filename}",
                "filename": final_filename
            }), 200
        else:
            return jsonify({"error": "合并失败，未能生成结果文件。"}), 500
            
    except Exception as e:
        shutil.rmtree(task_input_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # host='0.0.0.0' 使其可被外部访问 (Docker 需要)
    app.run(debug=False, host='0.0.0.0', port=5000)