import os
import uuid
import shutil
from flask import request, jsonify, render_template, send_from_directory
from core.processor import process_hospital_data
from core.merger import merge_excel_files

def register_routes(app):
    # 配置从 app 对象获取（假设在 app.py 中定义）
    UPLOAD_FOLDER = app.config['UPLOAD_FOLDER']
    DOWNLOAD_FOLDER = app.config['DOWNLOAD_FOLDER']

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/dashboard')
    def dashboard():
        return render_template('dashboard.html')

    @app.route('/api/download/<path:filename>')
    def download_file(filename):
        return send_from_directory(DOWNLOAD_FOLDER, filename, as_attachment=True)

    @app.route('/api/process_data', methods=['POST'])
    def api_process_data():
        if 'file' not in request.files:
            return jsonify({"error": "No file part"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No selected file"}), 400

        task_id = str(uuid.uuid4())
        task_upload_dir = os.path.join(UPLOAD_FOLDER, task_id)
        os.makedirs(task_upload_dir, exist_ok=True)
        
        src_path = os.path.join(task_upload_dir, file.filename)
        file.save(src_path)
        
        output_filename = f"{os.path.splitext(file.filename)[0]}_processed.xlsx"
        output_path = os.path.join(DOWNLOAD_FOLDER, output_filename)

        try:
            success = process_hospital_data(src_file=src_path, output_file=output_path)
            shutil.rmtree(task_upload_dir, ignore_errors=True)
            
            if success:
                return jsonify({
                    "message": "处理成功",
                    "download_url": f"/api/download/{output_filename}",
                    "filename": output_filename
                }), 200
            else:
                return jsonify({"error": "数据处理失败"}), 500
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route('/api/merge_files', methods=['POST'])
    def api_merge_files():
        if 'files' not in request.files:
            return jsonify({"error": "No files part"}), 400
        
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({"error": "No selected files"}), 400

        task_id = str(uuid.uuid4())
        task_input_dir = os.path.join(UPLOAD_FOLDER, task_id)
        os.makedirs(task_input_dir, exist_ok=True)

        for file in files:
            if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                file.save(os.path.join(task_input_dir, file.filename))
        
        custom_name = request.form.get('output_filename')
        output_filename = custom_name if custom_name else f"合并汇总_{task_id[:8]}.xlsx"
        if not output_filename.endswith(('.xlsx', '.xls')):
            output_filename += '.xlsx'
            
        try:
            result_path = merge_excel_files(
                input_dir=task_input_dir, 
                output_dir=DOWNLOAD_FOLDER, 
                output_filename=output_filename
            )
            shutil.rmtree(task_input_dir, ignore_errors=True)

            if result_path and os.path.exists(result_path):
                final_filename = os.path.basename(result_path)
                return jsonify({
                    "message": "成功合并文件",
                    "download_url": f"/api/download/{final_filename}",
                    "filename": final_filename
                }), 200
            else:
                return jsonify({"error": "合并失败"}), 500
        except Exception as e:
            shutil.rmtree(task_input_dir, ignore_errors=True)
            return jsonify({"error": str(e)}), 500
