from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation
import os
import tempfile
import traceback

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process-ppt', methods=['POST'])
def process_ppt():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if not file.filename.endswith(('.ppt', '.pptx')):
        return jsonify({'error': '请上传.ppt或.pptx格式的文件'}), 400

    try:
        # 创建临时文件来保存上传的PPT
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, 'input.pptx')
        output_path = os.path.join(temp_dir, 'output.pptx')
        
        # 保存上传的文件
        file.save(input_path)
        
        try:
            # 打开PPT文件
            prs = Presentation(input_path)
        except Exception as e:
            return jsonify({'error': f'无法打开PPT文件: {str(e)}'}), 400
        
        try:
            # 在每一页添加文字
            for slide in prs.slides:
                # 添加文本框
                left = prs.slide_width * 0.1  # 左边距为幻灯片宽度的10%
                top = prs.slide_height * 0.1   # 上边距为幻灯片高度的10%
                width = prs.slide_width * 0.8  # 宽度为幻灯片宽度的80%
                height = prs.slide_height * 0.1 # 高度为幻灯片高度的10%
                
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                
                # 添加文字
                p = tf.add_paragraph()
                p.text = "上海交大"
                p.font.size = 1000000  # 10pt
                p.font.name = '微软雅黑'
            
            # 保存修改后的PPT
            prs.save(output_path)
        except Exception as e:
            return jsonify({'error': f'处理PPT时出错: {str(e)}'}), 500
        
        try:
            # 发送文件
            return send_file(
                output_path,
                as_attachment=True,
                download_name='processed_' + file.filename,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        except Exception as e:
            return jsonify({'error': f'发送文件时出错: {str(e)}'}), 500
            
    except Exception as e:
        # 捕获所有其他异常
        error_msg = f'发生错误: {str(e)}\n{traceback.format_exc()}'
        print(error_msg)  # 在服务器端打印详细错误信息
        return jsonify({'error': '处理文件时发生错误，请重试'}), 500
        
    finally:
        # 清理临时文件
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
            os.rmdir(temp_dir)
        except Exception as e:
            print(f'清理临时文件时出错: {str(e)}')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 