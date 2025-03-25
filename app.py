from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation
import os
import tempfile
import traceback
import requests
import json

app = Flask(__name__)

# Dify API配置
DIFY_API_URL = "http://10.119.14.166/v1/chat-messages"
DIFY_API_KEY = "Bearer app-ujLJoBR6bFWdo33nqmgOoEdM"

def get_dify_response(slide_text):
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": DIFY_API_KEY
    }
    
    data = {
        "query": f"请为以下PPT内容生成一段简短的备注说明：{slide_text}",
        "response_mode": "blocking",
        "conversation_id": "",
        "user": "ppt_user",
        "inputs": {},
        "query_parameters": {}
    }
    
    try:
        print(f"正在调用Dify API，内容：{slide_text[:100]}...")  # 打印前100个字符
        response = requests.post(DIFY_API_URL, headers=headers, json=data)
        print(f"API响应状态码：{response.status_code}")
        print(f"API响应内容：{response.text[:200]}...")  # 打印前200个字符
        
        response.raise_for_status()
        result = response.json()
        
        if 'answer' in result:
            return result['answer']
        elif 'choices' in result and len(result['choices']) > 0:
            return result['choices'][0].get('message', {}).get('content', '')
        else:
            print(f"API返回结果格式：{json.dumps(result, ensure_ascii=False)}")
            return "无法解析API返回结果"
            
    except requests.exceptions.RequestException as e:
        print(f"API请求错误: {str(e)}")
        return f"API请求错误: {str(e)}"
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {str(e)}")
        print(f"原始响应内容: {response.text}")
        return "API返回格式错误"
    except Exception as e:
        print(f"其他错误: {str(e)}")
        print(f"错误堆栈: {traceback.format_exc()}")
        return f"发生错误: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process-ppt', methods=['POST'])
def process_ppt():
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if not (file.filename.endswith('.ppt') or file.filename.endswith('.pptx')):
        return jsonify({'error': '请上传PPT文件（.ppt或.pptx格式）'}), 400

    try:
        # 创建临时文件保存上传的PPT
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_input:
            file.save(temp_input.name)
            input_path = temp_input.name

        # 创建临时文件保存处理后的PPT
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_output:
            output_path = temp_output.name

        # 打开PPT文件
        prs = Presentation(input_path)

        # 为每一页添加备注
        for i, slide in enumerate(prs.slides, 1):
            print(f"\n处理第 {i} 页...")
            
            # 获取幻灯片文本内容
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text += shape.text + "\n"
            
            print(f"提取的文本内容：{slide_text[:100]}...")  # 打印前100个字符
            
            # 获取Dify生成的备注内容
            notes_content = get_dify_response(slide_text)
            print(f"生成的备注内容：{notes_content}")
            
            if not notes_content:
                notes_content = f"第{i}页备注生成失败"
            
            # 获取或创建备注页
            notes_slide = slide.notes_slide
            notes_text_frame = notes_slide.notes_text_frame
            
            # 添加备注
            notes_text_frame.text = notes_content

        # 保存处理后的PPT
        prs.save(output_path)
        print("\nPPT处理完成")

        # 删除临时输入文件
        os.unlink(input_path)

        # 发送处理后的文件
        return send_file(
            output_path,
            as_attachment=True,
            download_name='processed_ppt.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        print(f"处理PPT时出错: {str(e)}")
        print(f"错误堆栈: {traceback.format_exc()}")
        # 清理临时文件
        if 'input_path' in locals():
            os.unlink(input_path)
        if 'output_path' in locals():
            os.unlink(output_path)
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 