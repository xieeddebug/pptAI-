from flask import Flask, request, send_file, render_template, jsonify, make_response, Response
from pptx import Presentation
import os
import tempfile
import traceback
import requests
import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import base64

app = Flask(__name__)

# Dify API配置
DIFY_API_URL = "http://10.119.14.166/v1/chat-messages"
DIFY_API_KEY = "Bearer app-ujLJoBR6bFWdo33nqmgOoEdM"

# PPT对话助手API配置
CHAT_API_URL = "http://10.119.14.166/v1/chat-messages"
CHAT_API_KEY = "Bearer app-aaAtWp12EleFXiKB2L7lgm2J"

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

def generate_speech(ppt_text):
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": DIFY_API_KEY
    }
    
    data = {
        "query": f"请根据以下PPT内容生成一份完整的演讲稿，要求：\n1. 语言要正式、专业\n2. 加入适当的过渡语和连接词\n3. 每页PPT的内容要连贯\n4. 适合领导演讲使用\n5. 包含开场白和结束语\n\nPPT内容：{ppt_text}",
        "response_mode": "blocking",
        "conversation_id": "",
        "user": "ppt_user",
        "inputs": {},
        "query_parameters": {}
    }
    
    try:
        print("正在生成演讲稿...")
        response = requests.post(DIFY_API_URL, headers=headers, json=data)
        print(f"API响应状态码：{response.status_code}")
        
        response.raise_for_status()
        result = response.json()
        
        if 'answer' in result:
            return result['answer']
        elif 'choices' in result and len(result['choices']) > 0:
            return result['choices'][0].get('message', {}).get('content', '')
        else:
            return "无法生成演讲稿"
            
    except Exception as e:
        print(f"生成演讲稿时出错: {str(e)}")
        return f"生成演讲稿时出错: {str(e)}"

def create_word_document(speech_text):
    doc = Document()
    
    # 设置标题
    title = doc.add_heading('演讲稿', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 添加正文
    paragraphs = speech_text.split('\n')
    for para in paragraphs:
        if para.strip():
            p = doc.add_paragraph()
            # 设置段落格式
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.first_line_indent = Pt(24)
            # 添加文本
            run = p.add_run(para.strip())
            # 设置字体
            run.font.name = '宋体'
            run.font.size = Pt(12)
    
    return doc

def get_chat_response(prompt):
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": CHAT_API_KEY
    }
    
    data = {
        "query": prompt,
        "response_mode": "blocking",
        "conversation_id": "",
        "user": "ppt_user",
        "inputs": {},
        "query_parameters": {}
    }
    
    try:
        print(f"正在调用Chat API，内容：{prompt[:100]}...")  # 打印前100个字符
        response = requests.post(CHAT_API_URL, headers=headers, json=data)
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
    
    if not file.filename.endswith(('.ppt', '.pptx')):
        return jsonify({'error': '请上传PPT文件'}), 400

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
        
        # 发送处理后的文件
        return send_file(
            output_path,
            as_attachment=True,
            download_name='processed_ppt.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        print(f"错误堆栈: {traceback.format_exc()}")
        # 清理临时文件
        if 'input_path' in locals():
            os.unlink(input_path)
        if 'output_path' in locals():
            os.unlink(output_path)
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-speech', methods=['POST'])
def generate_speech_route():
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

        # 打开PPT文件
        prs = Presentation(input_path)
        
        # 提取所有文本
        all_text = ""
        for i, slide in enumerate(prs.slides, 1):
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text += shape.text + "\n"
            all_text += f"第{i}页：\n{slide_text}\n"

        # 生成演讲稿
        speech_text = generate_speech(all_text)
        
        # 创建Word文档
        doc = create_word_document(speech_text)
        
        # 保存Word文档
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_doc:
            doc.save(temp_doc.name)
            doc_path = temp_doc.name

        # 发送Word文档
        return send_file(
            doc_path,
            as_attachment=True,
            download_name='演讲稿.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        print(f"生成演讲稿时出错: {str(e)}")
        print(f"错误堆栈: {traceback.format_exc()}")
        # 清理临时文件
        if 'input_path' in locals():
            os.unlink(input_path)
        if 'doc_path' in locals():
            os.unlink(doc_path)
        return jsonify({'error': f'生成演讲稿时出错: {str(e)}'}), 500

@app.route('/api/get-ppt-content', methods=['POST'])
def get_ppt_content():
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if not file.filename.endswith(('.ppt', '.pptx')):
        return jsonify({'error': '请上传PPT文件'}), 400

    try:
        # 保存上传的文件
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        file.save(temp_file.name)
        
        # 读取PPT内容
        prs = Presentation(temp_file.name)
        content = []
        
        for slide in prs.slides:
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text += shape.text + "\n"
            content.append(slide_text.strip())
        
        # 清理临时文件
        os.unlink(temp_file.name)
        
        return jsonify({'content': '\n\n'.join(content)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/chat', methods=['POST'])
def chat():
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

        # 打开PPT文件
        prs = Presentation(input_path)
        
        # 提取所有文本
        all_text = ""
        for i, slide in enumerate(prs.slides, 1):
            slide_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text += shape.text + "\n"
            all_text += f"第{i}页：\n{slide_text}\n"

        # 获取用户问题
        data = request.form
        question = data.get('question', '')
        
        # 构建prompt
        prompt = f"""基于以下PPT内容回答问题。如果问题与PPT内容无关，请说明无法回答。

PPT内容：
{all_text}

问题：{question}

请提供准确、简洁的回答。"""
        
        # 调用Chat API
        response = get_chat_response(prompt)
        return jsonify({'answer': response})

    except Exception as e:
        print(f"处理对话时出错: {str(e)}")
        print(f"错误堆栈: {traceback.format_exc()}")
        # 清理临时文件
        if 'input_path' in locals():
            os.unlink(input_path)
        return jsonify({'error': f'处理对话时出错: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 