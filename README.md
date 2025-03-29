# PPT智能助手

一个基于AI的PPT智能辅助工具，帮助用户快速生成PPT备注、演讲稿和进行内容问答。

## 主要功能

1. **智能备注生成**
   - 自动分析PPT每页内容
   - 生成专业的备注说明
   - 支持下载带备注的PPT文件

2. **演讲稿生成**
   - 基于PPT内容自动生成完整演讲稿
   - 包含开场白和结束语
   - 使用仿宋字体排版
   - 适合领导演讲使用

3. **备注合集导出**
   - 将所有PPT备注整理成Word文档
   - 按页码组织内容
   - 使用仿宋字体排版
   - 支持同时下载带备注的PPT和备注合集

4. **智能问答**
   - 基于PPT内容回答问题
   - 提供准确、简洁的答案
   - 支持上下文理解

## 环境要求

- Python 3.8+
- 依赖包：见requirements.txt

## 安装步骤

1. 克隆项目到本地
```bash
git clone [项目地址]
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 配置API密钥
在app.py中配置以下变量：
```python
DIFY_API_KEY = "你的API密钥"
CHAT_API_KEY = "你的API密钥"
```

4. 运行应用
```bash
python app.py
```

## 使用说明

1. 打开浏览器访问：http://localhost:5000

2. 上传PPT文件（支持.ppt和.pptx格式）

3. 选择需要的功能：
   - 点击"生成备注"：为PPT添加智能备注
   - 点击"生成演讲稿"：生成完整的演讲稿Word文档
   - 点击"生成备注合集"：导出PPT备注的Word文档
   - 在问答框中输入问题：获取基于PPT内容的答案

## 注意事项

1. 请确保上传的PPT文件格式正确
2. 生成的文档默认使用仿宋字体
3. 需要稳定的网络连接以访问AI服务
4. 处理大型PPT文件可能需要较长时间

## 技术栈

- 后端：Flask
- 文档处理：python-pptx, python-docx
- AI服务：Dify API
- 前端：HTML/CSS/JavaScript
