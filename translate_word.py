# -*- coding: utf-8 -*-
"""
Word 文档翻译小工具
功能：读取桌面上的 input.docx，把每个段落的中文/英文文本翻译成中文，写入新的 output.docx
"""

# ========== 第 1 步：导入需要的工具库 ==========
import os
import time
import pandas as pd # 虽然不用处理Excel，但为了让整个项目看起来统一，可以留着
from docx import Document # python-docx 库，用于读写 Word 文档
from zhipuai import ZhipuAI # 智谱 AI 官方 SDK

# ========== 第 2 步：填写你自己的 API Key ==========
# 注意：千万不要把这个真实的 Key 上传到 GitHub！
# 在本地测试时，把下面的字符串换成你新申请的 API Key
API_KEY = "339a7e6674224771808b087862db1e75.B47GudQivNc61iNy" 

# ========== 第 3 步：准备工作（路径和文件名） ==========
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
input_file = os.path.join(desktop_path, "input.docx")
output_file = os.path.join(desktop_path, "output.docx")

# ========== 第 4 步：写一个“翻译文本”的函数 ==========
def translate_text(client, text):
    """
    把传入的文本翻译成中文。
    容错机制：如果出错，返回原文本，不让程序崩溃。
    """
    try:
        response = client.chat.completions.create(
            model="glm-4-flash",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "你是一个专业的翻译助手。"
                        "请把用户输入的英文翻译成准确、流畅的中文。"
                        "只输出翻译结果，不要包含任何解释或额外说明。"
                    ),
                },
                {
                    "role": "user",
                    "content": text,
                },
            ],
            temperature=0.1,  # 保持低随机性，保证翻译质量稳定
        )
        chinese = response.choices[0].message.content.strip()
        return chinese
    except Exception as error:
        print(f"  ⚠️ 翻译失败：{error}")
        return text # 翻译失败就返回原文，保证程序继续跑

# ========== 第 5 步：主程序 ==========
def main():
    # 5.1 检查 Key 是否替换
    if API_KEY == "你的keys":
        print("❌ 请先把代码里的 API_KEY 换成你自己的真实 Key 再运行！")
        return

    # 5.2 检查输入文件
    if not os.path.exists(input_file):
        print(f"❌ 没有找到输入文件：{input_file}")
        print("请确认桌面上有 input.docx，且里面有英文内容。")
        return

    # 5.3 读取 Word 文档
    print("✅ 成功读取 input.docx，开始翻译……")
    doc = Document(input_file)
    
    # 5.4 创建一个全新的 Word 文档用来存翻译结果
    new_doc = Document()
    
    # 5.5 创建智谱客户端
    client = ZhipuAI(api_key=API_KEY)

    # 5.6 遍历原文档的每一段
    for i, paragraph in enumerate(doc.paragraphs):
        original_text = paragraph.text.strip()
        
        # 如果这个段落是空的（比如空行），直接在写新文档里也加个空行，节约API调用
        if not original_text:
            new_doc.add_paragraph("")
            continue
        
        # 调用翻译函数
        translated_text = translate_text(client, original_text)
        
        # 把翻译结果写入新文档的一个新段落
        new_doc.add_paragraph(translated_text)
        
        # 打印进度
        print(f"  第 {i+1}/{len(doc.paragraphs)} 段：{original_text[:20]}... -> {translated_text[:20]}...")
        
        # 暂停 1 秒，防止触发API限流
        time.sleep(1)

    # 5.7 保存新文档
    new_doc.save(output_file)
    print(f"🎉 全部完成！翻译结果已保存到：{output_file}")

# ========== 第 6 步：程序入口 ==========
if __name__ == "__main__":
    main()
