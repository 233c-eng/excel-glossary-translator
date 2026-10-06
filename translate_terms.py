# -*- coding: utf-8 -*-
"""
术语翻译小工具
功能：读取桌面上的 input.xlsx，把 A 列的英文词汇翻译成中文，写入 B 列，保存为 output.xlsx
"""

# ========== 第 1 步：导入需要的工具库 ==========
import os                      # 用于拼出桌面文件夹的路径（不用手敲 C:\Users\xxx\Desktop）
import time                    # 用于在每次 API 请求之间暂停一下，防止请求太快被服务器限流
import pandas as pd            # 用于读取和写入 Excel 文件
from zhipuai import ZhipuAI    # 智谱 AI 官方 Python SDK，用于调用大模型

# ========== 第 2 步：填写你自己的 API Key ==========
# 去 https://open.bigmodel.cn 注册后，在“API Keys”页面创建一个 Key，粘贴到下面引号里
API_KEY = "你的keys"

# ========== 第 3 步：准备工作（路径和文件名） ==========
# os.path.expanduser("~") 会自动找到当前用户的主目录（Windows 一般是 C:\Users\你的用户名）
# 再和 "Desktop" 拼在一起，就得到了桌面路径，Windows / Mac 都能用
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# 输入文件：桌面上的 input.xlsx
input_file = os.path.join(desktop_path, "input.xlsx")
# 输出文件：翻译完成后也保存在桌面，叫 output.xlsx
output_file = os.path.join(desktop_path, "output.xlsx")

# ========== 第 4 步：写一个“翻译一个词”的函数 ==========
def translate_word(client, english_word):
    """
    把传入的一个英文单词/术语翻译成中文。
    关键设计：如果翻译过程中出任何错（网络中断、Key 失效、服务器限流……），
    函数不会让整个程序崩溃，而是返回字符串 "翻译失败"。
    """
    try:
        # 调用大模型的“对话补全”接口
        response = client.chat.completions.create(
            model="glm-4-flash",   # 模型名；glm-4-flash 是免费模型，翻译术语完全够用
            messages=[
                {
                    # system 消息：给大模型设定“人设”和任务规则
                    "role": "system",
                    "content": (
                        "你是一个专业的术语翻译助手。"
                        "只把用户给的英文词汇翻译成简洁、准确的中文，"
                        "不要输出任何解释、标点或多余文字。"
                    ),
                },
                {
                    # user 消息：本次真正要翻译的内容
                    "role": "user",
                    "content": f"请把下面的英文术语翻译成中文：{english_word}",
                },
            ],
            temperature=0.1,       # “温度”调低，让大模型的输出更稳定，不随意发挥
        )
        # 从返回结果中取出中文翻译
        # response.choices[0] 是第一条回答，.message.content 是回答的文字内容
        chinese = response.choices[0].message.content.strip()
        return chinese

    except Exception as error:
        # 任何异常都会跳到这里：打印一条警告，然后返回固定的“翻译失败”
        print(f"  ⚠️ “{english_word}” 翻译失败：{error}")
        return "翻译失败"

# ========== 第 5 步：主程序 ==========
def main():
    # 5.2 检查 input.xlsx 是否真的存在于桌面
    if not os.path.exists(input_file):
        print(f"❌ 没有找到输入文件：{input_file}")
        print("请确认桌面上有 input.xlsx，并且 A 列放的是英文词汇。")
        return

    # 5.3 用 pandas 读取 Excel 文件
    # header=None 表示：不要把第一行当成“标题行”，所有行都按数据处理
    df = pd.read_excel(input_file, header=None)
    print(f"✅ 成功读取 {len(df)} 行数据，开始翻译……")

    # 5.4 创建一个空列表，用来逐个收集每一行翻译出的中文
    translations = []

    # 5.5 创建智谱客户端（只创建一次，可以反复用来发很多条翻译请求）
    client = ZhipuAI(api_key=API_KEY)

    # 5.6 逐行读取 A 列（pandas 里第 0 列就是 Excel 的 A 列）并翻译
    for index in range(len(df)):
        english_word = df.iloc[index, 0]   # iloc[行号, 列号]：取第 index 行、第 0 列

        # 如果这一行是空白的（NaN），直接写空字符串，不去浪费一次 API 调用
        if pd.isna(english_word):
            translations.append("")
            continue  # continue 表示跳过本次循环，直接进入下一行

        # 把取到的内容转成字符串，并去掉首尾可能存在的空格或换行
        english_word = str(english_word).strip()

        # 调用上面写好的翻译函数，得到中文结果
        chinese = translate_word(client, english_word)
        translations.append(chinese)

        # 打印当前进度，让你知道程序正在正常运行、没有卡死
        print(f"  第 {index + 1}/{len(df)} 行：{english_word} → {chinese}")

        # 每次请求后暂停 1 秒，避免请求速度太快触发服务器的频率限制
        time.sleep(1)

    # 5.7 把收集好的翻译结果写进 B 列（pandas 里第 1 列就是 Excel 的 B 列）
    df[1] = translations

    # 5.8 把处理好的表格保存为 output.xlsx
    # index=False 表示不把“行号”写进文件；header=False 表示不额外写一行标题
    df.to_excel(output_file, index=False, header=False)
    print(f"🎉 全部完成！翻译结果已保存到：{output_file}")

# ========== 第 6 步：程序入口 ==========
# 意思是：只有“直接运行这个文件”时才会执行 main()；
# 如果哪天别的程序导入这个文件，不会自动跑翻译，这是一个 Python 的常用惯例
if __name__ == "__main__":
    main()