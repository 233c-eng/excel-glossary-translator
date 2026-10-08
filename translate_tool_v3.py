# -*- coding: utf-8 -*-
"""
AI 本地化翻译工具(glm-5.3 批量优化版）
==================================================
支持：桌面 input.xlsx(A 列英文 → B 列中文)、input.docx(逐段翻译）
可选：桌面 glossary.xlsx(A 列英文术语、B 列标准中文译法）作为术语库硬约束

批量模式说明：
  Excel 默认每 10 行合成一批，一次 API 请求翻译完，速度约为逐行模式的 8~10 倍。
  如果某一批解析失败或术语校验不通过，会自动降级为“逐行翻译”，保证每一行都有结果。
"""

import os
import re
import time
import pandas as pd
from docx import Document
from zhipuai import ZhipuAI

# ========== 1. 配置区 ==========
# ⚠️ 测试时填你自己的真实 Key，上传 GitHub 前务必换回来
API_KEY = "你的keys"
MODEL_NAME = "glm-4-flash"

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
glossary_file = os.path.join(desktop_path, "glossary.xlsx")
input_excel = os.path.join(desktop_path, "input.xlsx")
output_excel = os.path.join(desktop_path, "output.xlsx")
input_word = os.path.join(desktop_path, "input.docx")
output_word = os.path.join(desktop_path, "output.docx")

REQUEST_INTERVAL = 0.2      # 每次成功请求之间的间隔（秒）
FAIL_RETRY_TIMES = 2        # 普通翻译失败（网络抖动/限流）后的自动重试次数
RETRY_WAIT_SECONDS = 3      # 每次重试前等待的秒数
ENABLE_TERM_CHECK = True    # 是否开启“术语漏译自动校验”
MAX_TERM_RETRY = 1          # 术语校验不通过时，要求模型重译的次数

ENABLE_BATCH = True         # 是否开启批量翻译（Excel 有效，Word 始终是逐段）
BATCH_SIZE = 10             # 每批的行数；行数很多时可调到 20~30，太大容易出错

# ========== 2. 公共函数 ==========
def load_glossary(file_path):
    """读取术语库（glossary.xlsx），返回 {英文: 中文} 字典；文件不存在时返回空字典，不影响主流程。"""
    glossary = {}
    try:
        if not os.path.exists(file_path):
            print(f"⚠️ 未找到术语库：{file_path}，将退化为普通翻译。")
            return glossary
        df = pd.read_excel(file_path, header=None, dtype=str)
        for _, row in df.iterrows():
            en = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
            zh = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            # 跳过表头行（防止把 "English" 这种标题当成术语）
            if en and zh and en.lower() not in {"english", "en", "term", "source"}:
                glossary[en] = zh
        print(f"✅ 术语库加载成功（共 {len(glossary)} 条术语）")
    except Exception as e:
        print(f"⚠️ 术语库读取失败：{e}")
    return glossary


def build_system_prompt(glossary):
    """根据术语库拼出 system 提示词，让模型直接输出译文、不输出思考过程。"""
    base = (
        "你是一个专业的翻译助手。请把英文翻译成准确、流畅、地道的中文。"
        "直接输出最终的翻译结果，不要输出任何思考过程、解释或原文。"
    )
    if not glossary:
        return base

    terms = "\n".join(f"- {en} => {zh}" for en, zh in glossary.items())
    return (
        f"{base}\n\n【术语库硬约束】\n{terms}\n\n"
        "【强制规则】遇到术语库里的词，必须100%使用标准译法，不得意译或替换。"
    )


def find_missing_terms(source, translated, glossary):
    """检查原文里出现的术语，译文里是否用了标准译法；返回漏译的 [(英文, 中文), ...] 列表。"""
    missing = []
    for en, zh in glossary.items():
        # 用前后“负向断言”匹配完整单词，避免把 "AI" 匹配进 "said" 这类子串
        if re.search(rf"(?<![A-Za-z]){re.escape(en)}(?![A-Za-z])", source, flags=re.IGNORECASE):
            if zh not in translated:
                missing.append((en, zh))
    return missing


def call_model(client, prompt, text):
    """把“调一次模型”单独抽成函数，各处共用，避免代码重复。"""
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": prompt},
            {"role": "user", "content": text},
        ],
        thinking={"type": "enabled"},  # glm-5.3 深度思考开关
    )
    return response.choices[0].message.content.strip()


def translate_text(client, text, system_prompt, glossary=None):
    """
    翻译【一段】文本（单条模式），保证任何情况下都不崩溃：
    1) 调用出错 → 自动重试 FAIL_RETRY_TIMES 次；
    2) 术语漏译 → 要求模型重译 MAX_TERM_RETRY 次；
    3) 全部失败 → 返回原文，由调用方继续处理下一行。
    """
    # --- 第一阶段：正常调用 + 出错重试 ---
    result = None
    for attempt in range(FAIL_RETRY_TIMES + 1):   # +1 表示先正常试一次
        try:
            result = call_model(client, system_prompt, text)
            if not result:
                raise ValueError("模型返回了空内容")   # 空结果也视为失败，触发重试
            break                                      # 成功就跳出重试循环
        except Exception as e:
            print(f"  ⚠️ 调用出错（第 {attempt + 1} 次）：{e}")
            if attempt < FAIL_RETRY_TIMES:
                print(f"     等待 {RETRY_WAIT_SECONDS} 秒后自动重试……")
                time.sleep(RETRY_WAIT_SECONDS)

    if result is None:
        print("  ❌ 多次重试仍失败，保留原文，继续处理下一行。")
        return text

    # --- 第二阶段：术语漏译校验与重译（单条模式下的校验） ---
    if ENABLE_TERM_CHECK and glossary:
        result = term_check_loop(client, text, result, system_prompt, glossary)
    return result


def term_check_loop(client, source, result, system_prompt, glossary):
    """
    术语校验：检查译文是否漏用了术语库的标准译法，漏了就要求模型重译。
    单条翻译和批量翻译都会调用它。
    """
    if not (ENABLE_TERM_CHECK and glossary):
        return result
    for _ in range(MAX_TERM_RETRY):
        missing = find_missing_terms(source, result, glossary)
        if not missing:
            break
        force = "\n".join(f"- 「{en}」必须翻译成「{zh}」" for en, zh in missing)
        print(f"  ⚠️ 术语校验未通过（{len(missing)} 个术语缺失），正在要求模型重译……")
        retry_prompt = (
            f"{system_prompt}\n\n【再次强调】\n{force}\n"
            "请重新输出完整译文，严格使用标准译法。"
        )
        try:
            result = call_model(client, retry_prompt, source)
            if not result:
                raise ValueError("模型返回了空内容")
        except Exception as e:
            print(f"  ⚠️ 重译出错：{e}，保留当前结果。")
            break
        time.sleep(REQUEST_INTERVAL)   # 重译也要遵守请求间隔，防止限流
    return result


# ========== 3. 批量翻译核心逻辑 ==========
def build_batch_prompt(batch_items):
    """
    把一批 (行号, 英文) 拼成一次请求的 user 内容。
    要求模型按固定格式输出，方便我们用程序把译文“对号入座”：
        1. 译文一
        2. 译文二
        ...
    """
    numbered = "\n".join(f"{n}. {en}" for n, en in batch_items)
    return (
        "请把下面每一条英文翻译成中文，并按原编号逐行输出。\n"
        "严格要求：\n"
        "1) 每条译文单独一行，行首必须是编号和一个英文句点，格式为“1. 译文”；\n"
        "2) 不要输出任何解释、原文、空行或思考过程；\n"
        "3) 即使某一条很短，也必须占一行，总行数必须和输入条数一致。\n\n"
        f"待翻译内容：\n{numbered}"
    )


def parse_batch_result(content, expected_numbers):
    """
    从模型的批量回复里解析出 {行号: 译文}。
    用正则匹配行首的“数字.”，只认我们期望的行号；解析不到就返回 None 交由上层降级处理。
    """
    parsed = {}
    # 正则含义：行首可选空白 → 数字 → 英文句点/顿号/右括号 → 译文内容
    pattern = re.compile(r"^\s*(\d+)\s*[.、)]\s*(.+?)\s*$")
    for line in content.splitlines():
        m = pattern.match(line)
        if not m:
            continue                      # 跳过空行和不符合格式的行
        num = int(m.group(1))
        if num in expected_numbers:       # 只收本批的行号，防止模型串行
            parsed[num] = m.group(2)
    return parsed


def translate_batch(client, batch_items, system_prompt, glossary):
    """
    翻译【一批】文本。返回 {行号: 译文} 字典；失败时返回空字典，由调用方降级逐行翻译。
    batch_items 形如 [(excel行号, 英文文本), ...]
    """
    expected_numbers = {n for n, _ in batch_items}
    user_prompt = build_batch_prompt(batch_items)
    # 批量时把术语库约束也带上（和单条模式的 system_prompt 一致）
    batch_system = system_prompt

    # --- 第一阶段：批量调用 + 出错重试 ---
    content = None
    for attempt in range(FAIL_RETRY_TIMES + 1):
        try:
            content = call_model(client, batch_system, user_prompt)
            if not content:
                raise ValueError("模型返回了空内容")
            break
        except Exception as e:
            print(f"  ⚠️ 批量调用出错（第 {attempt + 1} 次）：{e}")
            if attempt < FAIL_RETRY_TIMES:
                time.sleep(RETRY_WAIT_SECONDS)

    if content is None:
        return {}   # 批量彻底失败，返回空，上层会逐行补翻

    # --- 第二阶段：解析 + 术语校验，不通过则整批重试一次 ---
    for _ in range(MAX_TERM_RETRY + 1):
        parsed = parse_batch_result(content, expected_numbers)

        # 校验一：行数必须凑齐，缺任何一行都视为失败
        if len(parsed) < len(batch_items):
            missing_nums = sorted(expected_numbers - set(parsed.keys()))
            print(f"  ⚠️ 批量解析不完整（缺行：{missing_nums}），正在重试该批……")
        else:
            # 校验二：整批拼起来做术语检查（和单条模式同一套逻辑）
            joined_source = "\n".join(en for _, en in batch_items)
            joined_result = "\n".join(parsed[n] for n, _ in batch_items)
            missing_terms = find_missing_terms(joined_source, joined_result, glossary) if glossary else []
            if not missing_terms:
                return parsed   # 全部通过，返回结果
            print(f"  ⚠️ 批量术语校验未通过（{len(missing_terms)} 个术语缺失），正在重试该批……")

        # 走到这里说明本批需要重试
        try:
            content = call_model(client, batch_system, user_prompt)
            if not content:
                raise ValueError("模型返回了空内容")
        except Exception as e:
            print(f"  ⚠️ 批量重试出错：{e}")
            return {}
        time.sleep(REQUEST_INTERVAL)

    # 重试机会用完仍不合格：返回空字典，上层逐行兜底
    return {}


# ========== 4. Excel 处理逻辑（批量 + 逐行兜底） ==========
def process_excel(client, system_prompt, glossary):
    if not os.path.exists(input_excel):
        print("⏭️ 未检测到 input.xlsx，跳过 Excel 处理。")
        return

    print("📊 开始处理 Excel ...")
    df = pd.read_excel(input_excel, header=None)
    total = len(df)

    # 4.1 先把 A 列全部读出来，做好清洗，并记录“原始行号”
    #     items 形如 [(0, "apple"), (1, "machine learning"), ...]，空行不放进列表
    items = []
    for i in range(total):
        cell = df.iloc[i, 0]
        # ⚠️ 顺序很重要：必须先判断 pd.isna，再做 str()，
        # 否则空单元格会被 str() 变成字符串 "nan"，被当成一个正常单词送去翻译
        if pd.isna(cell):
            continue
        text = str(cell).strip()
        if not text:
            continue
        items.append((i, text))

    # 4.2 用一个字典存放所有翻译结果：{原始行号: 译文}
    results = {}

    if ENABLE_BATCH and items:
        # --- 批量模式：每 BATCH_SIZE 行切成一批 ---
        batches = [items[k: k + BATCH_SIZE] for k in range(0, len(items), BATCH_SIZE)]
        print(f"  共 {len(items)} 行待翻译，分为 {len(batches)} 批（每批最多 {BATCH_SIZE} 行）...")

        for b_idx, batch_items in enumerate(batches, start=1):
            print(f"  ⏳ 正在批量翻译第 {b_idx}/{len(batches)} 批（行 {batch_items[0][0] + 1}~{batch_items[-1][0] + 1}）...")
            parsed = translate_batch(client, batch_items, system_prompt, glossary)

            if parsed:
                results.update(parsed)
                print(f"  ✅ 第 {b_idx}/{len(batches)} 批完成（{len(parsed)} 行）")
            else:
                # 批量失败：降级为逐行翻译，保证每一行都有结果
                print(f"  ⏬ 第 {b_idx}/{len(batches)} 批批量失败，改为逐行翻译兜底...")
                for row_no, en in batch_items:
                    print(f"    ⏳ 逐行翻译第 {row_no + 1} 行：{en}")
                    results[row_no] = translate_text(client, en, system_prompt, glossary)

            time.sleep(REQUEST_INTERVAL)

        # 4.3 二次兜底：批量跑完后检查有没有漏掉的行（理论上不该有）
        leftovers = [(n, t) for n, t in items if n not in results]
        for row_no, en in leftovers:
            print(f"  ⏬ 补翻遗漏的第 {row_no + 1} 行：{en}")
            results[row_no] = translate_text(client, en, system_prompt, glossary)
    else:
        # --- 逐行模式（把 ENABLE_BATCH 设为 False 时走这里） ---
        for row_no, en in items:
            print(f"  ⏳ 正在翻译第 {row_no + 1}/{total} 行：{en}")
            results[row_no] = translate_text(client, en, system_prompt, glossary)
            print(f"  ✅ 第 {row_no + 1}/{total} 行完成")
            time.sleep(REQUEST_INTERVAL)

    # 4.4 按原始行号把结果写进 B 列（没结果的行保持空白）
    df[1] = [results.get(i, "") for i in range(total)]
    df.to_excel(output_excel, index=False, header=False)
    print(f"✅ Excel 翻译完成，已保存：{output_excel}")


# ========== 5. Word 处理逻辑（逐段，保留段落对应关系） ==========
def process_word(client, system_prompt, glossary):
    if not os.path.exists(input_word):
        print("⏭️ 未检测到 input.docx，跳过 Word 处理。")
        return

    print("📄 开始处理 Word ...")
    doc = Document(input_word)
    new_doc = Document()
    total = len(doc.paragraphs)

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            new_doc.add_paragraph("")   # 空段落保留，保证译文段落和原文一一对应
            continue

        print(f"  ⏳ 正在调用 {MODEL_NAME}（第 {i + 1}/{total} 段）...")
        translated = translate_text(client, text, system_prompt, glossary)
        new_doc.add_paragraph(translated)
        print(f"  ✅ 第 {i + 1}/{total} 段完成")
        time.sleep(REQUEST_INTERVAL)

    new_doc.save(output_word)
    print(f"✅ Word 翻译完成，已保存：{output_word}")


# ========== 6. 主程序 ==========
def main():
    if not API_KEY or "你的keys" in API_KEY:
        print("❌ 请先把代码里的 API_KEY 换成你自己的真实 Key 再运行！")
        return

    mode = "批量" if ENABLE_BATCH else "逐行"
    print(f"🚀 启动 AI 本地化翻译工具 ({MODEL_NAME}，Excel {mode}模式)...")
    glossary = load_glossary(glossary_file)
    system_prompt = build_system_prompt(glossary)
    client = ZhipuAI(api_key=API_KEY)

    process_excel(client, system_prompt, glossary)
    process_word(client, system_prompt, glossary)

    print("🎉 全部任务执行完毕！请查看桌面上的 output.xlsx / output.docx。")


if __name__ == "__main__":
    main()
