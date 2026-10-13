"""
Word 文档翻译工具(tkinter 图形界面 · 双向翻译 · 格式保真 · 本地质检版）
=====================================================================
给不懂代码的客户使用：选方向 → 填 Key → 选文件 → 点「开始翻译」→ 看日志 → 拿结果。
界面功能：
    1. 翻译方向：单选「英译中」/「中译英」；
    2. 智谱 API Key 输入框（默认留空，留空时回退读环境变量 ZHIPUAI_API_KEY):
    3. 「选择Excel(术语库)」+「清除」按钮（不需要术语库时不选即可）；
    4. 「选择Word(待翻译)」按钮，弹出文件选择框；
    5. 「开始翻译」+「停止」+「打开输出文件夹」按钮 + 实时日志文本框。
术语库列含义（随方向自动切换）：
    · 英译中:A 列 = 英文（原文），B 列 = 中文（译文）；
    · 中译英:A 列 = 中文（原文），B 列 = 英文（译文）。
    两种方向都是「A 列原文 → B 列译文」，所以统一读成 {原文: 译文} 字典。
覆盖范围（都在原文件的完整副本上「原地替换文字」）：
    正文段落 + 表格（含嵌套表格）+ 页眉页脚 + 文本框/艺术字 + 脚注 + 尾注 + 批注。
技术要点：
    · 【格式保真】不新建文档，复制原文件后在副本上原地改；粗体、字号、标题样式、
      编号、图片、表格、页面设置全部保留；客户的原文件自始至终不碰。
    · 【段内混排】一段里只有几个词加粗/斜体时，用 <b>/<i>/<u> 标记让模型带回来，
      再按标记切成多个 run 还原格式（模型不听话时自动退回整段单格式，内容不丢）。
    · 【图片安全】python-docx 的 run.text = ... 会清空 run 的所有子元素，
      所以这里只改 w:t，绝不动 w:drawing / w:pict —— 图片和文本框不会被删。
    · 【防呆检测】先用底层 XML 查 w:txbxContent / 批注锚点 / 修订痕迹，
      本工具不处理的区域会在日志和报告里明确写出来，绝不静默漏翻。
    · 【术语库静默失效防线】术语库读到 0 条、或 A 列语言与方向不符时，
      日志和《翻译报告.txt》都会报警（这是最容易被忽略的失败）。
    · 含超链接或域代码（页码、目录、交叉引用）的段落整段跳过，不翻花，写进报告。
    · tkinter 不能跨线程刷新控件 → 耗时任务放后台线程，日志用 queue.Queue 传回主线程。
    · 术语库作为硬约束拼进 system 提示词，并做三重译后校验（术语/中英混杂/数字）+
      定向重译。
    · 【并发请求，排队写入】多个线程只负责向 API 发请求（ThreadPoolExecutor），
      拿到译文后回到翻译主线程按段落顺序写回 Word——Word 永远只有一个写入者，
      所以多线程不会把文档写坏；请求之间用全局节流控制 QPS，避免触发限流。
    · 【合批】连续的短段落会拼成「1. … / 2. …」一次请求翻完，API 调用次数大幅下降；
      合批结果一律按编号逐条校验，条数对不上或某条不合格就退回逐条翻译/单独重翻。
    · 【术语库按需裁剪】每段只把「这段原文里真正出现的术语」拼进提示词，
      提示词更短、模型更专注、更省钱；裁剪口径与术语校验完全一致。
    · 【断点续跑】每翻 20 段存一次进度文件；中途停止或关掉窗口后，再启动同一个文档
      会询问「是否从第 N 段继续」，避免已经花掉的 API 钱白费。
    · 【进度条】ttk.Progressbar + 「已翻 x/y 段 · 已用 · 预计还需」，
      刻意不依赖 tqdm——GUI 用 .pyw 启动时没有控制台，tqdm 的进度条根本看不见。
    · 【双服务商】界面上可切「智谱 GLM / DeepSeek」：两家都是 OpenAI 风格的接口，
      翻译主流程一行不用改，只在建客户端时换实现和模型名（DeepSeek 需 pip install openai）；
      服务商与模型名都会写进进度文件，换服务商续跑会被判无效（两家的译文风格不一致），
      必须重新开始，避免半篇智谱半篇 DeepSeek。
    · 空段直接跳过，不调用 API，不浪费 Token。
    · 模型输出里的 \\" \\' \\[ \\] 会被清理掉再写进 Word（见 clean_text）。
    · 网络类异常自动重试、单次请求有超时；「停止」按钮处理完当前段就停并保存已翻部分。
    · 每次运行生成《翻译报告.txt》：统计 + 【质检评分】+ 需要人工处理的段落清单。
    · 【统计口径】正文、文本框、脚注、尾注、批注共用一份「质检记账本」：附加部件的失败
      同样计入 api_failed / write_failed，完整性维度的分母 = 实际处理过的段数
      （正文 + 附加部件），空段和被跳过的域段落不参与分母。
    · 全流程 try-except：单段失败回退原文，整体错误弹窗 + 日志打印堆栈。
已知限制（都会写进报告，不会静默）：
    · 混排格式只还原「相对主体格式多出来的」加粗/斜体/下划线；
      主体本身加粗、个别词不加粗这种反向情况不还原。
    · 文本框内部不做混排格式还原（内容和字体保住，段内强调会统一）。
    · 含超链接/域代码的段落跳过不翻；艺术字的文字效果（描边、渐变）不受影响但也不重排。
运行环境：Windows + Python 3.8~3.12
    pip install pandas openpyxl python-docx zhipuai
    （想用 DeepSeek 再加装一个：pip install openai）
    （tkinter 是标准库自带；lxml 是 python-docx 的依赖，装 python-docx 时就有了）
运行：python translate_tool_v3.py   （存成 .pyw 双击运行则不弹黑框）
"""
import os
import re
import copy
import json
import time
import queue
import shutil
import zipfile
import threading
import traceback
import tkinter as tk
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from tkinter import ttk, filedialog, messagebox
import pandas as pd                    # 读取术语库 Excel
from docx import Document              # python-docx：读写 Word
from docx.oxml.ns import qn            # 把 "w:hyperlink" 这类名字转成 lxml 能用的格式
from docx.text.run import Run          # 手工新建 run 时要用（还原混排格式）
from lxml import etree                 # 直接看 docx 底层 XML（文本框/批注都在段落树之外）
from zhipuai import ZhipuAI            # 智谱 AI 官方 SDK
# ============================ 可调参数 ============================
MODEL_NAME        = "glm-4-plus"   # 智谱模型名，可换成 glm-4-air / glm-4-flash
# —— ★ 新增：多服务商（智谱 GLM / DeepSeek）——
# 两家的接口都是 OpenAI 风格（client.chat.completions.create），
# 所以整套翻译逻辑一行都不用改，只换「客户端对象 + 模型名」。
PROVIDER_ZHIPU      = "智谱 GLM"                    # 默认服务商
PROVIDER_DEEPSEEK   = "DeepSeek"                    # 备选服务商（OpenAI 兼容接口）
DEEPSEEK_MODEL_NAME = "deepseek-chat"               # DeepSeek 模型名
DEEPSEEK_BASE_URL   = "https://api.deepseek.com"    # DeepSeek 接口地址
_ACTIVE_MODEL       = [MODEL_NAME]                  # 本次实际使用的模型名（建客户端时写入）
REQUEST_INTERVAL  = 1.0             # 每段之间的间隔秒数，防止触发限流
REQUEST_TIMEOUT   = 60.0            # 单次 API 请求超时秒数（旧版 SDK 不支持时自动忽略）
API_MAX_RETRY     = 2               # 网络/限流类异常自动重试次数
API_RETRY_BACKOFF = 3.0             # 重试等待基数（秒）：第 n 次重试等待 n × 基数
ENABLE_TERM_CHECK   = True          # 是否开启「译文术语校验」
MAX_TERM_RETRY      = 1             # 术语校验不通过时，最多重译几次
ENABLE_NO_MIX_CHECK = True          # 是否开启「中英混杂校验」（译文残留源语言就重译）
MAX_NO_MIX_RETRY    = 1             # 混杂校验最多重译几次
MIN_LATIN_RUN_WORDS = 4             # 英译中时：连续多少个英文单词才算「整句没翻」
ENABLE_NUMBER_CHECK = True          # 是否开启「数字一致性校验」（数字消失=疑似漏译）
MAX_NUMBER_RETRY    = 1             # 数字校验最多重译几次
MIN_CHECK_NUMBER_DIGITS = 2         # 只查「几位数以上」的数字（见 find_number_mismatch 说明）
ENABLE_FOOTNOTE_TRANSLATE = True    # 是否翻译脚注正文
ENABLE_ENDNOTE_TRANSLATE  = True    # ★ 新增：是否翻译尾注正文（word/endnotes.xml）
ENABLE_COMMENT_TRANSLATE  = True    # 是否翻译批注正文（批注是审阅人写的，想保持原样就设 False）
ENABLE_TEXTBOX_TRANSLATE  = True    # 是否翻译文本框/艺术字里的文字
ENABLE_FORMAT_MARKUP      = True    # 是否开启「段内混排格式还原」（<b>/<i>/<u> 标记法）
ENABLE_DOC_FEATURE_CHECK  = True    # 是否开启文本框/批注/修订的「防呆检测」
# —— ★ 新增：并发 / 合批 / 术语裁剪 / 断点续跑 / 进度 ——
# 核心安全原则：「并发请求，排队写入」——多个线程只负责向 API 发请求，
# 拿到结果后回到唯一的翻译线程按顺序写回 Word。Word 文件永远只有一个写入者。
ENABLE_CONCURRENCY      = True      # 多线程并发发请求（关掉就退回逐段串行）
API_WORKERS             = 4         # 并发线程数。智谱有 QPS 限流，4 比较稳；报 429 就调小
GLOBAL_REQUEST_INTERVAL = 0.5       # 并发模式下「发起请求」的最小间隔秒数（全局节流，防限流）
REQUEST_WINDOW_FACTOR   = 2         # 预取窗口 = API_WORKERS × 该系数（限制内存，也限制「停止」的延迟）
ENABLE_BATCH            = True      # 短段落合批：把多条拼成一次请求，API 调用次数可降 ~70%
BATCH_MAX_ITEMS         = 8         # 一批最多合并几条
BATCH_MAX_CHARS         = 600       # 一批原文合计字符上限（超了就另起一批）
ENABLE_GLOSSARY_PRUNE   = True      # 术语库按需裁剪：只把「这批真用得到的术语」塞进提示词
ENABLE_RESUME           = True      # 断点续跑：中断后下次启动问「是否从第 N 段继续」
PROGRESS_SAVE_EVERY     = 20        # 每成功写入多少段，存一次进度文件
# —— 本地质检评分（零 API 费用）：权重只对「有数据」的维度生效，缺数据自动归一化 ——
WEIGHTS          = {"完整性": 30, "术语": 30, "语言一致性": 20, "数字一致": 10, "格式还原": 10}
PASS_SCORE       = 85               # ≥ 85 判定「通过」
WARN_SCORE       = 70               # 70~85 判定「需抽查」
RED_TERM_RATE    = 0.98             # 硬红线：术语合规率低于它直接不通过
RED_TARGET_RATIO = 0.60             # 硬红线：单个段落的目标语言字符占比低于它要人工复核
MIN_LENGTH_FOR_LANG_CHECK = 20      # 太短的段落不参与「目标语言占比」判断（误报多）
MANUAL_REVIEW_LIMIT = 30            # 报告里「疑似需人工确认」最多列多少条
WINDOW_TITLE      = "Word 文档翻译工具（智谱 GLM · 双向翻译 · 术语库硬约束 · 本地质检）"
WINDOW_SIZE       = "960x760"
DIRECTION_EN2ZH = "英译中"          # A 列英文 → B 列中文
DIRECTION_ZH2EN = "中译英"          # A 列中文 → B 列英文
# 表头/说明行关键字：术语库首行常写 “English / 中文” 之类，命中就跳过，别当术语
HEADER_WORDS = {
    "english", "en", "chinese", "zh", "term", "terms", "source", "target",
    "原文", "译文", "英文", "中文", "术语", "术语表",
}
CJK_PATTERN = re.compile(r"[\u4e00-\u9fff]")        # 是否含汉字
LATIN_PATTERN = re.compile(r"[A-Za-z]")             # 英文字母
NUMBER_PATTERN = re.compile(r"\d+")                 # 数字
# 出现这些关键字说明是「临时性故障」，值得重试；Key 无效、余额不足之类不在其中
RETRYABLE_KEYWORDS = (
    "timeout", "timed out", "connection", "connect", "network", "reset",
    "429", "rate limit", "too many requests", "1301", "1302",
    "502", "503", "504", "overloaded", "temporarily", "server error",
)
# 附加文字部件在 docx（本质是 zip）里的位置
FOOTNOTES_PART = "word/footnotes.xml"
ENDNOTES_PART  = "word/endnotes.xml"    # ★ 新增：尾注正文（结构和脚注一模一样）
COMMENTS_PART  = "word/comments.xml"
# 本工具「不处理」的区域：检测到就提醒客户，避免客户以为全文都翻了
UNSUPPORTED_REGION_PARTS   = ("document.xml", "footnotes.xml", "endnotes.xml")
TEXTBOX_LOCAL_NAMES        = ("txbxContent",)      # 文本框、艺术字的文字都装在这里面
COMMENT_ANCHOR_LOCAL_NAMES = ("commentRangeStart", "commentRangeEnd", "commentReference")
REVISION_LOCAL_NAMES       = ("ins", "del", "moveFrom", "moveTo")   # 修订痕迹
# 段内混排格式用的标记
MARKUP_TAG_PATTERN = re.compile(r"</?([biu])>")                     # 我们自己要求的三种标记
ANY_TAG_PATTERN    = re.compile(r"</?[A-Za-z][A-Za-z0-9]*\s*/?>")   # 任何 HTML 风格标记（兜底清理）
XML_SPACE          = "{http://www.w3.org/XML/1998/namespace}space"  # 保住文字首尾空格的属性
MARKUP_SYSTEM_RULES = (
    "\n\n【格式标记规则（必须严格遵守）】\n"
    "用户输入的原文里可能出现 <b>...</b>、<i>...</i>、<u>...</u> 三种标记，"
    "分别表示该片段在原文中是加粗、斜体、下划线。\n"
    "1. 翻译时必须保留这些标记：把对应译文用完全相同的标记包起来，"
    "标记覆盖的强调范围要和原文一致，不要移动标记的位置。\n"
    "2. 标记内的文字要完整翻译，一个词都不能漏。\n"
    "3. 除了 <b> </b> <i> </i> <u> </u>，不要输出任何其他标记、注释或说明。"
)
# ★ 新增：合批（一次请求翻多条）时追加的规则
BATCH_SYSTEM_RULES = (
    "\n\n【多条翻译规则（必须严格遵守）】\n"
    "这次用户会一次给你多条待翻译文本，每条前面都带「序号. 」编号。\n"
    "1. 必须逐条翻译，每一条单独翻，禁止把几条合并成一句。\n"
    "2. 输出的条数和编号必须与输入完全一致：每条一行，以相同编号开头，"
    "编号后紧跟译文；禁止增加、删减、重新排序、合并。\n"
    "3. 即使某条很短（一个词、一个编号）也要照翻，不得跳过。\n"
    "4. 除了「编号. 译文」之外，不要输出任何解释、注释、前后缀或多余文字。"
)
# 解析合批返回：行首的「1. 」「2、」「3) 」都认
BATCH_ITEM_PATTERN = re.compile(r"^\s*(\d+)\s*[\.、\)）]\s*(.*)$")
# =====================================================================
#                        第一部分：核心翻译逻辑
# =====================================================================
def load_glossary(file_path, direction=DIRECTION_EN2ZH, log=print):
    """
    用 pandas 读取术语库 Excel，返回 {原文: 译文} 字典。
    列的含义随翻译方向变化：
        · 英译中：A 列 = 英文（原文），B 列 = 中文（译文）；
        · 中译英：A 列 = 中文（原文），B 列 = 英文（译文）。
    两种情况都是「A 列原文、B 列译文」，所以取列下标的代码完全一致。
    · header=None → 不把第一行当表头，直接按位置取 A/B 列；
    · dtype=str   → 全部按文本读，避免 “5G”“2D” 被转成 5.0 这类脏数据；
    · 读不到文件 / 读取出错 → 返回空字典，退化成「无术语约束的普通翻译」，不崩。
    """
    glossary = {}
    try:
        if not file_path:
            log("ℹ️ 未选择术语库，本次按普通翻译处理。")
            return glossary
        if not os.path.exists(file_path):
            log(f"⚠️ 术语库文件不存在：{file_path}（本次按普通翻译处理）")
            return glossary
        if direction == DIRECTION_ZH2EN:
            column_desc = "A 列中文（原文）/ B 列英文（译文）"
        else:
            column_desc = "A 列英文（原文）/ B 列中文（译文）"
        dataframe = pd.read_excel(file_path, header=None, dtype=str)
        for _, row in dataframe.iterrows():
            source_term = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
            target_term = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            if not source_term or not target_term:      # 空行 / 只填了一半，跳过
                continue
            if source_term.lower() in HEADER_WORDS:     # 表头行，跳过
                continue
            glossary[source_term] = target_term         # 原文 → 译文
        if glossary:
            log(f"✅ 术语库加载成功：共 {len(glossary)} 条术语（{column_desc}）"
                f" → {os.path.basename(file_path)}")
            warning = check_glossary_direction(glossary, direction)
            if warning:
                log("⚠️ " + warning)
        else:
            log("⚠️ 术语库读到了 0 条术语！本次全程没有任何术语约束，请检查："
                "① A 列和 B 列是否都填了（两列缺一即跳过）；"
                "② 是否整个文件只有一行表头；"
                f"③ 数据是否为「{column_desc}」。")
    except Exception as error:
        log(f"⚠️ 术语库读取失败：{error}（本次按普通翻译处理）")
    return glossary
def check_glossary_direction(glossary, direction=DIRECTION_EN2ZH):
    """
    术语库方向自检：判断 A 列填的语言和当前翻译方向是否一致，没问题是空串。
    为什么要有这个检查：如果客户把「中译英」的术语库按英译中顺序填
    （A 列英文 / B 列中文），程序不会报任何错——只是一条术语都命中不了、
    日志全绿地跑完，客户以为术语库生效了。这是最危险的一类「静默错误」。
    """
    if not glossary:
        return ""
    chinese_ratio = sum(1 for term in glossary if CJK_PATTERN.search(term)) / len(glossary)
    if direction == DIRECTION_ZH2EN and chinese_ratio < 0.2:
        return ("术语库 A 列看起来几乎全是英文，但当前方向是「中译英」"
                "（应 A 列中文 / B 列英文），很可能把方向填反了，"
                "本次术语可能一条都命中不了，请核对术语库！")
    if direction == DIRECTION_EN2ZH and chinese_ratio > 0.8:
        return ("术语库 A 列看起来几乎全是中文，但当前方向是「英译中」"
                "（应 A 列英文 / B 列中文），很可能把方向填反了，"
                "本次术语可能一条都命中不了，请核对术语库！")
    return ""
def build_empty_glossary_warning(glossary_file, glossary):
    """
    术语库「选了文件、却一条都没读到」时的告警文本；正常情况返回空串。
    为什么要单独写这个：这种情况程序不会抛任何异常，客户会以为术语库生效了，
    属于最危险的一类静默失效——必须同时出现在日志和《翻译报告.txt》里。
    """
    if glossary or not glossary_file:
        return ""
    return ("选定了术语库，但读到了 0 条术语：本次全程没有任何术语约束。"
            "请检查 A、B 两列是否都填了、是否整个文件只有一行表头。")
def _common_rules(target_language, start_index):
    """
    两个方向、有无术语库都通用的收尾规则，返回带序号的多行文本。
    start_index：起始序号。有术语库时接在「1~4 条术语规则」后面（传 5），
                 没有术语库时从 1 开始（传 1）。
    第 1 条（严禁中英混杂）是客户实测反馈加上去的：
        中译英时模型偶尔把「偶然」译成「偶然ly」这种半中半英的写法，
        必须用提示词明确禁止，否则单靠模型自觉不可靠。
    """
    return (
        f"{start_index}. 严禁中英混杂：译文必须通篇是{target_language}，"
        "禁止在一种语言的词汇后面直接拼接另一种语言的后缀"
        "（例如「偶然ly」「数据base」这类写法一律不允许），"
        "每个词都必须完整翻译，不得只翻一半。\n"
        f"{start_index + 1}. 只输出{target_language}译文正文，"
        "禁止输出原文、禁止输出术语表本身、禁止输出任何解释或额外说明。\n"
        f"{start_index + 2}. 引号和方括号请直接输出 ' 和 \" 和 [ 和 ]，"
        "不要转义成 \\' 或 \\\" 或 \\[ 或 \\]。\n"
        f"{start_index + 3}. 原文里的数字、百分比、编号必须原样保留，"
        "不得遗漏、不得改写。"
    )
def build_system_prompt(glossary, direction=DIRECTION_EN2ZH):
    """
    按翻译方向拼接 system 提示词，并把术语库作为【硬约束】塞进去。
    提示词全文档只构造一次、反复复用，不会额外增加请求次数。
    """
    if direction == DIRECTION_ZH2EN:
        # 中译英：客户指定的提示词原文
        base_prompt = (
            "你是一个专业的翻译助手。"
            "请把用户输入的中文翻译成地道、准确的英文，符合英语母语者的表达习惯。"
        )
        target_language = "英文"
    else:
        base_prompt = (
            "你是一个专业的翻译助手。"
            "请把用户输入的英文翻译成准确、流畅、符合中文表达习惯的中文。"
        )
        target_language = "中文"
    if not glossary:
        return (
            f"{base_prompt}\n\n"
            "【强制翻译规则】\n"
            + _common_rules(target_language, start_index=1)
        )
    term_lines = "\n".join(f"- {source_term} => {target_term}"
                           for source_term, target_term in glossary.items())
    return (
        f"{base_prompt}\n\n"
        "【术语库（硬约束，优先级高于一切翻译习惯）】\n"
        f"{term_lines}\n\n"
        "【强制翻译规则】\n"
        "1. 遇到术语库里的词，必须 100% 使用术语库的标准翻译，不要自行发挥、"
        "不要同义替换、不要意译、不要改成更“通顺”的说法。\n"
        "2. 术语的单复数、大小写、时态变化同样必须按术语库执行。\n"
        "3. 术语库是硬约束：即使你认为术语库的译法不够优雅，也必须照用。\n"
        "4. 术语库之外的普通内容，按专业、准确、通顺的标准正常翻译。\n"
        + _common_rules(target_language, start_index=5)
    )
def _term_in_source(source_text, term):
    """
    判断术语是否出现在【原文】里。
    · 术语是中文（中译英场景）→ 直接判断包含关系；
    · 术语是英文（英译中场景）→ 用词边界匹配并忽略大小写，
      避免 “Port” 误命中 “Important”“Portable”。
    """
    if CJK_PATTERN.search(term):
        return term in source_text
    pattern = rf"(?<![A-Za-z]){re.escape(term)}(?![A-Za-z])"
    return re.search(pattern, source_text, flags=re.IGNORECASE) is not None
def _term_in_translation(translated_text, term):
    """
    判断【译文】里是否用了术语库给定的标准译法。
    · 标准译法是中文 → 直接判断包含关系；
    · 标准译法是英文 → 忽略大小写判断（Server / server 都算命中）。
    """
    if CJK_PATTERN.search(term):
        return term in translated_text
    return term.lower() in translated_text.lower()
def find_source_terms(source_text, glossary):
    """
    找出原文里出现过的术语，并做「最长匹配」过滤。
    术语之间常有包含关系（「数据」和「数据库」、「Port」和「Port Number」）：
    如果短术语被某个更长的、同样命中的术语包住，就只认长术语，
    否则「数据」会被判成「漏译」，白白触发一次重译、白花 Token。
    """
    if not glossary:
        return []
    present_terms = [term for term in glossary if _term_in_source(source_text, term)]
    return [term for term in present_terms
            if not any(term != longer and term in longer for longer in present_terms)]
def find_missing_terms(source_text, translated_text, glossary):
    """术语校验：返回 [(原文术语, 应有译文), ...]，空列表表示全部命中。"""
    missing_terms = []
    for source_term in find_source_terms(source_text, glossary):
        target_term = glossary[source_term]
        if not _term_in_translation(translated_text, target_term):
            missing_terms.append((source_term, target_term))
    return missing_terms
def find_mixed_language_residue(translated_text, direction=DIRECTION_EN2ZH):
    """
    中英混杂校验：找出译文里「没翻译干净」的残留，返回一句说明；没残留就返回空串。
    · 中译英（目标是英文）：出现【汉字】就是没翻干净，例如「偶然ly」。
      这是客户实测踩到的坑，按字符判定，最可靠。
    · 英译中（目标是中文）：不能一见到字母就报错——API、GDP、Word、公司名、
      以及术语库要求保留的英文词都属于正常。所以只有【连续 4 个以上英文单词】
      这种整句英文才算残留（阈值见 MIN_LATIN_RUN_WORDS）。
    """
    if not translated_text:
        return ""
    if direction == DIRECTION_ZH2EN:
        chinese_chars = CJK_PATTERN.findall(translated_text)
        if chinese_chars:
            sample = "".join(dict.fromkeys(chinese_chars))[:20]
            return f"译文里残留了中文：{sample}"
        return ""
    latin_run_pattern = re.compile(
        r"[A-Za-z][A-Za-z'\-]*(?:\s+[A-Za-z][A-Za-z'\-]*){%d,}" % (MIN_LATIN_RUN_WORDS - 1))
    match = latin_run_pattern.search(translated_text)
    if match:
        return f"译文里残留了整句英文：{match.group(0)[:40]}"
    return ""
def _checkable_numbers(text):
    """
    取出文本里「值得校验的数字」（两位及以上）。
    先把 1,000 的千分位逗号去掉，避免「1,000」和「1000」被当成两个不同的数。
    """
    cleaned = re.sub(r"(?<=\d),(?=\d)", "", text)
    return [number for number in NUMBER_PATTERN.findall(cleaned)
            if len(number) >= MIN_CHECK_NUMBER_DIGITS]
def has_checkable_number(text):
    """原文里是否含「两位及以上」的数字（数字校验只针对这类数字）。"""
    return bool(_checkable_numbers(text))
def find_number_mismatch(source_text, translated_text):
    """
    数字一致性校验：原文里「两位及以上」的数字，在译文里不能找不着。
    为什么只查两位以上：中文里一位数常用汉字表达（Chapter 3 → 第三章、第一名），
    拿一位数去比会产生大量误报；两位以上的数字中英文本里几乎都用阿拉伯数字。
    为什么这条有用：数字凭空消失，八成是整句/整段没翻译，或者模型自己「顺手改数」。
    """
    source_numbers = _checkable_numbers(source_text)
    if not source_numbers:
        return ""
    translated_numbers = _checkable_numbers(translated_text)
    missing_numbers = [number for number in source_numbers if number not in translated_numbers]
    if missing_numbers:
        return f"数字不一致：原文有 {('、'.join(missing_numbers))}，译文里找不到"
    return ""
def clean_text(text):
    """
    清理模型输出里多余的反斜杠转义。
    现象：GLM 偶尔会把符号输出成 \\" \\' \\[ \\]，
          例如 sparrow\\'s、\\"Things、脚注的 \\[1\\]。
    处理：\\' → '      \\" → "      \\[ → [      \\] → ]
    为什么「只替换这四种组合」而不是把反斜杠全删掉：
        译文里可能有正常用途的反斜杠（Windows 路径 C:\\Users、代码里的 \\n），
        一把梭全删会把这些内容改坏。
    幂等性：清过一遍的文本再清一遍不会有任何变化，可以放心在多处调用。
    """
    if not text:
        return text
    text = text.replace("\\'", "'")
    text = text.replace('\\"', '"')
    text = text.replace("\\[", "[")
    text = text.replace("\\]", "]")
    return text
def _is_retryable_error(error):
    """
    判断这个异常值不值得重试。
    超时、断网、限流、服务端 5xx → 重试有意义；
    API Key 无效、余额不足、参数错误 → 重试多少次都一样，直接放弃。
    """
    text = f"{type(error).__name__} {error}".lower()
    return any(keyword in text for keyword in RETRYABLE_KEYWORDS)
_GLOBAL_THROTTLE_LOCK = threading.Lock()
_LAST_REQUEST_TIME = [0.0]      # 用列表当可变容器（嵌套函数里改数字要 nonlocal，麻烦）
def throttle_requests():
    """
    ★ 新增：全局请求节流。
    为什么需要：并发 ≠ 可以把 QPS 拉到无限。智谱对 QPS 有限制，
    4 个线程放开猛发很容易吃 429（限流错误），反而更慢更贵。
    这里保证「两次请求发起」之间至少隔一会儿：
        · 并发模式：GLOBAL_REQUEST_INTERVAL（0.5 秒）
        · 串行模式：REQUEST_INTERVAL（1.0 秒，和旧版行为一致）
    多线程下用「锁 + 时间戳预约下一个发车时刻」，所以间隔是全局的，
    不是每个线程各算各的（各算各的等于没有节流）。
    """
    interval = GLOBAL_REQUEST_INTERVAL if ENABLE_CONCURRENCY else REQUEST_INTERVAL
    if interval <= 0:
        return
    with _GLOBAL_THROTTLE_LOCK:
        now = time.time()
        wait_seconds = interval - (now - _LAST_REQUEST_TIME[0])
        if wait_seconds < 0:
            wait_seconds = 0.0
        _LAST_REQUEST_TIME[0] = now + wait_seconds
    if wait_seconds:
        time.sleep(wait_seconds)
def call_glm(client, system_prompt, user_text, log=print):
    """
    调用一次 GLM 接口，返回清理过反斜杠转义、去掉首尾空白的译文。
    ★ 新增：从 translate_text 里抽出来变成公共函数，因为它现在有三个调用方——
      单段翻译、合批翻译、以及多线程并发（每个线程各调各的，互不干扰）。
    超时/断网/限流会自动重试（等待时间逐次拉长），其余异常直接抛出。
    """
    for attempt in range(API_MAX_RETRY + 1):
        try:
            throttle_requests()      # ★ 全局节流：多线程并发也不会把 QPS 拉爆
            response = client.chat.completions.create(
                model=_ACTIVE_MODEL[0],   # ★ 双服务商：跟着当前服务商走（智谱/DeepSeek）
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user",   "content": user_text},
                ],
                temperature=0.1,        # 低随机性，翻译更稳定
            )
            # 唯一的模型文本出口：在这里清理，所有路径一次改完
            return clean_text(response.choices[0].message.content.strip())
        except Exception as error:
            if attempt >= API_MAX_RETRY or not _is_retryable_error(error):
                raise
            wait_seconds = API_RETRY_BACKOFF * (attempt + 1)
            log(f"  ⏳ 网络/限流异常（{type(error).__name__}），"
                f"{wait_seconds:.0f} 秒后重试（第 {attempt + 1}/{API_MAX_RETRY} 次）……")
            time.sleep(wait_seconds)
    return ""
def build_client(provider, api_key):
    """
    ★ 新增：按「服务商」创建客户端，并把本次真正使用的模型名写进 _ACTIVE_MODEL。
    智谱和 DeepSeek 都提供 OpenAI 风格的接口（client.chat.completions.create），
    所以翻译主流程（call_glm 和后面所有校验逻辑）一行都不用改，只换客户端和模型名。
    DeepSeek 走官方 openai SDK：需要先 pip install openai。
    没装时给出人话提示（不是抛一堆栈让客户看不懂）。
    _ACTIVE_MODEL 用列表当可变容器（和 _LAST_REQUEST_TIME 一个写法）：
    日志、进度文件、《翻译报告.txt》都从它取模型名，切服务商就跟着变。
    """
    if provider == PROVIDER_DEEPSEEK:
        try:
            from openai import OpenAI          # 延迟导入：不用 DeepSeek 的客户不必装它
        except ImportError:
            raise RuntimeError(f"使用 {PROVIDER_DEEPSEEK} 需要先安装 openai：pip install openai")
        _ACTIVE_MODEL[0] = DEEPSEEK_MODEL_NAME
        try:
            return OpenAI(api_key=api_key, base_url=DEEPSEEK_BASE_URL,
                          timeout=REQUEST_TIMEOUT, max_retries=0)
        except TypeError:
            return OpenAI(api_key=api_key, base_url=DEEPSEEK_BASE_URL)   # 老版本 SDK
    _ACTIVE_MODEL[0] = MODEL_NAME
    try:
        return ZhipuAI(api_key=api_key, timeout=REQUEST_TIMEOUT, max_retries=0)
    except TypeError:
        return ZhipuAI(api_key=api_key)     # 旧版 SDK 不认这两个参数
def prune_glossary_for_text(text, glossary):
    """
    ★ 新增：术语库按需裁剪（上下文工程）。
    只保留「原文里真正出现的术语」，没出现的术语不再塞进提示词。
    为什么省钱又更准：
        · 客户术语库往往几百条，全文塞进去提示词很长，模型容易走神；
        · 一段话里通常只用得到两三条术语，裁完提示词短一大截；
        · 裁剪口径和 find_source_terms 完全一致（含最长匹配），
          所以「术语校验」的结果和用全量术语库时一模一样，不会漏判。
    返回裁剪后的 {原文: 译文}；关掉 ENABLE_GLOSSARY_PRUNE 时原样返回。
    """
    if not glossary or not ENABLE_GLOSSARY_PRUNE:
        return glossary
    return {term: glossary[term] for term in find_source_terms(text, glossary)}
_PROMPT_CACHE = {}      # 提示词缓存：(方向, 是否合批, 术语元组) → 拼好的 system 提示词
def build_prompt_for(glossary, direction, batch=False):
    """
    ★ 新增：按（裁剪后的）术语集构造 system 提示词，并缓存结果。
    为什么要缓存：并发/合批之后，几千段可能只有几十种不同的「术语子集」，
    每段都重新拼一次几百条术语的字符串是白白的 CPU 开销（而且都在抢 GIL）。
    """
    key = (direction, batch, tuple(glossary.items()) if glossary else None)
    cached = _PROMPT_CACHE.get(key)
    if cached is None:
        cached = build_system_prompt(glossary, direction)
        if batch:
            cached += BATCH_SYSTEM_RULES
        if len(_PROMPT_CACHE) > 256:     # 缓存无限长会吃内存，简单粗暴地清一次
            _PROMPT_CACHE.clear()
        _PROMPT_CACHE[key] = cached
    return cached
def _translate_one_item(client, source_item, glossary, direction, log=print,
                        should_stop=None, prompt_extra="",
                        api_source=None, check_source=None):
    """
    ★ 新增：翻译「一条」文本（单段翻译和合批回退都走它，保证口径一致）。
    做三件事：裁剪术语库 → 取缓存的提示词 → 交给 translate_text 做全套校验与重译。
    api_source / check_source 只有「段内混排」才用得上：发给模型的是带标记的原文，
    校验必须用不带标记的干净原文。
    """
    api_source = api_source if api_source is not None else source_item
    check_glossary = prune_glossary_for_text(check_source or source_item, glossary)
    prompt = build_prompt_for(check_glossary, direction) + prompt_extra
    diagnostics = {"retry": 0, "problems": []}
    translation = translate_text(
        client, api_source, prompt, check_glossary,
        direction=direction, log=log, should_stop=should_stop,
        diagnostics=diagnostics, source_for_check=check_source)
    return {"source": source_item, "translation": translation, "diagnostics": diagnostics}
def inspect_translation(source_text, translated_text, glossary, direction):
    """
    ★ 新增：单条译文的本地硬性校验（术语 / 中英混杂 / 数字），返回 [(类型, 说明), ...]。
    口径与 translate_text 内部的 collect_problems 完全一致；
    合批时用它筛选「哪一条翻坏了」，只把坏的那条拉出来单独重翻——不整批重翻。
    """
    problems = []
    if ENABLE_TERM_CHECK and glossary:
        missing_terms = find_missing_terms(source_text, translated_text, glossary)
        if missing_terms:
            problems.append(("术语未按术语库",
                             "、".join(f"{src}→{dst}" for src, dst in missing_terms)))
    if ENABLE_NO_MIX_CHECK:
        residue = find_mixed_language_residue(translated_text, direction)
        if residue:
            problems.append(("译文残留源语言", residue))
    if ENABLE_NUMBER_CHECK:
        number_issue = find_number_mismatch(source_text, translated_text)
        if number_issue:
            problems.append(("数字不一致", number_issue))
    return problems
def build_batch_user_text(source_items):
    """
    ★ 新增：把多条原文拼成「1. A / 2. B / 3. C」（每条一行）交给模型。
    条内换行、制表符替换成空格，保证「一条就是一行」——否则编号解析会乱。
    """
    lines = []
    for index, item in enumerate(source_items, start=1):
        single_line = re.sub(r"[\r\n\t]+", " ", item).strip()
        lines.append(f"{index}. {single_line}")
    return "\n".join(lines)
def split_batch_response(response_text, expected_count):
    """
    ★ 新增：把合批返回的「1. xxx / 2. yyy」解析回译文列表。
    解析不出来（条数对不上、编号缺失、某条是空的）就返回 None，
    调用方会整批退回逐条翻译。宁可多花一次请求的钱，也绝不把错位的译文写进 Word。
    """
    results = [None] * expected_count
    state = {"current": None, "buffer": []}
    def flush():
        current_index = state["current"]
        if current_index is not None and 1 <= current_index <= expected_count:
            results[current_index - 1] = " ".join(
                part for part in state["buffer"] if part).strip()
        state["buffer"] = []
    for raw_line in response_text.splitlines():
        match = BATCH_ITEM_PATTERN.match(raw_line)
        if match:
            flush()
            state["current"] = int(match.group(1))
            first_piece = match.group(2).strip()
            state["buffer"] = [first_piece] if first_piece else []
        elif state["current"] is not None:
            piece = raw_line.strip()
            if piece:
                state["buffer"].append(piece)
    flush()
    if any(not (item or "").strip() for item in results):
        return None
    return results
def translate_batch_items(client, source_items, glossary, direction, log=print,
                          should_stop=None):
    """
    ★ 新增：一次请求翻译多条短文本（合批）。
    返回 (结果列表, 批次诊断)：
        结果列表[i] = {"source": 原文, "translation": 译文, "diagnostics": {...}}
        批次诊断    = {"retry": 0, "problems": []}（合批调用本身的重译/失败记录）
    容错三层（一层比一层保守）：
        1. 合批请求异常或解析失败 → 整批退回逐条翻译（绝不把错位译文写进 Word）；
        2. 合批成功但某条本地校验不过 → 只把这一条单独重翻，其余照用；
        3. 单独重翻再失败 → translate_text 内部回退原文，并记账进报告。
    """
    batch_diagnostics = {"retry": 0, "problems": []}
    batch_source_text = "\n".join(source_items)
    prompt = build_prompt_for(prune_glossary_for_text(batch_source_text, glossary),
                              direction, batch=True)
    log(f"  📦 合批翻译 {len(source_items)} 条（少发 {len(source_items) - 1} 次请求）……")
    translations = None
    try:
        response_text = call_glm(client, prompt, build_batch_user_text(source_items), log)
        translations = split_batch_response(response_text, len(source_items))
    except Exception as error:
        log(f"  ⚠️ 合批请求失败（{error}），本批退回逐条翻译。")
        translations = None
    if translations is None:
        log("  ⚠️ 合批结果无效（请求失败或条数/编号对不上），本批退回逐条翻译。")
        results = []
        for source_item in source_items:
            results.append(_translate_one_item(
                client, source_item, glossary, direction,
                log=log, should_stop=should_stop))
        return results, batch_diagnostics
    results = []
    for source_item, translation in zip(source_items, translations):
        item_diagnostics = {"retry": 0, "problems": []}
        problems = inspect_translation(source_item, translation, glossary, direction)
        if problems and not (should_stop is not None and should_stop()):
            detail = "；".join(text for _, text in problems)
            log(f"  ⚠️ 合批里有 1 条不达标（{detail}），单独重翻这一条……")
            single = _translate_one_item(client, source_item, glossary, direction,
                                         log=log, should_stop=should_stop)
            results.append({"source": source_item,
                            "translation": single["translation"] or translation,
                            "diagnostics": single["diagnostics"]})
            continue
        results.append({"source": source_item, "translation": translation,
                        "diagnostics": item_diagnostics})
    return results, batch_diagnostics
def translate_text(client, text, system_prompt, glossary=None,
                   direction=DIRECTION_EN2ZH, log=print,
                   should_stop=None, diagnostics=None, source_for_check=None):
    """
    翻译一段文本。容错：任何异常都返回原文，绝不让整个流程中断。
    参数：
        direction        ：翻译方向（「中译英」/「英译中」）；
        should_stop      ：可调用的「是否要停」判断，点了「停止」就跳过剩下的重译；
        diagnostics      ：可选字典，回传 {"retry": 重译次数,
                           "problems": [{"kind":…, "detail":…}]}，用来写报告；
        source_for_check ：当「发给模型的原文」和「用来校验的原文」不一样时传它。
                           段内混排段落走这个模式：发给模型的是带 <b>/<i>/<u> 标记的原文，
                           校验必须用不带标记的干净原文；
                           同时译文里的标记在比对前也会被剥掉
                           （否则「<b>数据库</b>」会被误判成术语没用上）。
                           传 None（默认）= 普通模式，不做任何剥标记处理。
    译后校验（哪个开关打开就查哪个，多个问题在【同一轮】里合并成一次重译）：
        · 术语校验：原文出现的术语，译文必须用术语库的标准译法；
        · 混杂校验：译文不得残留源语言文字；
        · 数字校验：原文里的多位数字在译文里不能凭空消失。
    代价：只有校验不通过才多调 API，正常段落零额外开销；
         最坏情况每段最多多 1 次调用（轮数由 MAX_*_RETRY 控制）。
    """
    if diagnostics is not None:
        diagnostics.setdefault("retry", 0)
        diagnostics.setdefault("problems", [])
    markup_mode = source_for_check is not None
    check_text = source_for_check if markup_mode else text
    def note(kind, detail):
        """记一条「需要人工处理」的信息，交给调用方写进报告。"""
        if diagnostics is not None:
            diagnostics["problems"].append({"kind": kind, "detail": str(detail)})
    def checkable(current_translation):
        """校验前把格式标记剥掉（只有段内混排模式才需要，普通段落里的尖括号可能是正文内容）。"""
        if markup_mode:
            return ANY_TAG_PATTERN.sub("", current_translation)
        return current_translation
    def call_model(current_system_prompt, user_text):
        """
        调用一次 GLM 接口（★ 重试逻辑已抽到模块级 call_glm，
        这样合批翻译和多线程并发可以复用同一套超时/重试规则）。
        """
        return call_glm(client, current_system_prompt, user_text, log)
    def collect_problems(current_translation):
        """
        检查当前译文有哪些硬性要求没满足。
        返回 ([(问题类型, 问题说明), ...], [重译提示行, ...])。
        """
        found_problems, force_lines = [], []
        target_language = "英文" if direction == DIRECTION_ZH2EN else "中文"
        comparison = checkable(current_translation)
        # 1) 术语硬约束
        if ENABLE_TERM_CHECK and glossary:
            missing = find_missing_terms(check_text, comparison, glossary)
            if missing:
                found_problems.append(
                    ("术语未按术语库", "、".join(f"{src}→{dst}" for src, dst in missing)))
                force_lines.extend(
                    f"- 「{src}」必须翻译成「{dst}」，不得使用其他译法" for src, dst in missing)
        # 2) 中英混杂残留
        if ENABLE_NO_MIX_CHECK:
            residue = find_mixed_language_residue(comparison, direction)
            if residue:
                found_problems.append(("译文残留源语言", residue))
                force_lines.append(
                    f"- {residue}；必须把残留的部分也完整翻译成{target_language}，"
                    "译文里不得出现任何未被翻译的原文文字")
        # 3) 数字一致性
        if ENABLE_NUMBER_CHECK:
            number_issue = find_number_mismatch(check_text, comparison)
            if number_issue:
                found_problems.append(("数字不一致", number_issue))
                force_lines.append(
                    f"- {number_issue}；说明可能有内容没翻到，请把原文每一部分都完整翻译，"
                    "数字、百分比、编号必须与原文一致")
        return found_problems, force_lines
    try:
        translated = call_model(system_prompt, text)
        retry_rounds = max(
            MAX_TERM_RETRY if (ENABLE_TERM_CHECK and glossary) else 0,
            MAX_NO_MIX_RETRY if ENABLE_NO_MIX_CHECK else 0,
            MAX_NUMBER_RETRY if ENABLE_NUMBER_CHECK else 0,
        )
        for attempt in range(retry_rounds):
            if should_stop is not None and should_stop():
                log("  ⏹ 已请求停止，跳过本段的重译。")
                break
            found_problems, force_lines = collect_problems(translated)
            if not found_problems:
                break       # 全部达标，不用重译
            problem_text = "；".join(detail for _, detail in found_problems)
            log(f"  ⚠️ 第 {attempt + 1} 次译后校验未通过：{problem_text}，正在重译……")
            if diagnostics is not None:
                diagnostics["retry"] += 1
            retry_prompt = (
                f"{system_prompt}\n\n"
                "【再次强调（上一版译文违反了下面的硬性要求）】\n"
                + "\n".join(force_lines) + "\n"
                "请重新输出完整译文，并严格满足上述全部要求。"
            )
            try:
                retranslated = call_model(retry_prompt, text)
            except Exception as retry_error:
                log(f"  ⚠️ 重译失败，保留上一版译文：{retry_error}")
                break
            if retranslated:
                translated = retranslated       # 模型返回空串时不覆盖，保住上一版译文
        # 收尾：重译后仍不达标的，按问题类型分别记进日志和报告
        final_problems, _ = collect_problems(translated)
        for kind, detail in final_problems:
            log(f"  ⚠️ 重译后仍未达标（{kind}）：{detail}")
            note(kind, detail)
        return translated
    except Exception as error:
        log(f"  ⚠️ 该段翻译失败，已回退为原文：{error}")
        note("API失败，已回退原文", error)
        return text          # 回退的是「原文」，不做清理，避免改坏客户原稿
# =====================================================================
#              第二部分：格式保真（混排格式 + 图片安全）
# =====================================================================
def _run_is_text_only(run):
    """
    判断这个 run 里除了格式（w:rPr）和文字（w:t）之外没有别的东西。
    为什么必须判断：python-docx 给 run.text 赋值时会【先清空这个 run 的所有子元素】
    （clear_content 只保留 w:rPr），所以 run 里挂着的图片（w:drawing）、
    文本框/图形（w:pict）、域（w:fldChar）、换行（w:br）都会被一起删掉。
    因此：只有「纯文字 run」才允许整体赋值，其它 run 一律只改 w:t。
    """
    for child in run._r:
        if child.tag in (qn("w:rPr"), qn("w:t")):
            continue
        return False
    return True
def _set_run_text_safely(run, new_text):
    """
    只改 run 里的 w:t 文字，绝不碰 run 里的图片 / 文本框 / 图形 / 换行。
    多余的 w:t 直接删掉（一段文字对应的 w:t 留一个就够）。
    返回 False 表示这个 run 里没有 w:t（只挂着图片之类），那就什么都不做。
    """
    text_nodes = run._r.findall(qn("w:t"))
    if not text_nodes:
        return False
    text_nodes[0].text = new_text
    if new_text != new_text.strip():
        text_nodes[0].set(XML_SPACE, "preserve")    # 保住首尾空格
    for extra_node in text_nodes[1:]:
        run._r.remove(extra_node)
    return True
def _clear_run_text(run):
    """清空一个 run 的文字，但保住 run 里的图片 / 文本框 / 图形 / 换行。"""
    if _run_is_text_only(run):
        run.text = ""
    else:
        _set_run_text_safely(run, "")
def replace_paragraph_text(paragraph, new_text):
    """
    把译好的文字整段写回，尽量保住原有格式和段落里的图片/文本框（混排段落用它兜底）。
    做法：挑一个「主体 run」——文字最长的那个——把译文写进去，其余 run 清空文字。
    挑最长的是因为：像 “**注意**：这是说明” 这种段落，第一个 run 是加粗的「注意」，
    用它当载体的话整段译文都会变加粗；挑文字最长的那个，拿到的一般就是正文格式。
    返回 False 表示这个段落没法安全写入（调用方会记进报告）。
    """
    runs = list(paragraph.runs)
    text_runs = [run for run in runs if run.text]
    if not text_runs:
        return False                # 没有文字（比如只挂着一张图），不碰
    anchor_run = max(text_runs, key=lambda run: len(run.text))
    for run in runs:
        if run is not anchor_run:
            _clear_run_text(run)    # 带图片的 run 只清文字，图片保住
    if _run_is_text_only(anchor_run):
        anchor_run.text = new_text
    else:
        _set_run_text_safely(anchor_run, new_text)
    return True
def _run_format_signature(run):
    """
    把 run 的「看得见的格式」压成一个可比较的元组，用来判断段内是否混排。
    只比最常见的几项（加粗/斜体/下划线/字号/字体名）：
    比得太多会把本来一致的段落误判成混排，白白多花钱。
    """
    font = run.font
    return (
        bool(run.bold),
        bool(run.italic),
        bool(run.underline),
        font.size.pt if font.size is not None else None,
        font.name,
    )
def _extra_emphasis_tags(signature, dominant_signature):
    """列出这个 run 比段落主体格式「多出来」的强调（b/i/u）。"""
    tags = []
    if signature[0] and not dominant_signature[0]:
        tags.append("b")
    if signature[1] and not dominant_signature[1]:
        tags.append("i")
    if signature[2] and not dominant_signature[2]:
        tags.append("u")
    return tags
def _paragraph_format_plan(paragraph):
    """
    分析段落的 run 格式，返回「写回方案」字典：
        anchor_run    用来当格式模板的 run（段落主体格式），没有则为 None；
        marked_source 发给模型的原文（混排时带 <b>/<i>/<u> 标记）；
        plain_source  不带标记的干净原文（校验用）；
        has_markup    是否真的生成了标记；
        used_tags     用到了哪几种标记。
    思路：先按「格式签名」把 run 分组，哪一类的文字最多，哪一类就是段落主体格式；
    比主体格式「多出」加粗/斜体/下划线的 run 会被包上对应标记交给模型。
    只表达「多出来的强调」：主体本身加粗、个别词不加粗这种反向情况不还原。
    """
    runs = list(paragraph.runs)
    plain_source = paragraph.text
    plan = {"anchor_run": None, "marked_source": plain_source, "plain_source": plain_source,
            "has_markup": False, "used_tags": set()}
    # 只考虑「纯文字 run」：带图片/文本框的 run 不能当格式模板，里面也不能塞标记
    candidate_runs = [run for run in runs if run.text.strip() and _run_is_text_only(run)]
    if not candidate_runs:
        return plan
    weight_by_signature = {}
    for run in candidate_runs:
        signature = _run_format_signature(run)
        weight_by_signature[signature] = weight_by_signature.get(signature, 0) + len(run.text)
    dominant_signature = max(weight_by_signature, key=weight_by_signature.get)
    plan["anchor_run"] = max(
        (run for run in candidate_runs if _run_format_signature(run) == dominant_signature),
        key=lambda run: len(run.text))
    # 只有真的混排（超过一种格式）才启用标记模式
    if not ENABLE_FORMAT_MARKUP or len(weight_by_signature) < 2:
        return plan
    # 原文里本来就带尖括号标签的段落不碰（会和我们的标记混淆）
    if ANY_TAG_PATTERN.search(plain_source):
        return plan
    marked_pieces, used_tags = [], set()
    for run in runs:
        piece = run.text
        if not piece:
            continue
        tags = (_extra_emphasis_tags(_run_format_signature(run), dominant_signature)
                if _run_is_text_only(run) else [])
        if tags:
            opening = "".join(f"<{tag}>" for tag in tags)
            closing = "".join(f"</{tag}>" for tag in reversed(tags))
            piece = f"{opening}{piece}{closing}"
            used_tags.update(tags)
        marked_pieces.append(piece)
    if not used_tags:
        return plan                 # 混排只是字号/字体不同，没法用标记表达
    plan["marked_source"] = "".join(marked_pieces)
    plan["has_markup"] = True
    plan["used_tags"] = used_tags
    return plan
def _strip_markup(text):
    """把所有 HTML 风格的标记从译文里去掉（模型不听话时的兜底）。"""
    return ANY_TAG_PATTERN.sub("", text)
def _split_marked_translation(marked_translation):
    """
    把模型返回的带标记译文切成 [(文字, 标记集合), ...]。
    容错解析：标记不配对也不会丢文字；相邻且标记相同的片段会合并，避免生成一堆碎 run。
    """
    segments = []
    active_tags = set()
    position = 0
    pending_text = ""
    for match in MARKUP_TAG_PATTERN.finditer(marked_translation):
        pending_text += marked_translation[position:match.start()]
        position = match.end()
        if pending_text:
            segments.append((pending_text, set(active_tags)))
            pending_text = ""
        if match.group(0).startswith("</"):
            active_tags.discard(match.group(1))     # 闭合标记
        else:
            active_tags.add(match.group(1))         # 开始标记
    pending_text += marked_translation[position:]
    if pending_text:
        segments.append((pending_text, set(active_tags)))
    merged_segments = []
    for text, tags in segments:
        text = _strip_markup(text)                  # 顺手清掉模型自己加的其它标记
        if not text:
            continue
        if merged_segments and merged_segments[-1][1] == tags:
            merged_segments[-1] = (merged_segments[-1][0] + text, tags)
        else:
            merged_segments.append((text, tags))
    return merged_segments
def _apply_tag_flags(run, tags):
    """给新 run 补上标记要求的格式（字体、字号已由模板 run 深拷贝带过来）。"""
    if "b" in tags:
        run.bold = True
    if "i" in tags:
        run.italic = True
    if "u" in tags:
        run.underline = True
def apply_marked_translation(paragraph, marked_translation, format_plan):
    """
    把（可能带标记的）译文写回段落，还原段内混排格式。
    返回实际使用的方式，供日志和报告统计：
        "标记还原"          —— 模型保留了标记，格式按片段还原；
        "退化为整段单格式"  —— 模型没保留标记，退回老办法（内容不丢，只是格式统一）；
        "无法写入"          —— 没法安全写入，一个字都不改。
    """
    anchor_run = format_plan.get("anchor_run")
    fallback_text = _strip_markup(marked_translation)
    segments = _split_marked_translation(marked_translation)
    if not segments or not fallback_text:
        return "无法写入"           # 模型返回空内容，宁可不动
    # 模型没保留任何标记 → 退回「整段一个格式」
    if not any(tags for _, tags in segments):
        return "退化为整段单格式" if replace_paragraph_text(paragraph, fallback_text) else "无法写入"
    # 主体 run 里挂着图片/文本框：深拷贝模板会把图片复制出来，绝对不能用
    if anchor_run is None or not _run_is_text_only(anchor_run):
        return "退化为整段单格式" if replace_paragraph_text(paragraph, fallback_text) else "无法写入"
    original_runs = list(paragraph.runs)        # 先记下原有 run（新插入的不能被当成要清空的）
    first_text, first_tags = segments[0]
    anchor_run.text = first_text                # 已确认是纯文字 run，可以整体赋值
    _apply_tag_flags(anchor_run, first_tags)
    template_element = copy.deepcopy(anchor_run._r)     # 格式模板（字体字号样式全在里面）
    last_element = anchor_run._r
    for text, tags in segments[1:]:
        new_element = copy.deepcopy(template_element)
        last_element.addnext(new_element)               # 依次插在主体 run 后面，保证文字顺序
        last_element = new_element
        new_run = Run(new_element, paragraph)
        new_run.text = text
        _apply_tag_flags(new_run, tags)
    for run in original_runs:
        if run is not anchor_run:
            _clear_run_text(run)                        # 原有 run 只清文字，图片保住
    return "标记还原"
# =====================================================================
#     第二部分之后、第三部分之前：质检「记账本」（四条路径共用同一份）
# ★ 本次修改：新增整节
# =====================================================================
# 为什么要有记账本：正文、文本框、脚注、批注是四个独立的循环。
# 以前只有正文循环在累加 api_failed / write_failed 等计数器，
# 文本框/脚注/批注的失败只写进了「需要人工处理的段落」清单、没进计数器，
# 结果「完整性」维度只反映了正文——附加部件翻失败了几段，分数照样很高。
# 现在四个循环共用同一份记账本（按引用传递），并且只有一个记账出口。
PROBLEM_API_FAILED   = "API失败，已回退原文"
PROBLEM_TERM         = "术语未按术语库"
PROBLEM_NUMBER       = "数字不一致"
PROBLEM_WRITE_FAILED = "段落结构特殊，未写入译文"
def make_stats_counter():
    """
    建一个质检统计「记账本」。四个循环共用同一份，全部按引用传递。
    各字段的含义（分母千万别漏，否则会出现「合规率 >100%」或负数这种荒唐结果）：
        translated_count  ：已成功写入译文的段数（正文 + 文本框 + 脚注 + 批注）
        body_count        ：其中正文写入的段数（写报告用）
        attachment_count  ：其中文本框/脚注/批注写入的段数
        api_failed        ：API 失败、回退成原文的段数
        write_failed      ：结构特殊、没能写进文档的段数
        same_count        ：译文与原文完全相同的段数
        term_paragraphs   ：含术语的段数（术语合规率的分母）
        term_bad          ：最终术语仍不达标的段数
        number_paragraphs ：含多位数字的段数（数字一致的分母）
        number_bad        ：最终数字对不上的段数
        attachment_failed ：其中属于文本框/脚注/批注的异常段数（写报告用）
        retry_total       ：定向重译总次数
        problems          ：需要人工处理的段落清单
    """
    return {
        "translated_count": 0,
        "body_count": 0,
        "attachment_count": 0,
        "api_failed": 0,
        "write_failed": 0,
        "same_count": 0,
        "term_paragraphs": 0,
        "term_bad": 0,
        "number_paragraphs": 0,
        "number_bad": 0,
        "attachment_failed": 0,
        "retry_total": 0,
        "problems": [],
    }
def count_source_features(counter, original_text, glossary):
    """
    按【原文】统计分母特征：这段是不是含术语、是不是含多位数字。
    必须在调用 API 之前调（统计的是原文，不是译文）。
    """
    if counter is None:
        return
    if glossary and find_source_terms(original_text, glossary):
        counter["term_paragraphs"] += 1
    if has_checkable_number(original_text):
        counter["number_paragraphs"] += 1
def count_written(counter, attachment=False):
    """
    记一段「已成功写入译文」。
    ★ 正文和附加部件都必须调这里：完整性维度的分母就是
      translated_count + write_failed（= 真正动手处理过的段落数），
      附加部件只要在这里 +1，分母自然就包含它们了。
    """
    if counter is None:
        return
    counter["translated_count"] += 1
    if attachment:
        counter["attachment_count"] += 1
    else:
        counter["body_count"] += 1
def record_problem(counter, kind, detail, location="", number=0, attachment=False):
    """
    记一条「需要人工处理」的问题：既写进清单，又累加到对应计数器。
    ★ 全流程只有这一个出口：正文、文本框、脚注、批注四条路径都调它，
      所以不会再出现「某一类失败只进清单、不进统计」的漏洞。
      attachment=True 表示这条问题来自文本框/脚注/批注（用于报告分类统计）。
    """
    if counter is None:
        return
    counter["problems"].append({"number": number, "location": location,
                                "kind": kind, "detail": str(detail)})
    if attachment:
        counter["attachment_failed"] += 1
    if kind == PROBLEM_API_FAILED:
        counter["api_failed"] += 1
    elif kind == PROBLEM_TERM:
        counter["term_bad"] += 1
    elif kind == PROBLEM_NUMBER:
        counter["number_bad"] += 1
    elif kind == PROBLEM_WRITE_FAILED:
        counter["write_failed"] += 1
    # 「译文残留源语言」这类问题只进清单，不参与任何评分维度
def record_diagnostics(counter, diagnostics, location="", number=0, attachment=False):
    """
    把 translate_text() 回传的 diagnostics（重译次数 + 问题清单）一次性并入记账本。
    四条路径都走这里，重译次数和问题计数就不会再漏。
    """
    if counter is None or diagnostics is None:
        return
    counter["retry_total"] += diagnostics.get("retry", 0)
    for problem in diagnostics.get("problems", []):
        record_problem(counter, problem["kind"], problem["detail"],
                       location, number, attachment)
# =====================================================================
#          第三部分：覆盖范围（文本框/艺术字、页眉页脚、脚注、批注）
# =====================================================================
def _has_special_content(paragraph):
    """
    判断段落里是否含超链接、域代码（页码、目录、交叉引用）。
    为什么要跳过这类段落：Word 的超链接和域不是普通的 run，
    直接把译文覆盖上去会把文字翻花、甚至把页码/目录的域代码弄坏，
    而且这种损坏客户很难发现。宁可整段不翻、在报告里列出来交给人工。
    """
    xml = paragraph._p
    if xml.findall(".//" + qn("w:fldChar")) or xml.findall(".//" + qn("w:instrText")):
        return True
    return xml.find(qn("w:hyperlink")) is not None or xml.find(qn("w:fldSimple")) is not None
def iter_translatable_paragraphs(doc):
    """
    遍历文档里所有「有文字可翻」的段落，每段只产出一次，并给出位置说明。
    覆盖范围：正文、正文表格（含嵌套表格）、每节的页眉/页脚（含首页/偶数页变体）。
    两个容易踩的坑，这里都处理了：
      1. 合并单元格：row.cells 会把同一个单元格重复返回多次 → 按 XML 元素去重，避免翻两遍；
      2. 链接到上一节的页眉/页脚：内容其实属于上一节，而且 python-docx 一读它的内容
         就会顺手给本节新建一个空定义（破坏「链接到上一节」）→ 先判断 is_linked_to_previous，
         是就整块跳过，只翻真正有独立内容的页眉页脚。
    """
    visited = set()         # 记录已经产出过的段落元素（lxml 元素按身份去重）
    def yield_paragraphs(paragraphs, location):
        for paragraph in paragraphs:
            element = paragraph._p
            if element in visited:
                continue
            visited.add(element)
            yield paragraph, location
    def yield_tables(tables, location):
        for table_index, table in enumerate(tables, start=1):
            for row_index, row in enumerate(table.rows, start=1):
                for column_index, cell in enumerate(row.cells, start=1):
                    cell_location = f"{location}表格{table_index} 第{row_index}行 第{column_index}列"
                    yield from yield_paragraphs(cell.paragraphs, cell_location)
                    yield from yield_tables(cell.tables, cell_location + " 内嵌")   # 嵌套表格
    yield from yield_paragraphs(doc.paragraphs, "正文")
    yield from yield_tables(doc.tables, "正文")
    section_count = len(doc.sections)
    for section_index, section in enumerate(doc.sections, start=1):
        containers = (
            ("页眉", section.header),
            ("页脚", section.footer),
            ("首页页眉", section.first_page_header),
            ("首页页脚", section.first_page_footer),
            ("偶数页页眉", section.even_page_header),
            ("偶数页页脚", section.even_page_footer),
        )
        for name, container in containers:
            if container.is_linked_to_previous:
                continue        # 见函数说明第 2 点
            label = f"{name}（第{section_index}节）" if section_count > 1 else name
            yield from yield_paragraphs(container.paragraphs, label)
            yield from yield_tables(container.tables, label)
def _iter_txbx_contents(root):
    """
    找出所有 w:txbxContent 节点（文本框和艺术字的文字容器），跳过 mc:Fallback 里的副本。
    mc:AlternateContent 会同时存两份等价内容：mc:Choice（新式，Word 实际显示这份）
    和 mc:Fallback（给老版本看的兼容副本）。两处文字一模一样，都翻会白花钱，
    所以只翻 Choice 那份，Fallback 整棵子树跳过。
    """
    def walk(node):
        for child in node:
            if not isinstance(child.tag, str):      # 注释、处理指令节点
                continue
            local_name = etree.QName(child).localname
            if local_name == "Fallback":
                continue                            # 兼容副本整棵子树跳过
            if local_name == "txbxContent":
                yield child                         # 文本框/艺术字的文字都装在这里
            yield from walk(child)                  # 继续往下找嵌套文本框
    yield from walk(root)
def _iter_textbox_roots(doc):
    """
    产出可能出现文本框/艺术字的「活」XML 树：正文 + 已存在的页眉页脚部件。
    这些部件都是 python-docx 的 XML 部件（正文是 DocumentPart，页眉页脚是 HeaderPart/
    FooterPart），改内存里的元素会被后面的 doc.save() 正常写回，不需要手动替换 zip。
    """
    yield "正文", doc.element
    package = doc.part.package
    parts = package.iter_parts() if hasattr(package, "iter_parts") else package.parts
    for part in parts:
        part_name = str(part.partname)
        if not part_name.startswith("/word/"):
            continue
        file_name = part_name.rsplit("/", 1)[-1]
        if not (file_name.startswith("header") or file_name.startswith("footer")):
            continue
        element = getattr(part, "_element", None)
        if element is not None:
            yield file_name, element
def translate_textboxes(doc, client, system_prompt, glossary, direction,
                        log=print, should_stop=None, counter=None):
    """
    翻译文本框 / 艺术字里的文字，返回翻译的段落数。
    就地改内存里的 XML，靠后面的 doc.save() 落盘（必须在 doc.save() 之前调用）。
    为什么能改：文本框和艺术字的正文文字都装在 w:txbxContent 里，
    它下面就是普通的 w:p / w:r / w:t，处理和正文一样——
    找到每个 w:p 的 w:t，第一个放译文、其余清空。不做任何解压提取。
    counter：四条路径共用的质检记账本（见 make_stats_counter）。
    """
    translated_count = 0
    box_index = 0
    for part_label, root in _iter_textbox_roots(doc):
        for txbx_content in _iter_txbx_contents(root):
            box_index += 1
            for paragraph_index, paragraph in enumerate(txbx_content.findall(qn("w:p")), start=1):
                text_nodes = paragraph.findall(".//" + qn("w:t"))
                original_text = "".join(node.text or "" for node in text_nodes).strip()
                if not original_text:
                    continue
                # 文本框里也可能有域代码（少见），跳过以免弄坏
                if (paragraph.findall(".//" + qn("w:fldChar"))
                        or paragraph.findall(".//" + qn("w:instrText"))):
                    continue
                location = f"{part_label} 文本框{box_index} 第{paragraph_index}段"
                # ★ 本次修改：分母（含术语 / 含多位数字）按原文统计
                count_source_features(counter, original_text, glossary)
                paragraph_diagnostics = {"retry": 0, "problems": []}
                translated = translate_text(
                    client, original_text,
                    build_prompt_for(prune_glossary_for_text(original_text, glossary),
                                     direction),         # ★ 术语同样按需裁剪
                    glossary,
                    direction=direction, log=log, should_stop=should_stop,
                    diagnostics=paragraph_diagnostics)
                # ★ 本次修改：失败与重译统一并入记账本（attachment=True 便于报告分类）
                record_diagnostics(counter, paragraph_diagnostics, location, attachment=True)
                # ★ 本次修改：防御性检查——没有文字节点就记「未写入」，绝不清空原文
                if not text_nodes:
                    record_problem(counter, PROBLEM_WRITE_FAILED,
                                   "该段落没有文字节点，未写入译文",
                                   location, attachment=True)
                    continue
                text_nodes[0].text = translated
                if translated != translated.strip():
                    text_nodes[0].set(XML_SPACE, "preserve")
                for extra_node in text_nodes[1:]:
                    extra_node.text = ""
                # ★ 本次修改：写入成功，分子分母一起记
                count_written(counter, attachment=True)
                translated_count += 1
                log(f"  🖼 {location}：{original_text[:20]}... → {translated[:20]}...")
                if should_stop is not None and should_stop():
                    log("  ⏹ 已请求停止，文本框翻译到此为止。")
                    return translated_count
    if translated_count:
        log(f"✅ 文本框/艺术字翻译完成：共 {translated_count} 段。")
    return translated_count
def _replace_zip_entry(docx_file, entry_name, new_bytes):
    """
    把 docx（本质是 zip 包）里的某一个部件换成新内容，其余部件原样复制。
    用「先写临时文件再原子替换」的方式，避免中途失败留下半个损坏的文档。
    """
    temp_file = docx_file + ".tmp"
    with zipfile.ZipFile(docx_file, "r") as source:
        with zipfile.ZipFile(temp_file, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = new_bytes if item.filename == entry_name else source.read(item.filename)
                target.writestr(item, data)
    os.replace(temp_file, docx_file)
def translate_extra_part_paragraphs(docx_file, part_name, entry_tag, skip_ids, label,
                                   client, system_prompt, glossary, direction,
                                   log=print, should_stop=None, counter=None):
    """
    翻译 docx 里某个「附加文字部件」的段落，返回翻译的段落数。
    脚注（word/footnotes.xml）和批注（word/comments.xml）共用这一套逻辑。
    为什么必须单独处理：这两个部件都是独立的，doc.paragraphs 根本看不见它们；
    而且 python-docx 把它们当普通二进制部件加载，改了内存里的树也不会被 doc.save()
    写回——所以只能在 doc.save() 之后，自己替换 zip 里的这一个部件。
    容错：任何异常都只记日志、保留原文，绝不影响已经翻好的正文。
    counter：四条路径共用的质检记账本（见 make_stats_counter）。
    """
    if not os.path.exists(docx_file):
        return 0
    try:
        with zipfile.ZipFile(docx_file, "r") as package:
            if part_name not in package.namelist():
                return 0                    # 这份文档没有这个部件
            raw_xml = package.read(part_name)
        root = etree.fromstring(raw_xml)
        translated_count = 0
        for entry in root.findall(qn(entry_tag)):
            entry_id = entry.get(qn("w:id"))
            if entry_id in skip_ids:
                continue                    # Word 用 -1/0 画脚注分隔线，不能翻
            for paragraph_index, paragraph in enumerate(entry.findall(qn("w:p")), start=1):
                text_nodes = paragraph.findall(".//" + qn("w:t"))
                original_text = "".join(node.text or "" for node in text_nodes).strip()
                if not original_text:
                    continue
                location = f"{label}{entry_id} 第{paragraph_index}段"
                # ★ 本次修改：分母（含术语 / 含多位数字）按原文统计
                count_source_features(counter, original_text, glossary)
                paragraph_diagnostics = {"retry": 0, "problems": []}
                translated = translate_text(
                    client, original_text,
                    build_prompt_for(prune_glossary_for_text(original_text, glossary),
                                     direction),         # ★ 术语同样按需裁剪
                    glossary,
                    direction=direction, log=log, should_stop=should_stop,
                    diagnostics=paragraph_diagnostics)
                # ★ 本次修改：失败（API 回退 / 术语 / 数字）与重译统一记账
                record_diagnostics(counter, paragraph_diagnostics, location, attachment=True)
                # ★ 本次修改：防御性检查——没有文字节点就记「未写入」
                if not text_nodes:
                    record_problem(counter, PROBLEM_WRITE_FAILED,
                                   "该段落没有文字节点，未写入译文",
                                   location, attachment=True)
                    continue
                text_nodes[0].text = translated
                if translated != translated.strip():
                    text_nodes[0].set(XML_SPACE, "preserve")
                for extra_node in text_nodes[1:]:
                    extra_node.text = ""
                # ★ 本次修改：写入成功，分子分母一起记
                count_written(counter, attachment=True)
                translated_count += 1
                if should_stop is not None and should_stop():
                    log(f"  ⏹ 已请求停止，{label}翻译到此为止。")
                    break
            if should_stop is not None and should_stop():
                break
        if not translated_count:
            return 0
        new_xml = etree.tostring(root, xml_declaration=True,
                                 encoding="UTF-8", standalone=True)
        _replace_zip_entry(docx_file, part_name, new_xml)
        log(f"✅ {label}翻译完成：共 {translated_count} 段。")
        return translated_count
    except Exception as error:
        log(f"⚠️ {label}翻译失败（{label}保持原文，正文不受影响）：{error}")
        return 0
def translate_footnotes_in_docx(docx_file, client, system_prompt, glossary, direction, **kwargs):
    """翻译脚注正文（跳过 id 为 -1/0 的分隔线脚注）。counter 通过 **kwargs 透传，无需改动。"""
    return translate_extra_part_paragraphs(
        docx_file, FOOTNOTES_PART, "w:footnote", ("-1", "0"), "脚注",
        client, system_prompt, glossary, direction, **kwargs)
def translate_comments_in_docx(docx_file, client, system_prompt, glossary, direction, **kwargs):
    """
    翻译批注正文（word/comments.xml）。
    批注是审阅人写的，语言通常和正文一致，所以按同一个方向翻。
    想让批注保持原样：把 ENABLE_COMMENT_TRANSLATE 改成 False。
    counter 通过 **kwargs 透传，无需改动。
    """
    return translate_extra_part_paragraphs(
        docx_file, COMMENTS_PART, "w:comment", (), "批注",
        client, system_prompt, glossary, direction, **kwargs)
def translate_endnotes_in_docx(docx_file, client, system_prompt, glossary, direction, **kwargs):
    """
    ★ 新增：翻译尾注正文（word/endnotes.xml）。
    尾注和脚注在 docx 里的结构完全一样，只是部件名和标签不同：
        · 脚注：word/footnotes.xml，标签 w:footnote
        · 尾注：word/endnotes.xml，标签 w:endnote
    所以直接复用 translate_extra_part_paragraphs，不用另写一套。
    跳过 id 为 -1/0 的分隔符尾注（Word 用它们画尾注分隔线，翻了会破坏排版）。
    文档里没有尾注部件时函数内部直接返回 0，不报错。
    counter 通过 **kwargs 透传，无需改动。
    """
    return translate_extra_part_paragraphs(
        docx_file, ENDNOTES_PART, "w:endnote", ("-1", "0"), "尾注",
        client, system_prompt, glossary, direction, **kwargs)
# =====================================================================
#                     第四部分：防呆检测（不处理的区域）
# =====================================================================
def _iter_region_roots(doc):
    """
    产出文档里「可能藏文字」的那几个 XML 部件的解析结果：
    正文 / 页眉 / 页脚 / 脚注 / 尾注。
    为什么不直接用 doc.paragraphs：文本框、艺术字、批注都不在正文段落树里，
    doc.paragraphs 根本看不见它们，只能直接看部件本身。
    """
    package = doc.part.package
    parts = package.iter_parts() if hasattr(package, "iter_parts") else package.parts
    for part in parts:
        part_name = str(part.partname)
        if not part_name.startswith("/word/"):
            continue
        file_name = part_name.rsplit("/", 1)[-1]
        if not (file_name in UNSUPPORTED_REGION_PARTS
                or file_name.startswith("header")
                or file_name.startswith("footer")):
            continue
        try:
            try:
                xml_bytes = part.blob
            except AttributeError:
                xml_bytes = part._element.xml.encode("utf-8")   # 兜底：直接序列化元素
            yield etree.fromstring(xml_bytes)
        except Exception:
            continue        # 某个部件解析不了就跳过，检测本身绝不能把主流程搞崩
def _count_elements_by_local_name(root, local_names):
    """
    统计整棵树里指定「本地名」的元素个数。
    用本地名（localname）而不是 "w:xxx"：命名空间前缀是各软件自己定的，
    Word 用 w:，别的软件可能用 ns0:，按本地名比才认得出，
    而且不会把 w:ins 和 w:instrText 搞混。
    """
    count = 0
    for element in root.iter():
        if not isinstance(element.tag, str):        # 注释、处理指令的 tag 不是字符串
            continue
        if etree.QName(element).localname in local_names:
            count += 1
    return count
def _package_has_comments(doc):
    """
    批注正文不在 document.xml 里，而是独立部件 word/comments.xml。
    这里只检查「有没有这个部件」（不解压、不解析），有就说明文档带批注。
    """
    package = doc.part.package
    parts = package.iter_parts() if hasattr(package, "iter_parts") else package.parts
    for part in parts:
        if str(part.partname).startswith("/word/comments"):
            return True
    return False
def detect_unsupported_regions(doc):
    """
    【防呆检测】找出本工具「不处理」的区域，返回中文提示列表；没发现就返回空列表。
    为什么要做这个：文本框、艺术字、批注里的文字不在 doc.paragraphs 里，
    本工具翻不到它们——但客户不知道，拿到的译文少了内容也不会察觉。
    所以宁可吵一句，也不能静默。
    查三类（只读不改，不影响正常的段落翻译）：
      · 文本框 / 艺术字：找 w:txbxContent 节点；
      · 批注：正文里留下的是 w:commentRangeStart / w:commentReference 锚点，
              批注正文则在 word/comments.xml，两边都查；
      · 修订痕迹：w:ins（插入）/ w:del（删除）——带修订的文档翻出来容易半中半英。
    返回示例：["文本框/艺术字（3 处）", "批注（5 个锚点，含批注正文）"]
    """
    textbox_count = 0
    comment_anchor_count = 0
    revision_count = 0
    for root in _iter_region_roots(doc):
        textbox_count += _count_elements_by_local_name(root, TEXTBOX_LOCAL_NAMES)
        comment_anchor_count += _count_elements_by_local_name(root, COMMENT_ANCHOR_LOCAL_NAMES)
        revision_count += _count_elements_by_local_name(root, REVISION_LOCAL_NAMES)
    has_comment_part = _package_has_comments(doc)
    messages = []
    if textbox_count and not ENABLE_TEXTBOX_TRANSLATE:
        messages.append(f"文本框/艺术字（{textbox_count} 处，已关闭翻译）")
    elif textbox_count:
        messages.append(f"文本框/艺术字（{textbox_count} 处，工具会翻译）")
    if has_comment_part and comment_anchor_count:
        messages.append(f"批注（{comment_anchor_count} 个锚点，含批注正文）")
    elif has_comment_part:
        messages.append("批注（文档带批注正文 comments.xml）")
    elif comment_anchor_count:
        messages.append(f"批注锚点（{comment_anchor_count} 处）")
    if revision_count:
        messages.append(f"修订痕迹（{revision_count} 处插入/删除标记）")
    return messages
# =====================================================================
#                  第五部分：输出路径、本地质检、报告
# =====================================================================
# ★ 新增：断点续跑用的进度文件（放在原文件旁边，和输出文件的 (1)(2) 序号无关）
PROGRESS_FILE_SUFFIX = "_翻译进度.json"
def progress_file_path(word_file, direction=DIRECTION_EN2ZH):
    """
    进度文件的路径：与待翻译文档同目录，文件名带方向后缀。
    刻意不含输出文件的 (1)(2) 序号——否则重启一次就找不到上次的进度了。
    """
    suffix = "_英文译文" if direction == DIRECTION_ZH2EN else "_中文译文"
    folder = os.path.dirname(os.path.abspath(word_file))
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    return os.path.join(folder, f"{base_name}{suffix}{PROGRESS_FILE_SUFFIX}")
def file_signature(path):
    """
    给文件打个「大小 + 修改时间」的指纹。
    断点续跑前用它确认：客户这期间有没有改过原文件。
    改过就不敢续跑了——段落早就不对号了，续下去会张冠李戴。
    """
    try:
        stat_info = os.stat(path)
        return {"size": stat_info.st_size, "mtime": int(stat_info.st_mtime)}
    except Exception:
        return {}
def load_progress(word_file, direction=DIRECTION_EN2ZH):
    """读取进度文件；不存在或读坏了都返回 None（当成没有断点）。"""
    path = progress_file_path(word_file, direction)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as progress_file:
            data = json.load(progress_file)
        return data if isinstance(data, dict) else None
    except Exception:
        return None
def save_progress(word_file, direction, payload):
    """
    写进度文件（先写临时文件再原子替换，避免中途断电留下半个坏文件）。
    存不下也不能影响翻译本身，所以整段都是 try-except。
    """
    path = progress_file_path(word_file, direction)
    try:
        payload = dict(payload)
        payload["saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        temp_path = path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as progress_file:
            json.dump(payload, progress_file, ensure_ascii=False, indent=2)
        os.replace(temp_path, path)
    except Exception:
        pass
def remove_progress(word_file, direction=DIRECTION_EN2ZH):
    """正常跑完后删掉进度文件（留着会让下次误以为是未完成任务）。"""
    try:
        path = progress_file_path(word_file, direction)
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass
def inspect_progress(word_file, direction, glossary_file, provider=PROVIDER_ZHIPU):
    """
    ★ 新增：判断进度文件能不能用来续跑，返回 (状态, 进度数据, 原因)。
        ("none",   None, "")      没有进度文件 / 功能关掉了 → 正常从零开始
        ("usable", data, "")      可以续跑
        ("stale",  data, 原因)    文件在，但和当前情况对不上 → 不续跑，重新开始
    为什么必须校验「对不上」：续跑靠的是「第几段」这个序号，
    客户要是换了个文档、改了原文件、换了方向或换了术语库，
    序号就对不上了，续下去会把 A 段译文写到 B 段上——比重新翻更糟。
    """
    if not ENABLE_RESUME:
        return "none", None, ""
    data = load_progress(word_file, direction)
    if not data:
        return "none", None, ""
    if data.get("stage") == "done":
        return "stale", data, "上次的任务其实已经跑完了（进度文件只是没来得及删除）"
    if data.get("source") != word_file or data.get("direction") != direction:
        return "stale", data, "进度文件记录的原文件或翻译方向与本次不一致"
    # ★ 新增：换了服务商（用的模型不同）也不许续跑。旧进度文件里没有 provider 字段时
    #   按「智谱」算——升级前的进度文件都是智谱跑的，这样不会误伤。
    if (data.get("provider") or PROVIDER_ZHIPU) != provider:
        return "stale", data, "进度文件记录的服务商与本次选择的不一致（换服务商请重新开始）"
    if data.get("file_signature") != file_signature(word_file):
        return "stale", data, "原文件在上次翻译之后被修改过（大小或修改时间变了）"
    if data.get("glossary_file") != glossary_file:
        return "stale", data, "本次的术语库和上次不一致（续跑会导致前后半段术语不统一）"
    if not data.get("output_path") or not os.path.exists(data["output_path"]):
        return "stale", data, "上次生成的结果文件已经不在了"
    return "usable", data, ""
def format_duration(seconds):
    """把秒数写成人话：45 秒 / 12 分 30 秒 / 1 小时 5 分。"""
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds} 秒"
    minutes, second = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes} 分 {second} 秒"
    hours, minutes = divmod(minutes, 60)
    return f"{hours} 小时 {minutes} 分"
def build_output_path(word_file, direction=DIRECTION_EN2ZH):
    """
    根据输入文件和翻译方向生成输出路径：与原文件同目录，文件名加方向后缀。
    若已存在同名结果，自动加 (1)(2) 序号，避免覆盖客户之前的成果。
    """
    suffix = "_英文译文" if direction == DIRECTION_ZH2EN else "_中文译文"
    folder = os.path.dirname(os.path.abspath(word_file))
    base_name = os.path.splitext(os.path.basename(word_file))[0]
    candidate = os.path.join(folder, f"{base_name}{suffix}.docx")
    index = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{base_name}{suffix}({index}).docx")
        index += 1
    return candidate
def count_target_characters(text, direction=DIRECTION_EN2ZH):
    """数一数文本里有多少个「目标语言字符」（中译英数英文字母，英译中数汉字）。"""
    if not text:
        return 0
    if direction == DIRECTION_ZH2EN:
        return len(LATIN_PATTERN.findall(text))
    return len(CJK_PATTERN.findall(text))
def target_language_ratio(text, direction=DIRECTION_EN2ZH):
    """
    统计文本里「目标语言字符」的占比。
    用来抓「方向选反了」「整段没翻译」这类问题：正常译文这个比例应该很高。
    """
    if not text:
        return 1.0
    return count_target_characters(text, direction) / len(text)
def _looks_untranslated(original_text, translated_text):
    """
    判断「译文和原文一模一样」——这是 API 失败回退原文、或模型原样吐回原文的痕迹。
    只做归一化后比对（去空白、忽略大小写），不做翻译质量判断。
    注意：纯专有名词/编号段落（例如 “ISO 9001”“A-01”）本来就可能完全相同，会有少量误报。
    """
    def normalize(text):
        return re.sub(r"\s+", "", text).lower()
    return normalize(original_text) == normalize(translated_text)
def _language_baseline(direction=DIRECTION_EN2ZH):
    """
    「目标语言字符占比」的满分基准，按方向取不同的值：
    · 英译中：目标是中文，汉字几乎占满整段 → 0.98；
    · 中译英：目标是英文，字母只占约 80%（空格约 15%、标点约 5%），
      拿 0.98 当基准会把完全正常的英文译文判成「不达标」（实测踩到过）→ 0.80。
    """
    return 0.80 if direction == DIRECTION_ZH2EN else 0.98
def calculate_quality_score(stats, direction=DIRECTION_EN2ZH):
    """
    本地质检评分（零 API 费用）：按维度算分，缺数据的维度不计入并按权重归一化。
    返回 {"score": 0~100, "verdict": 通过/需抽查/不通过, "red_lines": [...], "lines": [...]}。
    为什么要有评分：报告只统计不评分时，客户拿到一堆数字也不知道「能不能交」。
    这个分数不是翻译文笔分（那要靠 LLM 裁判 + 黄金集回归），而是「交付风险分」：
    漏翻、回退原文、术语不达标、方向选反，这些问题它一定能抓到。
    """
    dimension_scores = {}
    detail_lines = []
    # ★ 本次修改：维度 1 完整性 —— 分母改成「真正动手处理过的段落数」。
    # 为什么这样改（审查意见）：
    #   ① 附加部件（文本框/脚注/批注）的写入现在也会累加到 translated_count 里，
    #      所以分母自然包含它们，不再只统计正文；
    #   ② 原口径用的「待处理段落总数」里含空段和「含域代码被跳过」的段落，
    #      那些本来就不该翻，拿它们当分母会把分数虚高。
    handled_total = stats["translated_count"] + stats["write_failed"]
    abnormal = stats["api_failed"] + stats["write_failed"] + stats["same_count"]
    dimension_scores["完整性"] = max(0.0, 1 - abnormal / max(handled_total, 1))
    detail_lines.append(
        f"完整性：实际处理 {handled_total} 段"
        f"（正文 {stats.get('body_count', 0)} 段 + 文本框/脚注/批注 "
        f"{stats.get('attachment_count', 0)} 段），"
        f"其中 API 失败回退 {stats['api_failed']} 段、结构特殊未写入 {stats['write_failed']} 段、"
        f"疑似未翻译 {stats['same_count']} 段")
    # 维度 2：术语合规率（没术语库就不计这一维）
    term_rate = None
    if stats["term_paragraphs"]:
        term_rate = (stats["term_paragraphs"] - stats["term_bad"]) / stats["term_paragraphs"]
        dimension_scores["术语"] = term_rate
        detail_lines.append(f"术语：含术语段落 {stats['term_paragraphs']} 个"
                            f"（含文本框/脚注/批注），"
                            f"最终仍未达标 {stats['term_bad']} 个（合规率 {term_rate:.1%}）")
    elif stats.get("glossary_selected"):
        detail_lines.append("术语：⚠️ 选定了术语库但读到 0 条术语，本次没有任何术语约束"
                            "（该维度不计分）")
    else:
        detail_lines.append("术语：本次未使用术语库，该维度不计分")
    # 维度 3：语言一致性（正文译文的目标语言字符占比，基准按方向取）
    baseline = _language_baseline(direction)
    language_rate = (min(1.0, stats["target_char_ratio"] / baseline)
                     if stats["translated_count"] else 0.0)
    dimension_scores["语言一致性"] = language_rate
    detail_lines.append(f"语言一致性：正文译文里目标语言字符占比 "
                        f"{stats['target_char_ratio']:.1%}"
                        f"（{direction}，正常约 {baseline:.0%} 以上）")
    # 维度 4：数字一致
    if stats["number_paragraphs"]:
        number_rate = (stats["number_paragraphs"] - stats["number_bad"]) / stats["number_paragraphs"]
        dimension_scores["数字一致"] = number_rate
        detail_lines.append(f"数字：含多位数字段落 {stats['number_paragraphs']} 个"
                            f"（含文本框/脚注/批注），"
                            f"最终数字对不上 {stats['number_bad']} 个")
    else:
        detail_lines.append("数字：原文里没有多位数字，该维度不计分")
    # 维度 5：混排格式还原（只有正文会用到标记，附加部件不涉及）
    if stats["markup_paragraphs"]:
        format_rate = (stats["markup_paragraphs"] - stats["markup_fallback"]) / stats["markup_paragraphs"]
        dimension_scores["格式还原"] = format_rate
        detail_lines.append(f"混排格式：需要标记还原的段落 {stats['markup_paragraphs']} 个，"
                            f"其中退回整段单格式 {stats['markup_fallback']} 个")
    else:
        detail_lines.append("混排格式：本次没有段内混排格式的段落，该维度不计分")
    # 加权归一化：只算有数据的维度
    total_weight = sum(WEIGHTS[name] for name in dimension_scores)
    score = 0.0
    if total_weight:
        score = sum(WEIGHTS[name] * rate for name, rate in dimension_scores.items()) / total_weight * 100
    score = round(score, 1)
    # 硬红线：命中直接判「不通过」（分数只作参考）
    red_lines = []
    if term_rate is not None and term_rate < RED_TERM_RATE:
        red_lines.append(f"术语合规率 {term_rate:.1%} 低于红线 {RED_TERM_RATE:.0%}")
    if stats["translated_count"] and stats["target_char_ratio"] < RED_TARGET_RATIO:
        red_lines.append(f"目标语言字符占比 {stats['target_char_ratio']:.1%} 低于红线 "
                         f"{RED_TARGET_RATIO:.0%}，可能是翻译方向选反了，或大量段落没翻")
    if stats["api_failed"]:
        red_lines.append(f"有 {stats['api_failed']} 段因 API 失败回退成了原文")
    if red_lines:
        verdict = "不通过"
    elif score >= PASS_SCORE:
        verdict = "通过"
    elif score >= WARN_SCORE:
        verdict = "需抽查"
    else:
        verdict = "不通过"
    return {"score": score, "verdict": verdict, "red_lines": red_lines, "lines": detail_lines}
def _verify_real_structure(src_path, dst_path):
    """
    真实回读原文和译文，比对段落数、表格数、图片数。
    这是报告中「结构未变」声明的唯一事实依据。
    ★ 为什么要真的把两个文件打开数一遍？
      报告里原来那句「译文文档结构未变」，依据只是「在同文件副本上就地替换文字，
      段落数天然一致」——那是**推理**，不是**证据**。推理会错（某段写入失败、
      附加部件被改坏、文件被别的东西覆盖都可能发生），而这份报告是要连译文一起
      交给客户的。所以现在改成：真的把磁盘上的原文和译文各打开一次，把段落数、
      表格数、图片数各数一遍，数字对得上才敢写「一致」，对不上就如实报出来。
    返回：{"paragraphs": (原文数, 译文数), "tables": (原文数, 译文数),
           "images": (原文数, 译文数)}
          读文件失败时返回 {"error": "错误信息"}，报告里会照实写「核对失败」，
          而不是假装一切正常。
    """
    # ★ 本次修改（修复1）：先把 zip 部件级比对做掉，它跟 python-docx 那三项互相独立，
    #   就算 Document() 打不开文件，部件级结论照样能报出来。
    parts_result = _compare_package_parts(src_path, dst_path)
    try:
        doc_src = Document(src_path)
        doc_dst = Document(dst_path)
        return {
            "paragraphs": (len(doc_src.paragraphs), len(doc_dst.paragraphs)),
            "tables": (len(doc_src.tables), len(doc_dst.tables)),
            "images": (len(doc_src.inline_shapes), len(doc_dst.inline_shapes)),
            "parts": parts_result,
        }
    except Exception as e:
        return {"error": str(e), "parts": parts_result}
# ★ 本次修改（修复1）：结构核对要看到「文字部件」清单。
# 为什么必须回到 zip 包这一层：python-docx 的 doc.paragraphs / tables /
# inline_shapes 只看得到正文（连正文表格里的段落都不全），脚注、尾注、批注、
# 页眉页脚全都看不见。只比正文那三项的话，脚注/批注整个丢了（或被写坏），
# 报告照样写「结构一致」——那是假绿灯，比没有校验更危险。
STRUCTURE_EXACT_PART_NAMES = ("word/document.xml", "word/footnotes.xml",
                              "word/endnotes.xml", "word/comments.xml")
STRUCTURE_PART_PREFIXES = ("word/header", "word/footer")
# 这些部件要逐个数 w:p 段落数（正文的段落数由上面 python-docx 那行负责，
# 两边口径不一样，混在一起报会让人以为是两组互相矛盾的数字）。
STRUCTURE_COUNT_EXACT_NAMES = ("word/footnotes.xml", "word/endnotes.xml",
                               "word/comments.xml")
def _structure_part_label(part_name):
    """把 zip 里的部件名翻成人看得懂的名字（页眉1、页脚2、脚注…）。"""
    if part_name.startswith("word/header"):
        return "页眉" + part_name[len("word/header"):-len(".xml")]
    if part_name.startswith("word/footer"):
        return "页脚" + part_name[len("word/footer"):-len(".xml")]
    return {
        "word/document.xml": "正文",
        "word/footnotes.xml": "脚注",
        "word/endnotes.xml": "尾注",
        "word/comments.xml": "批注",
    }.get(part_name, part_name)
def _compare_package_parts(src_path, dst_path):
    """
    ★ 本次修改（修复1）：在 zip 包这一层，比对原文和译文里的「文字部件」。
    查两件事：
      1) 部件在不在：正文、脚注、尾注、批注、每一个页眉/页脚。
         原文有、译文没有 = 真丢了（这正是「写盘时丢部件」会留下的痕迹）。
      2) 部件里的 w:p 段落数一不一样。本工具只换文字、不动结构，
         所以脚注/批注/页眉页脚的段落数必须和原文一模一样。
    返回：{"checked": True, "missing_parts": [...], "extra_parts": [...],
           "part_paragraphs": {"word/footnotes.xml": (原文段数, 译文段数), ...}}
          整个 zip 打不开时返回 {"checked": False, "error": "错误信息"}。
    """
    result = {
        "checked": False,
        "missing_parts": [],
        "other_missing_parts": [],
        "extra_parts": [],
        "part_paragraphs": {},
    }
    try:
        with zipfile.ZipFile(src_path, "r") as zip_src, \
                zipfile.ZipFile(dst_path, "r") as zip_dst:
            names_src = set(zip_src.namelist())
            names_dst = set(zip_dst.namelist())
            # 1) 关键部件在不在
            key_parts = sorted(
                name for name in names_src
                if name in STRUCTURE_EXACT_PART_NAMES
                or name.startswith(STRUCTURE_PART_PREFIXES))
            result["missing_parts"] = sorted(set(key_parts) - names_dst)
            # 其他部件（图片、样式、主题…）也一起记下来，但只当提示用：
            # 有些「孤立部件」（在 zip 里但没被任何关系引用）本来就会被 doc.save() 丢掉。
            result["other_missing_parts"] = sorted(
                (names_src - set(key_parts)) - names_dst)
            result["extra_parts"] = sorted(names_dst - names_src)
            # 2) 部件里的段落数对不对（单个部件解析失败只跳过它，不影响别的）
            for part_name in key_parts:
                if part_name not in names_dst:
                    continue
                if not (part_name in STRUCTURE_COUNT_EXACT_NAMES
                        or part_name.startswith(STRUCTURE_PART_PREFIXES)):
                    continue
                try:
                    src_root = etree.fromstring(zip_src.read(part_name))
                    dst_root = etree.fromstring(zip_dst.read(part_name))
                except Exception:
                    continue
                result["part_paragraphs"][part_name] = (
                    _count_elements_by_local_name(src_root, ("p",)),
                    _count_elements_by_local_name(dst_root, ("p",)),
                )
            result["checked"] = True
    except Exception as e:
        result["error"] = str(e)
    return result
def format_structure_parts_lines(parts):
    """
    ★ 本次修改（修复1）：把部件级比对结果拼成报告里直接可读的几行。
    返回 list[str]，调用方 lines += 进去即可。
    """
    if not parts:
        return ["结构一致性校验（附加部件）：⚠️ 未做比对"]
    if parts.get("error"):
        return [f"结构一致性校验（附加部件）：⚠️ 无法比对（{parts['error']}）"]
    lines = []
    detail = []
    for part_name, (count_src, count_dst) in sorted(parts["part_paragraphs"].items()):
        label = _structure_part_label(part_name)
        mark = "" if count_src == count_dst else " ⚠️"
        detail.append(f"{label} {count_src} → {count_dst} 段{mark}")
    if detail:
        lines.append("结构一致性校验（附加部件段落数）：" + "，".join(detail))
    missing = parts.get("missing_parts") or []
    if missing:
        lines.append("结构一致性校验（部件缺失）：⚠️ 译文里少了 "
                     f"{len(missing)} 个部件 —— "
                     + "、".join(_structure_part_label(name) for name in missing))
    else:
        lines.append("结构一致性校验（部件缺失）：✅ 原文里的正文/脚注/尾注/批注/"
                     "页眉页脚部件，译文里一个都没少")
    # 非文字部件（图片、样式、主题、设置…）理论上也不该少，但它们不承载译文，
    # 而且有些孤立部件本来就会被 Word/python-docx 顺手丢掉，所以只做提示，不判 ❗。
    other_missing = parts.get("other_missing_parts") or []
    if other_missing:
        lines.append(f"（提示）另少了 {len(other_missing)} 个非文字部件"
                     f"（图片/样式/主题等）：" + "、".join(other_missing[:8]))
    extra = parts.get("extra_parts") or []
    if extra:
        lines.append(f"（提示）译文比原文多出 {len(extra)} 个部件："
                     + "、".join(extra[:8]))
    return lines
def write_translation_report(output_path, info):
    """
    在译文旁边生成《翻译报告.txt》：统计 + 质检评分 + 需要人工处理的段落清单。
    用 utf-8-sig 编码，双击用记事本打开不会乱码。返回报告文件路径。
    报告是这个工具「能交付」的关键一环：翻译完客户最关心两件事——
    「有没有没翻到的地方」「这份东西能不能直接交」，答案都必须落成文件，
    而不是滚过去就没了的日志。
    """
    report_path = os.path.splitext(output_path)[0] + "_翻译报告.txt"
    quality = info.get("quality") or {}
    lines = [
        "Word 文档翻译报告",
        "=" * 62,
        f"生成时间：{time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"翻译方向：{info['direction']}",
        f"服务商  ：{info.get('provider', PROVIDER_ZHIPU)}",
        f"使用模型：{info['model']}",
        f"原文档  ：{info['source_file']}",
        f"结果文档：{output_path}",
        f"术语库  ：{info['glossary_file'] or '（未使用）'}",
        f"术语条数：{info['glossary_count']}",
    ]
    if info.get("glossary_warning"):
        lines += ["", "⚠️ 术语库提醒：", "   " + info["glossary_warning"]]
    if info.get("unsupported_regions"):
        lines += ["", "⚠️ 检测到的特殊区域（见下方处理说明）：",
                  "   " + "、".join(info["unsupported_regions"])]
    # —— 质检评分 ——
    if quality:
        lines += [
            "",
            "-" * 62,
            "【质检评分（本地复算，不消耗任何 API 费用）】",
            f"评分：{quality['score']} / 100　　判定：{quality['verdict']}",
        ]
        for detail_line in quality["lines"]:
            lines.append("· " + detail_line)
        if quality["red_lines"]:
            lines.append("")
            lines.append("❗ 命中硬红线（直接判不通过）：")
            for red_line in quality["red_lines"]:
                lines.append("   - " + red_line)
    # —— 统计 ——
    lines += [
        "",
        "-" * 62,
        "【统计】",
        f"待处理段落总数（含表格/页眉页脚）：{info['total_targets']}",
        f"实际翻译并写入的段落：{info['translated_count']}",
        f"跳过空行（未消耗 Token）：{info['skipped_blank']}",
        f"跳过含超链接/域代码的段落：{len(info['skipped_special'])}",
        f"文本框/艺术字翻译段数：{info.get('textbox_count', 0)}",
        f"脚注翻译段数：{info['footnote_count']}",
        f"批注翻译段数：{info.get('comment_count', 0)}",
        # ★ 新增：尾注（word/endnotes.xml，和脚注同一套部件处理）
        f"尾注翻译段数：{info.get('endnote_count', 0)}",
        # ★ 本次修改：新增 3 行，把「附加部件」的处理量与异常量单独列出来
        f"附加部件已处理段数（文本框+脚注+批注+尾注）：{info.get('attachment_count', 0)}",
        f"附加部件异常段数：{info.get('attachment_failed', 0)}",
        f"实际处理段数（进质检分母）：{info.get('body_count', 0)} 正文 + "
        f"{info.get('attachment_count', 0)} 附加部件",
        f"定向重译次数：{info['retry_total']}",
        f"需要人工处理的段落：{len(info['problems'])}",
        # ★ 本次修改：原来这里写「译文文档结构未变（同文件就地替换，段落数天然一致）」，
        #   那只是推理、没有证据。结构到底有没有变，改由下面那行「结构一致性校验」
        #   真实回读原文/译文数出来的数字说话，这里不下结论。
        f"是否中途停止：{'是（结果是半成品，只含已翻译的部分）' if info['stopped'] else '否'}",
    ]
    # ★ 本次修改：真实回读比对，不再写虚假的「结构未变」
    struct_check = _verify_real_structure(info['source_file'], output_path)
    if "error" in struct_check:
        lines.append(f"结构一致性：⚠️ 无法回读比对（{struct_check['error']}）")
    else:
        p_src, p_dst = struct_check["paragraphs"]
        t_src, t_dst = struct_check["tables"]
        i_src, i_dst = struct_check["images"]
        struct_ok = (p_src == p_dst and t_src == t_dst and i_src == i_dst)
        status = "一致" if struct_ok else "⚠️ 不一致！"
        lines.append(f"结构一致性校验：段落 {p_src} → {p_dst}（{status}），"
                     f"表格 {t_src} → {t_dst}，图片 {i_src} → {i_dst}")
    # ★ 本次修改（修复1）：上面那行只有正文。脚注/批注/页眉页脚在 zip 包里是独立
    #   部件，python-docx 的 paragraphs/tables/inline_shapes 根本看不见它们，
    #   只查正文会出现「脚注整个丢了，报告却写结构一致」的假绿灯。这里再补一层
    #   部件级证据：部件在不在 + 部件里的段落数对不对。
    lines += format_structure_parts_lines(struct_check.get("parts"))
    # —— 需要人工处理的段落 ——
    lines += ["", "-" * 62, "【需要人工处理的段落】"]
    if not info["problems"] and not info["skipped_special"]:
        lines.append("（无，全部段落都已正常翻译）")
    else:
        index = 1
        for number, location in info["skipped_special"]:
            lines.append(f"{index}) 第 {number} 个待翻译段落（{location}）："
                         f"含超链接或域代码，为避免翻花页码/目录已跳过，请人工处理")
            index += 1
        for problem in info["problems"]:
            number = problem["number"]
            prefix = f"第 {number} 个待翻译段落" if number else ""
            lines.append(f"{index}) {prefix}（{problem['location']}）："
                         f"{problem['kind']} —— {problem['detail']}")
            index += 1
    # —— 疑似需要人工确认的段落（本地质检发现的可疑项，不一定是错）——
    manual_review = info.get("manual_review") or []
    lines += ["", "-" * 62, "【疑似需要人工确认的段落（本地质检提示，仅供参考）】"]
    if manual_review:
        for index, item in enumerate(manual_review, start=1):
            lines.append(f"{index}) {item}")
        if info.get("manual_review_overflow"):
            lines.append(f"（还有更多疑似项未列出，共 {info['manual_review_overflow']} 条）")
    else:
        lines.append("（无）")
    lines += [
        "",
        "-" * 62,
        "【说明】",
        "· 本工具在「原文件的完整副本」上只替换文字，粗体、字号、标题样式、编号、表格、",
        "  图片、页面设置都会保留；客户的原文件自始至终没有被修改。",
        "· 段内混排格式：只还原「相对段落主体格式多出来的」加粗/斜体/下划线；",
        "  主体本身加粗、个别词不加粗这种反向情况不还原。",
        "· 含超链接/域代码的段落跳过不翻；文本框内部不做混排格式还原。",
        "· 跳过空行不消耗任何 Token。",
        "· 质检评分是「交付风险分」，用来抓漏翻和术语不达标；翻译文笔的好坏",
        "  需要另用质检脚本（LLM 裁判 + 黄金集回归）评定。",
        "· 结构一致性校验：报告里那几行数字，是译文写完后重新打开原文和译文各数一遍",
        "  得到的真实数字（正文的段落/表格/图片 + 脚注/尾注/批注/页眉页脚的部件与段落数），",
        "  不是「应该不会变」的推理结论。python-docx 看不见脚注/批注/页眉页脚这些部件，",
        "  所以后面那几行是回到 zip 包这一层数出来的。",
    ]
    try:
        with open(report_path, "w", encoding="utf-8-sig") as report_file:
            report_file.write("\n".join(lines) + "\n")
    except Exception:
        pass        # 报告写不出来也不能影响译文文件
    return report_path
# =====================================================================
#                          第六部分：图形界面
# =====================================================================
class TranslatorApp:
    """整个 GUI 应用。所有控件只在主线程操作，耗时任务交给后台线程。"""
    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(800, 600)
        # ---------- 界面状态变量 ----------
        self.api_key_var        = tk.StringVar()
        self.api_key_label_var  = tk.StringVar()        # ★ 新增：Key 标签随服务商变
        self.api_key_hint_var   = tk.StringVar()        # ★ 新增：Key 留空提示随服务商变
        self.provider_var       = tk.StringVar(value=PROVIDER_ZHIPU)   # ★ 新增：服务商
        self.direction_var      = tk.StringVar(value=DIRECTION_EN2ZH)
        self.glossary_path      = tk.StringVar()
        self.word_path          = tk.StringVar()
        self.use_glossary       = tk.BooleanVar(value=True)
        self.glossary_label_var = tk.StringVar()
        self.output_hint        = tk.StringVar(value="输出文件：选择 Word 后自动生成（与原文件同目录）")
        # ---------- 运行状态 ----------
        self.log_queue = queue.Queue()          # 后台线程 → 主线程的日志队列
        self.is_running = False                 # 是否正在翻译（防止重复点击）
        self.stop_event = threading.Event()     # 「停止」标志：后台线程每段检查一次
        self.last_output_path = ""              # 最近一次成功生成的结果文件
        self.last_quality_note = ""             # 最近一次的质检评分摘要（弹窗里显示）
        self._build_ui()
        self._refresh_direction_texts()
        self.root.after(100, self._poll_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    # ------------------------------------------------------------------
    # 界面搭建
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(9, weight=1)
        ttk.Label(main, textvariable=self.api_key_label_var).grid(row=0, column=0, sticky="w", pady=(0, 2))
        ttk.Entry(main, textvariable=self.api_key_var, width=60).grid(
            row=0, column=1, columnspan=2, sticky="we", pady=(0, 2))
        ttk.Label(main, textvariable=self.api_key_hint_var,
                  foreground="#777777").grid(row=1, column=1, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(main, text="翻译方向：").grid(row=2, column=0, sticky="w", pady=3)
        direction_frame = ttk.Frame(main)
        direction_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=3)
        ttk.Radiobutton(direction_frame, text="英译中（English → 中文）",
                        variable=self.direction_var, value=DIRECTION_EN2ZH,
                        command=self._on_direction_change).pack(side="left")
        ttk.Radiobutton(direction_frame, text="中译英（中文 → English）",
                        variable=self.direction_var, value=DIRECTION_ZH2EN,
                        command=self._on_direction_change).pack(side="left", padx=(20, 0))
        # ★ 新增：服务商选择（放在「翻译方向」同一行，界面整体布局不用重排）
        ttk.Label(direction_frame, text="　服务商：").pack(side="left", padx=(26, 0))
        ttk.Radiobutton(direction_frame, text=PROVIDER_ZHIPU,
                        variable=self.provider_var, value=PROVIDER_ZHIPU,
                        command=self._on_provider_change).pack(side="left")
        ttk.Radiobutton(direction_frame, text=PROVIDER_DEEPSEEK,
                        variable=self.provider_var, value=PROVIDER_DEEPSEEK,
                        command=self._on_provider_change).pack(side="left", padx=(12, 0))
        ttk.Label(main, textvariable=self.glossary_label_var).grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(main, textvariable=self.glossary_path, state="readonly").grid(
            row=3, column=1, sticky="we", pady=3)
        glossary_btn_frame = ttk.Frame(main)
        glossary_btn_frame.grid(row=3, column=2, sticky="e", padx=(6, 0), pady=3)
        ttk.Button(glossary_btn_frame, text="选择Excel", command=self.choose_glossary).pack(side="left")
        ttk.Button(glossary_btn_frame, text="清除", width=6,
                   command=self.clear_glossary).pack(side="left", padx=(4, 0))
        ttk.Checkbutton(main, text="使用术语库（不勾选则忽略上面选择的 Excel）",
                        variable=self.use_glossary).grid(row=4, column=1, sticky="w", pady=(0, 8))
        ttk.Label(main, text="待翻译 Word：").grid(row=5, column=0, sticky="w", pady=3)
        ttk.Entry(main, textvariable=self.word_path, state="readonly").grid(
            row=5, column=1, sticky="we", pady=3)
        ttk.Button(main, text="选择Word", command=self.choose_word).grid(
            row=5, column=2, sticky="e", padx=(6, 0), pady=3)
        ttk.Label(main, textvariable=self.output_hint, foreground="#777777",
                  wraplength=880).grid(row=6, column=1, columnspan=2, sticky="w", pady=(0, 8))
        button_frame = ttk.Frame(main)
        button_frame.grid(row=7, column=0, columnspan=3, sticky="we", pady=(4, 10))
        self.start_button = ttk.Button(button_frame, text="开始翻译", command=self.start_translate)
        self.start_button.pack(side="left", ipadx=26, ipady=6)
        self.stop_button = ttk.Button(button_frame, text="停止",
                                      command=self.request_stop, state="disabled")
        self.stop_button.pack(side="left", padx=(10, 0))
        self.open_folder_button = ttk.Button(button_frame, text="打开输出文件夹",
                                            command=self.open_output_folder, state="disabled")
        self.open_folder_button.pack(side="left", padx=(10, 0))
        # ★ 新增：进度条 + 「已翻 x/y 段 · 已用 · 预计还需」文案
        progress_frame = ttk.Frame(main)
        progress_frame.grid(row=8, column=0, columnspan=3, sticky="we", pady=(2, 0))
        progress_frame.columnconfigure(0, weight=1)
        ttk.Label(progress_frame, text="运行日志：").grid(row=0, column=0, sticky="w")
        self.progress_text = tk.StringVar(value="")
        ttk.Label(progress_frame, textvariable=self.progress_text,
                  foreground="#00875A").grid(row=0, column=1, sticky="e")
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="we", pady=(3, 0))
        log_frame = ttk.Frame(main)
        log_frame.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(2, 0))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap="word", height=18,
                                bg="#111111", fg="#33FF99", insertbackground="#33FF99",
                                font=("Consolas", 10), state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
    def _on_direction_change(self):
        """切换方向时：更新提示文案 + 重新推算输出文件名 + 写一条日志。"""
        self._refresh_direction_texts()
        self.write_log(f"🔁 已切换翻译方向：{self.direction_var.get()}")
    def _on_provider_change(self):
        """★ 新增：切换服务商时：刷新 Key 标签/提示 + 写一条日志。"""
        self._refresh_direction_texts()
        active_model = (DEEPSEEK_MODEL_NAME if self.provider_var.get() == PROVIDER_DEEPSEEK
                        else MODEL_NAME)
        self.write_log(f"🤖 已切换服务商：{self.provider_var.get()}（模型：{active_model}）")
    def _refresh_direction_texts(self):
        """根据当前方向/服务商刷新界面上的说明文字。"""
        self.glossary_label_var.set("术语库 Excel：")
        # ★ 新增：Key 的标签和环境变量名随服务商切换
        #   （否则客户容易把 DeepSeek 的 Key 贴进「智谱 API Key」框里，报错还看不懂）
        if self.provider_var.get() == PROVIDER_DEEPSEEK:
            self.api_key_label_var.set("DeepSeek API Key：")
            self.api_key_hint_var.set("默认留空即可；留空时会自动读取环境变量 DEEPSEEK_API_KEY"
                                      "（DeepSeek 需要先执行 pip install openai）")
        else:
            self.api_key_label_var.set("智谱 API Key：")
            self.api_key_hint_var.set("默认留空即可；若留空会自动尝试读取环境变量 ZHIPUAI_API_KEY")
        word_file = self.word_path.get().strip()
        if word_file:
            self.output_hint.set("输出文件：" + build_output_path(word_file, self.direction_var.get()))
    # ------------------------------------------------------------------
    # 文件选择
    # ------------------------------------------------------------------
    def choose_glossary(self):
        """选择术语库 Excel（英译中时 A 列英文/B 列中文；中译英时 A 列中文/B 列英文）。"""
        if self.direction_var.get() == DIRECTION_ZH2EN:
            title = "请选择术语库 Excel（A 列中文，B 列英文）"
        else:
            title = "请选择术语库 Excel（A 列英文，B 列中文）"
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.glossary_path.set(path)
            self.use_glossary.set(True)
            self.write_log(f"📗 已选择术语库：{path}")
    def choose_word(self):
        """选择待翻译的 Word 文档。"""
        path = filedialog.askopenfilename(
            title="请选择要翻译的 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.word_path.set(path)
            self.output_hint.set("输出文件：" + build_output_path(path, self.direction_var.get()))
            self.write_log(f"📄 已选择待翻译文档：{path}")
    def clear_glossary(self):
        """清除术语库选择：客户不需要术语约束时用。"""
        self.glossary_path.set("")
        self.use_glossary.set(False)
        self.write_log("🧹 已清除术语库，本次将按普通翻译处理。")
    # ------------------------------------------------------------------
    # 日志
    # ------------------------------------------------------------------
    def write_log(self, message):
        """线程安全的写日志入口（可以被后台线程直接调用）。"""
        self.log_queue.put(("log", str(message)))
    def _append_log_line(self, line):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {line}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
    def _reset_progress(self):
        """★ 新增：任务开始前把进度条清零。"""
        self.progress_bar.configure(value=0)
        self.progress_text.set("")
    def _update_progress(self, done, total, extra_text):
        """
        ★ 新增：刷新进度条和「已翻 x/y 段 · 已用 · 预计还需」文案。
        为什么不用 tqdm：tqdm 是往控制台（stderr）打印进度条的，
        而这个工具是 GUI，客户常用 .pyw 双击启动——根本没有控制台窗口，
        装了 tqdm 也看不见，还多一个要额外 pip install 的依赖。
        ttk.Progressbar 是 tkinter 自带的，零依赖、能嵌在窗口里。
        """
        percent = (done / total * 100) if total else 0
        self.progress_bar.configure(value=max(0.0, min(100.0, percent)))
        self.progress_text.set(extra_text)
    def _poll_log_queue(self):
        """主线程定时任务：把队列里的消息搬到界面上。"""
        try:
            while True:
                item = self.log_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._append_log_line(item[1])
                elif kind == "progress":
                    self._update_progress(item[1], item[2], item[3])
                elif kind == "done":
                    self._on_worker_done(success=item[1], output_path=item[2])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)
    # ------------------------------------------------------------------
    # 开始 / 停止
    # ------------------------------------------------------------------
    def start_translate(self):
        """点「开始翻译」：先做输入校验，再启动后台线程。"""
        if self.is_running:
            messagebox.showinfo("提示", "正在翻译中，请稍候……")
            return
        # ★ 新增：服务商 → 决定用哪家的 Key、读哪个环境变量
        provider = self.provider_var.get()
        env_key_name = "DEEPSEEK_API_KEY" if provider == PROVIDER_DEEPSEEK else "ZHIPUAI_API_KEY"
        api_key = self.api_key_var.get().strip() or os.environ.get(env_key_name, "").strip()
        if not api_key:
            messagebox.showwarning("还差一步",
                                   f"请先填写 {provider} 的 API Key"
                                   f"（或设置环境变量 {env_key_name}）。")
            return
        word_file = self.word_path.get().strip()
        if not word_file or not os.path.exists(word_file):
            messagebox.showwarning("还差一步", "请先点击「选择Word」选择要翻译的文档。")
            return
        direction = self.direction_var.get()
        glossary_file = self.glossary_path.get().strip()
        if not self.use_glossary.get():
            glossary_file = ""
        elif glossary_file and not os.path.exists(glossary_file):
            messagebox.showwarning("文件不存在", f"术语库文件找不到了：\n{glossary_file}")
            return
        # ★ 断点续跑：先看看这个文档 + 这个方向有没有「没跑完的进度文件」
        output_path = build_output_path(word_file, direction)
        resume_data = None
        resume_note = ""
        resume_status, progress_data, resume_reason = inspect_progress(
            word_file, direction, glossary_file, provider)
        if resume_status == "usable":
            already_done = int(progress_data.get("next_number") or 0)
            total_planned = int(progress_data.get("total_targets") or 0)
            output_path = progress_data.get("output_path") or output_path
            if messagebox.askyesno(
                    "检测到上次没跑完的任务",
                    f"上次这个文档翻到第 {already_done}/{total_planned} 段就中断了"
                    f"（{progress_data.get('saved_at', '时间未知')}）。\n\n"
                    f"「是」= 从第 {already_done + 1} 段继续（已经翻好的部分不再花钱）\n"
                    f"「否」= 从头重新开始（会覆盖上次那份半成品）"):
                resume_data = progress_data
            else:
                remove_progress(word_file, direction)    # 不续跑：删进度，覆盖半成品重来
        elif resume_status == "stale":
            remove_progress(word_file, direction)
            resume_note = f"ℹ️ 发现上次的进度文件，但不能续跑（{resume_reason}），本次重新开始。"
        self._clear_log()
        self._reset_progress()
        if resume_note:
            self.write_log(resume_note)
        if resume_data:
            self.write_log(f"▶️ 接着上次跑：从第 "
                           f"{int(resume_data.get('next_number') or 0) + 1} 段继续。")
        self.write_log(f"🚀 任务开始……（{direction} · {provider} · "
                       f"{DEEPSEEK_MODEL_NAME if provider == PROVIDER_DEEPSEEK else MODEL_NAME}）")
        self.is_running = True
        self.stop_event.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.open_folder_button.configure(state="disabled")
        self.last_output_path = ""
        self.last_quality_note = ""
        threading.Thread(
            target=self._translate_worker,
            args=(api_key, word_file, glossary_file, direction, output_path, resume_data, provider),
            daemon=True,
        ).start()
    def request_stop(self):
        """
        点「停止」：只设一个标志位，不强行杀线程。
        后台线程会在处理完当前这段之后停下，并把已经翻译好的部分保存下来，
        这样客户发现选错了方向/术语库时，不必白花已经跑掉的钱。
        """
        if not self.is_running:
            return
        self.stop_event.set()
        self.stop_button.configure(state="disabled")
        self.write_log("⏹ 已收到停止请求，等当前这段翻完就会停下并保存……")
    def _translate_worker(self, api_key, word_file, glossary_file, direction,
                          output_path, resume_data=None, provider=PROVIDER_ZHIPU):
        """
        后台线程里真正干活的函数。这里是整个程序【唯一】写 Word 的地方。
        结构（★ 重写：并发请求，排队写入）：
            ① 准备：术语库、输出文件、防呆检测、收集待翻译段落
            ② 切任务：短而干净的连续段落合并成一批，混排/超长段落各成一个任务
            ③ 多个线程并发调 API（只发请求，绝不碰 Word 文档）
            ④ 回到本线程，严格按段落原始顺序逐个写回 Word（唯一写入者 → 文件不会坏）
            ⑤ 文本框 → doc.save() → 脚注 → 批注（顺序不能变，见各自的注释）
            ⑥ 质检评分 + 写《翻译报告.txt》+ 删/存进度文件
        续跑（resume_data 非空）：用的还是上次那份半成品文档，从上次断的段落继续，
        并且把上次的统计一起接过来，否则最终报告只会统计后半段。
        注意：这里绝对不能直接碰 tkinter 控件，只能 self.write_log(...) 入队。
        """
        try:
            started_at = time.time()
            resume_data = resume_data or {}
            # 1) 术语库 → {原文: 译文} 字典（提示词按需裁剪：见 prune_glossary_for_text）
            glossary = load_glossary(glossary_file, direction, self.write_log)
            system_prompt = build_system_prompt(glossary, direction)
            # 2) 准备输出文件：整份复制原文件，再在副本上「原地替换文字」。
            #    这一步是格式保真的关键——图片、表格、样式、页面设置全在副本里，
            #    我们只改文字，别的东西一个字节都不动；客户的原文件自始至终不碰。
            #    续跑时绝不能重新复制：那样会把上次已经翻好的部分整段冲掉。
            if resume_data:
                doc = Document(output_path)
                self.write_log(f"📂 继续在上次的结果文件上翻译：{output_path}")
            else:
                shutil.copyfile(word_file, output_path)
                doc = Document(output_path)
            # 3) 【防呆检测】文本框/艺术字/批注/修订：这些区域要如实告诉客户
            unsupported_regions = detect_unsupported_regions(doc) if ENABLE_DOC_FEATURE_CHECK else []
            if unsupported_regions:
                self.write_log("⚠️ 检测到文档包含" + "、".join(unsupported_regions)
                               + "，详见翻译报告（这些区域的处理方式见报告说明）。")
            else:
                self.write_log("✅ 防呆检测：未发现文本框/艺术字/批注/修订。")
            # 4) 收集所有待翻译段落（正文 + 表格 + 页眉页脚），每段只算一次
            targets = list(iter_translatable_paragraphs(doc))
            total_targets = len(targets)
            self.write_log(f"✅ 读取成功：共找到 {total_targets} 个待处理段落"
                           f"（含表格、页眉页脚），开始逐段翻译……")
            if total_targets == 0:
                self.write_log("ℹ️ 没有找到可翻译的段落，可能文档是空的或全是图片。")
            # 创建客户端：加请求超时，并关掉 SDK 自带的隐形重试（我们自己的重试会写日志）
            # ★ 本次修改：按界面选的服务商创建客户端（智谱 GLM / DeepSeek）。
            #   两种客户端的接口形状一样，所以下面所有翻译代码一行都不用改。
            try:
                client = build_client(provider, api_key)
            except Exception as client_error:
                self.write_log(f"❗ 创建 {provider} 客户端失败：{client_error}")
                raise
            self.write_log(f"🤖 服务商：{provider}；模型：{_ACTIVE_MODEL[0]}")
            # ★ 质检记账本 —— 正文/文本框/脚注/批注 四条路径共用同一份
            counter = make_stats_counter()
            # 统计（正文循环与附加部件共用一套）
            skipped_blank = 0           # 跳过的空段数
            skipped_special = []        # 含超链接/域代码而跳过的段落 [(段号, 位置), ...]
            manual_review = []          # 疑似需人工确认（本地质检提示，最多留前 N 条）
            manual_review_total = 0     # ★ 疑似需人工确认的「总条数」（不受上限截断）
            markup_paragraphs = 0       # 启用混排标记的段数（格式还原的分母）
            markup_fallback = 0         # 标记还原失败、退回整段单格式的段数
            translated_chars = 0        # 正文译文字符总数（语言占比用）
            target_chars = 0            # 正文译文里目标语言字符数
            textbox_count = 0           # 文本框/艺术字处理段数
            footnote_count = 0          # 脚注处理段数
            comment_count = 0           # 批注处理段数
            endnote_count = 0           # ★ 新增：尾注处理段数
            batch_saved_calls = 0       # 合批省下的请求次数
            # ★ 断点续跑：把上次的统计接过来（否则最终报告只会统计后半段）
            stage = resume_data.get("stage") or "body"      # 见下面 stage_order 说明
            next_number = int(resume_data.get("next_number") or 0)
            last_written_number = next_number
            if resume_data:
                stored_counters = resume_data.get("counters") or {}
                for counter_key in ("translated_count", "body_count", "attachment_count",
                                    "api_failed", "write_failed", "same_count",
                                    "term_paragraphs", "term_bad", "number_paragraphs",
                                    "number_bad", "attachment_failed", "retry_total"):
                    counter[counter_key] = int(stored_counters.get(counter_key) or 0)
                counter["problems"] = list(stored_counters.get("problems") or [])
                stored_extra = resume_data.get("extra_stats") or {}
                skipped_blank = int(stored_extra.get("skipped_blank") or 0)
                skipped_special = list(stored_extra.get("skipped_special") or [])
                manual_review = list(stored_extra.get("manual_review") or [])
                # ★ 总条数也要接着数；老进度文件里没这个字段时，用已有条数兜底
                manual_review_total = int(stored_extra.get("manual_review_total")
                                          or len(manual_review))
                markup_paragraphs = int(stored_extra.get("markup_paragraphs") or 0)
                markup_fallback = int(stored_extra.get("markup_fallback") or 0)
                translated_chars = int(stored_extra.get("translated_chars") or 0)
                target_chars = int(stored_extra.get("target_chars") or 0)
                textbox_count = int(stored_extra.get("textbox_count") or 0)
                footnote_count = int(stored_extra.get("footnote_count") or 0)
                comment_count = int(stored_extra.get("comment_count") or 0)
                endnote_count = int(stored_extra.get("endnote_count") or 0)   # ★ 新增：尾注
                self.write_log(f"↩️ 已接上次统计：上次完成到第 {next_number}/{total_targets} 段，"
                               f"本次从第 {next_number + 1} 段继续。")
            # 阶段顺序：body（正文进行中）→ textbox（正文好了，文本框待做）
            #           → footnote（脚注待做）→ comment（批注待做）→ endnote（尾注待做）
            #           → done（全好了）
            # 每完成一个阶段才往后挪一格，所以续跑时能准确知道该从哪一步接着做。
            stage_order = ["body", "textbox", "footnote", "comment", "endnote", "done"]
            stage_index = stage_order.index(stage) if stage in stage_order else 0
            def build_progress_payload(current_stage, current_number):
                """把「翻到哪儿了 + 统计能不能接上」打包成进度文件的内容。"""
                return {
                    "version": 1,
                    "source": word_file,
                    "file_signature": file_signature(word_file),
                    "direction": direction,
                    # ★ 本次修改：模型名跟着界面选的服务商走；进度文件里也记下服务商，
                    #   换服务商续跑会被 inspect_progress 判为 stale（两家模型不一样）
                    "provider": provider,
                    "model": _ACTIVE_MODEL[0],
                    "glossary_file": glossary_file,
                    "glossary_count": len(glossary),
                    "output_path": output_path,
                    "total_targets": total_targets,
                    "stage": current_stage,
                    "next_number": current_number,
                    "counters": dict(counter),
                    "extra_stats": {
                        "skipped_blank": skipped_blank,
                        "skipped_special": skipped_special,
                        "manual_review": manual_review,
                        "manual_review_total": manual_review_total,
                        "markup_paragraphs": markup_paragraphs,
                        "markup_fallback": markup_fallback,
                        "translated_chars": translated_chars,
                        "target_chars": target_chars,
                        "textbox_count": textbox_count,
                        "footnote_count": footnote_count,
                        "comment_count": comment_count,
                        "endnote_count": endnote_count,     # ★ 新增：尾注
                    },
                }
            # ★ 5) 切任务：连续的短段落合并成一批（一次请求翻多条），
            #      混排段落和超长段落各自单独成一个任务（提示词和格式互不干扰）。
            jobs = []
            current_batch = []
            def flush_batch():
                """把攒着的这一批变成一个任务（可能只有 1 条）。"""
                if not current_batch:
                    return
                if len(current_batch) == 1:
                    label = f"第 {current_batch[0]['number']} 段"
                else:
                    label = (f"第 {current_batch[0]['number']}~"
                             f"{current_batch[-1]['number']} 段")
                jobs.append({"entries": list(current_batch), "label": label})
                current_batch.clear()
            for number, (paragraph, location) in enumerate(targets, start=1):
                if number <= next_number:
                    continue        # ★ 续跑：上次已经翻好的段落直接跳过，不再花钱
                original_text = paragraph.text.strip()
                # —— 空行：直接跳过，不调 API（省 Token），段落本身已经在副本里 ——
                if not original_text:
                    skipped_blank += 1
                    continue
                # —— 含超链接/域代码的段落：整段跳过，避免翻花页码、目录 ——
                if _has_special_content(paragraph):
                    skipped_special.append((number, location))
                    continue
                # —— 段内混排格式：先分析，再决定要不要带标记去翻 ——
                format_plan = _paragraph_format_plan(paragraph)
                if format_plan["has_markup"]:
                    api_source_text = format_plan["marked_source"]     # 发给模型：带标记
                    check_source = format_plan["plain_source"]         # 校验用：不带标记
                    prompt_extra = MARKUP_SYSTEM_RULES
                else:
                    api_source_text = original_text
                    check_source = None                                # None = 普通模式
                    prompt_extra = ""
                entry = {
                    "number": number,
                    "paragraph": paragraph,
                    "location": location,
                    "source": original_text,            # 干净原文（校验、统计、日志用）
                    "format_plan": format_plan,
                    "prompt_extra": prompt_extra,
                    "api_source": api_source_text,      # 真正发给模型的文本
                    "check_source": check_source,
                }
                # 能不能合批：普通段落（没有混排标记）、不算长、加了它不超上限
                batchable = (ENABLE_BATCH and not format_plan["has_markup"]
                             and len(original_text) <= BATCH_MAX_CHARS)
                batch_chars = sum(len(item["source"]) for item in current_batch) + len(original_text)
                if not batchable:
                    flush_batch()                   # 混排/超长：自己一个任务，格式最稳
                    jobs.append({"entries": [entry], "label": f"第 {number} 段"})
                elif current_batch and (len(current_batch) >= BATCH_MAX_ITEMS
                                        or batch_chars > BATCH_MAX_CHARS):
                    flush_batch()                   # 这批满了，另起一批
                    current_batch.append(entry)
                else:
                    current_batch.append(entry)
            flush_batch()
            batch_saved_calls = sum(len(job["entries"]) - 1 for job in jobs
                                    if len(job["entries"]) > 1)
            # ★ 6) 并发请求，排队写入
            worker_count = max(1, API_WORKERS) if ENABLE_CONCURRENCY else 1
            request_window = max(1, worker_count * REQUEST_WINDOW_FACTOR)
            if jobs:
                self.write_log(f"⚙️ 共切成 {len(jobs)} 个翻译任务（覆盖 {total_targets} 个段落）："
                               f"合批省下 {batch_saved_calls} 次请求；"
                               f"{worker_count} 路并发发请求，写回 Word 始终单线程按顺序进行。")
            last_saved_body = counter["body_count"]
            def write_entry(entry, translated_text, diagnostics):
                """
                把一条译文写回 Word。
                ★ 只有本线程（翻译主线程）会调用它 → Word 文件永远只有一个写入者。
                  这正是「并发请求，排队写入」的关键：API 那一步可以 4 路并发，
                  写文件这一步永远串行。
                """
                nonlocal manual_review_total    # ★ 总条数要改外层变量，必须声明 nonlocal
                nonlocal markup_paragraphs, markup_fallback, last_saved_body
                nonlocal translated_chars, target_chars, last_written_number
                paragraph = entry["paragraph"]
                format_plan = entry["format_plan"]
                number, location = entry["number"], entry["location"]
                # ★ 质检分母（含术语 / 含多位数字）按原文统计，只统计真正处理过的段落
                count_source_features(counter, entry["source"], glossary)
                if format_plan["has_markup"]:
                    markup_paragraphs += 1
                # ★ 重译次数 + API 失败/术语/数字问题，统一并入记账本
                record_diagnostics(counter, diagnostics, location, number)
                # —— 写回副本：混排段落走标记还原，普通段落整段单格式 ——
                if format_plan["has_markup"]:
                    write_mode = apply_marked_translation(
                        paragraph, clean_text(translated_text), format_plan)
                else:
                    write_mode = ("整段单格式"
                                  if replace_paragraph_text(paragraph, clean_text(translated_text))
                                  else "无法写入")
                if write_mode == "无法写入":
                    # ★ 统一记账入口：write_failed 会在里面 +1，清单也会记上
                    record_problem(counter, PROBLEM_WRITE_FAILED,
                                   "该段没有可用的文字 run，已跳过以免破坏文档",
                                   location, number)
                    return
                if write_mode == "退化为整段单格式":
                    markup_fallback += 1
                    manual_review_total += 1        # ★ 先记总条数，再决定要不要存进清单
                    if len(manual_review) < MANUAL_REVIEW_LIMIT:
                        manual_review.append(
                            f"第 {number} 个待翻译段落（{location}）：原文段内有混排格式，"
                            "模型未保留格式标记，译文已按整段统一格式写入（内容完整，仅格式简化）")
                # ★ 写入成功，分子分母一起记
                count_written(counter)
                # —— 本地质检统计（零 API 费用）——
                clean_translation = (_strip_markup(translated_text)
                                     if format_plan["has_markup"] else translated_text)
                translated_chars += len(clean_translation)
                target_chars += count_target_characters(clean_translation, direction)
                if _looks_untranslated(entry["source"], clean_translation):
                    counter["same_count"] += 1
                    manual_review_total += 1        # ★ 先记总条数
                    if len(manual_review) < MANUAL_REVIEW_LIMIT:
                        manual_review.append(
                            f"第 {number} 个待翻译段落（{location}）：译文与原文完全相同，"
                            "疑似未翻译或 API 失败回退，请人工确认")
                if (len(clean_translation) >= MIN_LENGTH_FOR_LANG_CHECK
                        and target_language_ratio(clean_translation, direction) < RED_TARGET_RATIO):
                    language_ratio = target_language_ratio(clean_translation, direction)
                    manual_review_total += 1        # ★ 先记总条数
                    if len(manual_review) < MANUAL_REVIEW_LIMIT:
                        manual_review.append(
                            f"第 {number} 个待翻译段落（{location}）：译文里目标语言字符只占 "
                            f"{language_ratio:.0%}，疑似未翻译或方向异常")
                self.write_log(f"  [{number}/{total_targets}] "
                               f"{entry['source'][:24]}... → {clean_translation[:24]}...")
                last_written_number = number
                # ★ 每 PROGRESS_SAVE_EVERY 段存一次断点（续跑的粒度）
                if counter["body_count"] - last_saved_body >= PROGRESS_SAVE_EVERY:
                    last_saved_body = counter["body_count"]
                    save_progress(word_file, direction,
                                  build_progress_payload("body", number))
                # ★ 进度条 + 预计剩余时间（只发一条消息给主线程，不直接碰控件）
                done = counter["body_count"]
                elapsed = time.time() - started_at
                eta_seconds = (elapsed / done * max(0, total_targets - done)) if done else 0.0
                self.log_queue.put(("progress", done, total_targets,
                                    f"已翻 {done}/{total_targets} 段 · "
                                    f"已用 {format_duration(elapsed)} · "
                                    f"预计还需 {format_duration(eta_seconds)}"))
            stopped_early = False
            executor = ThreadPoolExecutor(max_workers=worker_count)
            try:
                pending = deque()
                job_iterator = iter(jobs)
                all_submitted = False
                while True:
                    # 预取：窗口没满就继续提交；点了「停止」就不再提交新任务
                    while (not all_submitted and not self.stop_event.is_set()
                           and len(pending) < request_window):
                        next_job = next(job_iterator, None)
                        if next_job is None:
                            all_submitted = True
                            break
                        pending.append((next_job, executor.submit(
                            self._run_translate_job, next_job, client, glossary, direction)))
                    if not pending:
                        break
                    job, future = pending.popleft()
                    try:
                        job_result = future.result()      # 按提交顺序取结果 = 按段落顺序
                    except Exception as job_error:
                        self.write_log(f"⚠️ 任务「{job['label']}」翻译失败，该段保留原文：{job_error}")
                        job_result = {"entries": [], "batch_diagnostics": None}
                    if job_result.get("batch_diagnostics"):
                        record_diagnostics(counter, job_result["batch_diagnostics"],
                                           f"合批任务（{job['label']}）")
                    if not job_result["entries"]:
                        # 整个任务没拿到结果：每段都记一条「API 失败，已回退原文」
                        for entry in job["entries"]:
                            count_source_features(counter, entry["source"], glossary)
                            record_problem(counter, PROBLEM_API_FAILED,
                                           "该段所在的翻译任务整体失败，保留原文",
                                           entry["location"], entry["number"])
                        continue
                    for entry, translation, diagnostics in job_result["entries"]:
                        write_entry(entry, translation, diagnostics)
                if self.stop_event.is_set():
                    stopped_early = True
                    self.write_log("⏹ 已停止：已经发出去的请求都写完了，不再提交新请求。")
            finally:
                executor.shutdown(wait=True)
            # 说明：写回顺序 = 提交顺序，且提交顺序严格按段落序号，
            #       所以即使 4 路并发，Word 里的段落顺序也不会乱。
            # ★ 7) 正文跑完 → 阶段推进到 textbox；续跑时按进度文件里的阶段接着做
            #      注意：中途被「停止」时绝不能推进阶段，否则续跑会跳过还没翻的正文！
            if stage_index == 0 and not self.stop_event.is_set():
                stage, stage_index = "textbox", 1
                save_progress(word_file, direction,
                              build_progress_payload(stage, total_targets))
            elif stage_index > 0:
                self.write_log("↩️ 正文部分上次已经处理过了，本次直接做附加部件。")
            else:
                self.write_log("⏹ 正文还没翻完就停了：本次先不处理文本框/脚注/批注，"
                               "下次续跑会接着翻剩下的正文。")
            # ★ 8) 文本框 / 艺术字：改内存里的 XML，必须在 doc.save() 之前
            if ENABLE_TEXTBOX_TRANSLATE and stage_index <= 1 and not self.stop_event.is_set():
                self.write_log("🖼 正在处理文本框/艺术字……")
                textbox_count += translate_textboxes(
                    doc, client, system_prompt, glossary, direction,
                    log=self.write_log, should_stop=self.stop_event.is_set,
                    counter=counter)
                stage, stage_index = "footnote", 2
                save_progress(word_file, direction,
                              build_progress_payload(stage, total_targets))
            elif stage_index > 1:
                self.write_log("↩️ 文本框/艺术字上次已经处理过，本次跳过。")
            # 说明：文本框的失败/重译/分母都由 translate_textboxes 内部并入 counter，
            #       外面不用再补一遍（以前是外面手动 append，容易漏）。
            # 9) 保存正文/表格/页眉页脚/文本框的改动
            # 保存前，先备份原始的附加部件
            original_footnotes = None
            original_comments = None
            original_endnotes = None            # ★ 新增：尾注部件
            try:
                with zipfile.ZipFile(word_file, "r") as source_zip:
                    source_names = source_zip.namelist()
                    if FOOTNOTES_PART in source_names:
                        original_footnotes = source_zip.read(FOOTNOTES_PART)
                    if COMMENTS_PART in source_names:
                        original_comments = source_zip.read(COMMENTS_PART)
                    if ENDNOTES_PART in source_names:
                        original_endnotes = source_zip.read(ENDNOTES_PART)
            except Exception as backup_error:
                self.write_log(f"⚠️ 备份附加部件失败（万一被丢弃将无法回填）：{backup_error}")
            doc.save(output_path)
            # ★ 本次修改：保存后，无条件回填附加部件（防止 python-docx 清空内容）
            # 注意：只在还没处理这些部件的阶段回填，避免冲掉已翻译的译文。
            # （唯一的加固：回填本身再包一层 try，万一译文文档正被 Word 打开导致
            #   zip 写不进去，也只记一条日志，不会让整轮翻译在最后一步崩掉。）
            if stage_index <= 2 and original_footnotes is not None:
                try:
                    _replace_zip_entry(output_path, FOOTNOTES_PART, original_footnotes)
                    self.write_log("🧩 附加部件检查：脚注已回填原文备份（防止被 python-docx 清空）")
                except Exception as restore_error:
                    self.write_log(f"⚠️ 脚注部件回填失败（正文译文不受影响）：{restore_error}")
            if stage_index <= 3 and original_comments is not None:
                try:
                    _replace_zip_entry(output_path, COMMENTS_PART, original_comments)
                    self.write_log("🧩 附加部件检查：批注已回填原文备份（防止被 python-docx 清空）")
                except Exception as restore_error:
                    self.write_log(f"⚠️ 批注部件回填失败（正文译文不受影响）：{restore_error}")
            # ★ 新增：尾注部件同样只在其翻译阶段（<=4）之前回填
            if stage_index <= 4 and original_endnotes is not None:
                try:
                    _replace_zip_entry(output_path, ENDNOTES_PART, original_endnotes)
                    self.write_log("🧩 附加部件检查：尾注已回填原文备份（防止被 python-docx 清空）")
                except Exception as restore_error:
                    self.write_log(f"⚠️ 尾注部件回填失败（正文译文不受影响）：{restore_error}")
            # 10) 脚注 + 批注：必须在 doc.save() 之后！
            #     python-docx 保存时会用它内存里那份旧的 XML 覆盖整个包，
            #     先改这两个部件会被冲掉。
            if ENABLE_FOOTNOTE_TRANSLATE and stage_index <= 2 and not self.stop_event.is_set():
                self.write_log("📎 正在处理脚注……")
                footnote_count += translate_footnotes_in_docx(
                    output_path, client, system_prompt, glossary, direction,
                    log=self.write_log, should_stop=self.stop_event.is_set,
                    counter=counter)
                stage, stage_index = "comment", 3
                save_progress(word_file, direction,
                              build_progress_payload(stage, total_targets))
            elif stage_index > 2:
                self.write_log("↩️ 脚注上次已经处理过，本次跳过。")
            if ENABLE_COMMENT_TRANSLATE and stage_index <= 3 and not self.stop_event.is_set():
                self.write_log("💬 正在处理批注……")
                comment_count += translate_comments_in_docx(
                    output_path, client, system_prompt, glossary, direction,
                    log=self.write_log, should_stop=self.stop_event.is_set,
                    counter=counter)
                stage, stage_index = "endnote", 4
                save_progress(word_file, direction,
                              build_progress_payload(stage, total_targets))
            elif stage_index > 3:
                self.write_log("↩️ 批注上次已经处理过，本次跳过。")
            # ★ 新增：尾注（word/endnotes.xml，和脚注是同一套部件处理，只是标签是 w:endnote）
            if ENABLE_ENDNOTE_TRANSLATE and stage_index <= 4 and not self.stop_event.is_set():
                self.write_log("📑 正在处理尾注……")
                endnote_count += translate_endnotes_in_docx(
                    output_path, client, system_prompt, glossary, direction,
                    log=self.write_log, should_stop=self.stop_event.is_set,
                    counter=counter)
                stage, stage_index = "done", 5
                save_progress(word_file, direction,
                              build_progress_payload(stage, total_targets))
            elif stage_index > 4:
                self.write_log("↩️ 尾注上次已经处理过，本次跳过。")
            # 说明：脚注/批注/尾注的失败同样由 translate_extra_part_paragraphs 内部并入 counter。
            # 9) 本地质检评分（零 API 费用）：直接用记账本，正文 + 附加部件一起算
            quality_stats = dict(counter)
            quality_stats.update({
                "total_targets": total_targets,
                "markup_paragraphs": markup_paragraphs,
                "markup_fallback": markup_fallback,
                "target_char_ratio": (target_chars / translated_chars) if translated_chars else 0.0,
                "glossary_selected": bool(glossary_file),
            })
            quality = calculate_quality_score(quality_stats, direction)
            # 10) 写《翻译报告.txt》
            stopped = self.stop_event.is_set()
            report_path = write_translation_report(output_path, {
                "direction": direction,
                "provider": provider,               # ★ 新增：报告里写清这次用的哪家服务商
                "model": _ACTIVE_MODEL[0],          # ★ 本次修改：模型名跟着服务商走
                "source_file": word_file,
                "glossary_file": glossary_file,
                "glossary_count": len(glossary),
                "glossary_warning": (check_glossary_direction(glossary, direction)
                                     or build_empty_glossary_warning(glossary_file, glossary)),
                "unsupported_regions": unsupported_regions,
                "total_targets": total_targets,
                "translated_count": counter["translated_count"],
                "skipped_blank": skipped_blank,
                "skipped_special": skipped_special,
                "textbox_count": textbox_count,
                "footnote_count": footnote_count,
                "comment_count": comment_count,
                "endnote_count": endnote_count,     # ★ 新增：尾注
                "retry_total": counter["retry_total"],
                "problems": counter["problems"],
                "manual_review": manual_review,
                # ★ 本次修改：溢出条数要用「总条数」算；用被截断的清单长度算永远是 0
                "manual_review_overflow": max(0, manual_review_total - len(manual_review)),
                # ★ 本次修改：附加部件的处理量与异常量（报告里单独列一行）
                "attachment_count": counter["attachment_count"],
                "attachment_failed": counter["attachment_failed"],
                "body_count": counter["body_count"],
                "quality": quality,
                "stopped": stopped,
            })
            # ★ 12) 收尾：全部完成就删掉进度文件；中途停止则留下进度，下次可以接着跑
            if stopped:
                if stage == "body":
                    saved_payload = build_progress_payload("body", last_written_number)
                else:
                    saved_payload = build_progress_payload(stage, total_targets)
                save_progress(word_file, direction, saved_payload)
                self.write_log(f"💾 断点已保存：下次选择同一个文档时会问你要不要从第 "
                               f"{last_written_number} 段之后继续。")
            else:
                remove_progress(word_file, direction)
                self.log_queue.put(("progress", total_targets, total_targets,
                                    f"全部完成 · 共 {format_duration(time.time() - started_at)}"))
            # 13) 收尾日志
            self.write_log("⏹ 已停止，已翻译的部分已保存。" if stopped else "🎉 全部完成！")
            self.write_log(f"   翻译方向：{direction}；服务商：{provider}；模型：{_ACTIVE_MODEL[0]}")
            self.write_log(f"   正文 {counter['body_count']} 段；空行跳过 {skipped_blank} 段；"
                           f"特殊段落跳过 {len(skipped_special)} 段；"
                           f"文本框 {textbox_count} 段；脚注 {footnote_count} 段；"
                           f"批注 {comment_count} 段；尾注 {endnote_count} 段。")
            self.write_log(f"   合批省下 {batch_saved_calls} 次请求；并发 {worker_count} 路发请求；"
                           f"进度文件：{'已保留（下次可续跑）' if stopped else '已删除（任务完成）'}。")
            self.write_log(f"   定向重译 {counter['retry_total']} 次；"
                           f"需要人工处理的段落 {len(counter['problems'])} 处"
                           f"（其中文本框/脚注/批注 {counter['attachment_failed']} 处）。")
            self.write_log(f"   🧪 质检评分：{quality['score']}/100（{quality['verdict']}）")
            if quality["red_lines"]:
                for red_line in quality["red_lines"]:
                    self.write_log(f"      ❗ {red_line}")
            self.write_log(f"   结果文件：{output_path}")
            self.write_log(f"   翻译报告：{report_path}")
            # ★ 本次修改：结构核对结论同步打到日志（报告里是白纸黑字的证据，日志里是即时提醒）
            structure_check = _verify_real_structure(word_file, output_path)
            if structure_check.get("error"):
                self.write_log(f"   ⚠️ 结构核对失败：{structure_check['error']}"
                               f"（请人工打开两个文档确认后再交付）")
            else:
                # ⚠️ 注意：不能再拿 structure_check.values() 去解包了——现在多了一个
                #    "parts" 键（部件级比对结果，本身是个 dict），解包会直接抛异常。
                counts_ok = all(
                    count_src == count_dst
                    for count_src, count_dst in (structure_check["paragraphs"],
                                                 structure_check["tables"],
                                                 structure_check["images"]))
                # ★ 本次修改（修复1）：正文之外的部件（脚注/批注/页眉页脚）单独判一次。
                #   True=都在且段落数一致；False=丢了/段落数对不上/压根没比对成；
                #   None=没做成比对，也不能当通过。
                parts_check = structure_check.get("parts") or {}
                if parts_check.get("error"):
                    parts_ok = False
                elif not parts_check.get("checked"):
                    parts_ok = None
                else:
                    parts_ok = (not parts_check.get("missing_parts")) and all(
                        count_src == count_dst
                        for count_src, count_dst in parts_check["part_paragraphs"].values())
                if counts_ok and parts_ok is True:
                    if parts_check.get("part_paragraphs"):
                        self.write_log("   ✅ 结构核对：真实回读原文与译文，段落数/表格数/图片数"
                                       "一致；脚注/批注/页眉页脚部件也都在、段落数一致。")
                    else:
                        # 部件级结果为空 = 原文档里本来就没有脚注/批注/页眉页脚，没什么可丢的
                        self.write_log("   ✅ 结构核对：真实回读原文与译文，段落数/表格数/图片数"
                                       "一致；文档里本来就没有脚注/批注/页眉页脚部件。")
                elif counts_ok and parts_ok is None:
                    self.write_log("   ⚠️ 结构核对：正文计数一致，但附加部件没能比对"
                                   "（详见翻译报告），请人工确认后再交付。")
                else:
                    self.write_log("   ❗ 结构核对：结构有差异（正文计数"
                                   + ("一致" if counts_ok else "不一致")
                                   + "，附加部件"
                                   + ("一致" if parts_ok is True else "不一致或未比对")
                                   + "），详见翻译报告，请人工确认后再交付。")
            self.last_quality_note = f"质检评分：{quality['score']}/100（{quality['verdict']}）"
            self.log_queue.put(("done", True, output_path))
        except Exception as error:
            self.write_log(f"❌ 运行出错：{error}")
            self.write_log("---- 详细报错 ----")
            for line in traceback.format_exc().splitlines():
                self.write_log(line)
            self.log_queue.put(("done", False, ""))
    def _run_translate_job(self, job, client, glossary, direction):
        """
        ★ 新增：在【工作线程】里执行一个翻译任务。
        这里只做一件事：调 API 拿译文。绝不碰 Word 文档（paragraph、doc 一个都不传进来），
        所以 4 个线程同时跑也不会损坏文件——这就是「并发请求，排队写入」的前半句。
        返回 {"entries": [(段落条目, 译文, 诊断), ...], "batch_diagnostics": {...}|None}
        """
        entries = job["entries"]
        # —— 多条合批：一次请求翻一批（省 API 调用次数）——
        if len(entries) > 1:
            batch_results, batch_diagnostics = translate_batch_items(
                client, [entry["source"] for entry in entries], glossary, direction,
                log=self.write_log, should_stop=self.stop_event.is_set)
            return {"entries": [(entry, result["translation"], result["diagnostics"])
                                for entry, result in zip(entries, batch_results)],
                    "batch_diagnostics": batch_diagnostics}
        # —— 单条：混排段落或超长段落，走普通单段翻译 ——
        entry = entries[0]
        result = _translate_one_item(
            client, entry["source"], glossary, direction,
            log=self.write_log, should_stop=self.stop_event.is_set,
            prompt_extra=entry["prompt_extra"],
            api_source=entry["api_source"], check_source=entry["check_source"])
        return {"entries": [(entry, result["translation"], result["diagnostics"])],
                "batch_diagnostics": None}
    def _on_worker_done(self, success, output_path):
        """后台线程结束后，回到主线程做收尾（恢复按钮、弹提示框）。"""
        self.is_running = False
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        if success:
            self.last_output_path = output_path
            self.open_folder_button.configure(state="normal")
            report_path = os.path.splitext(output_path)[0] + "_翻译报告.txt"
            report_hint = f"\n\n翻译报告：\n{report_path}" if os.path.exists(report_path) else ""
            score_hint = f"\n\n{self.last_quality_note}" if self.last_quality_note else ""
            messagebox.showinfo("翻译完成",
                                f"已完成！结果保存在：\n{output_path}{score_hint}{report_hint}")
        else:
            messagebox.showerror("出错了", "翻译过程中出现错误，详情请看下方日志。")
    def open_output_folder(self):
        """用资源管理器打开结果文件所在目录（Windows 专用）。"""
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showinfo("提示", "还没有生成翻译结果。")
            return
        try:
            os.startfile(os.path.dirname(self.last_output_path))
        except Exception as error:
            messagebox.showerror("打开失败", str(error))
    def _on_close(self):
        """关闭窗口：正在翻译时先确认，避免客户误关丢结果。"""
        if self.is_running:
            if not messagebox.askyesno(
                    "正在翻译",
                    "翻译还没结束，确定要退出吗？\n"
                    "（直接退出不会保存已翻译的部分，建议先点「停止」）"):
                return
        self.root.destroy()
# =====================================================================
#                            程序入口
# =====================================================================
def main():
    root = tk.Tk()
    TranslatorApp(root)
    root.mainloop()
if __name__ == "__main__":
    main()