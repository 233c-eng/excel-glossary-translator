## 📄 多格式支持 (Multi-format Support)

本工具目前已支持多格式文件的本土化处理：
- **Excel (.xlsx)**：`translate_terms.py` —— 读取Excel的A列英文术语，批量翻译并写入B列，注重术语一致性。
- **Word (.docx)**：`translate_word.py` —— 读取Word文档段落，批量翻译成中文并生成新的Word文档，解决译员手动复制粘贴的痛点。
