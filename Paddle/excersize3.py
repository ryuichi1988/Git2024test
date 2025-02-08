import fitz  # PyMuPDF
import numpy as np
import re
from exc import pdf_path

# 打开 PDF 并提取文本
with fitz.open(pdf_path) as doc:
    lines = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text("text")
        lines.extend(text.split("\n"))  # 按行拆分文本


print(lines)
print(len(lines))