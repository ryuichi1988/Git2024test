import fitz  # PyMuPDF
import numpy as np
import re
from exc import pdf_path

# 打开 PDF 并提取文本
with fitz.open(pdf_path) as doc:
    data = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text("text")
        data.extend(text.split("\n"))  # 按行拆分文本

# 结果存储
structured_data = []
current_group = None

for item in data:
    print("item:")
    print(item)
    if isinstance(item, str) and item.startswith("NP"):
        # 遇到新的 "NPxxx" 开头的字符串，开始新分组
        if current_group:
            structured_data.append(current_group)  # 存储上一个组
        current_group = [item, []]  # 初始化新组
    else:
        # 把数字加入当前组
        if current_group:
            current_group[1].append(item)

# 添加最后一个组
if current_group:
    structured_data.append(current_group)

# 转换为 NumPy 数组，确保结构正确


structured_array = np.array([(group[0], np.array(group[1])) for group in structured_data], dtype=object)


# 打印结果
print(structured_array)