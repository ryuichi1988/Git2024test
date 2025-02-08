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

# 正则表达式匹配出勤数据
for i, line in enumerate(lines):
    print(f"Line {i}: {repr(line)}")
attendance_data = []
for line in (lines):
    print(f"Checking type: {type(line)}")  # 看看是 str 还是 bytes

    match = re.match(r"(\d{1,2}日) (\d{4})-(\d{4}) (\d{1,2}:\d{2}) (\d{1,2}:\d{2})",
                     line)
    print(f"Checking line: {line} -> Match: {match}")

    if match:
        date = match.group(1)
        start_time = match.group(2)
        end_time = match.group(3)
        clock_in = match.group(4)
        clock_out = match.group(5)
        work_hours = float(match.group(6))
        other1 = float(match.group(7))
        other2 = float(match.group(8))
        total_hours = float(match.group(9))

        attendance_data.append(
            [date, start_time, end_time, clock_in, clock_out, work_hours, other1, other2, total_hours])

# 转换为 NumPy 数组
attendance_array = np.array(attendance_data, dtype=object)

# 打印结果
print(attendance_array)
