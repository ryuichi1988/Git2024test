import fitz  # PyMuPDF
import numpy as np
import re
from exc import pdf_path

# 读取 PDF 并提取文本
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
    if isinstance(item, str) and item.startswith("NP"):
        # 遇到新的 "NPxxx" 开头的字符串，开始新分组
        if current_group:
            structured_data.append(current_group)  # 存储上一个组
        current_group = [item, []]  # 初始化新组
    else:
        # 把数据加入当前组
        if current_group:
            current_group[1].append(item)

# 添加最后一个组
if current_group:
    structured_data.append(current_group)

# **🔹 合并相同名字的出勤记录**
merged_data = {}

for name, records in structured_data:
    # **🔴 如果 "name" 以 "合計" 结尾，则跳过**
    if name.endswith("合計"):
        print(f"Skipping: {name}")  # 调试输出
        continue

    raw_records = np.array(records, dtype=object)  # 转换为 NumPy 数组

    # **按 "x日" 进行分组**
    reshaped_records = []
    temp_group = []

    for item in raw_records:
        if re.match(r"\d{1,2}日", item):  # 如果是 "x日"，表示新的一组
            if temp_group:  # 存入上一个分组
                reshaped_records.append(temp_group[:4])  # 只取前 4 个
            temp_group = [item]  # 开启新分组
        else:
            temp_group.append(item)

    if temp_group:
        reshaped_records.append(temp_group[:4])  # 处理最后一组

    # **🔴 过滤掉第一个元素不包含 "日" 的行**
    filtered_records = [row for row in reshaped_records if row[0].endswith("日")]

    # **如果过滤后数据为空，则跳过这个组**
    if not filtered_records:
        print(f"Skipping group {name} due to no valid records")
        continue

    # **调整上班时间**
    for record in filtered_records:
        work_start_hour = record[1][:2]  # 时间段开始时间 (前 2 位小时)
        arrival_hour, arrival_minute = map(int, record[2].split(":"))  # 上班时间

        # **如果上班时间早于时间段开始时间**
        if work_start_hour == "00" and arrival_hour == 23:
            arrival_hour = -1

        if arrival_hour < int(work_start_hour):
            new_hour = int(work_start_hour)

            # **如果是 23 点，改为 00:00**


            # **修改上班时间**
            record[2] = f"{new_hour:02d}:00"

    # **合并相同名字的记录**
    if name in merged_data:
        merged_data[name] = np.vstack((merged_data[name], filtered_records))
    else:
        merged_data[name] = np.array(filtered_records, dtype=object)

# **转换为 NumPy 数组**
structured_array = np.array([(name, records) for name, records in merged_data.items()], dtype=object)

# **打印结果**
print(structured_array)
