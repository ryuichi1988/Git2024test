import numpy as np
import re
import fitz
import openpyxl
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from os import  system

final_mention_list = []

wb = openpyxl.load_workbook("CCtestM.xlsm", keep_vba=True)
source = wb["99999　ニッセープロダクツ"]
number_master_sheet = wb['Number_Master']
number_name_dict = {}
name_now = None
current_date = None
last_name = None  # 记录名字

for row in range(1, number_master_sheet.max_row + 1):
    CCNP_number = number_master_sheet.cell(row=row, column=1).value
    staff_name = number_master_sheet.cell(row=row, column=2).value
    NP_number = CCNP_number
    temp = {CCNP_number: (NP_number, staff_name)}
    number_name_dict.update(temp)

"""
根据您的需求，我们可以编写一个函数来实现这个功能。
这个函数将接受name作为参数，遍历字典，查找匹配的staff_name，
并返回相应的NP_number。如果找不到匹配的名字，
则返回"不明"。以下是实现这个功能的代码：

使用示例
name_to_search = "张三"  # 要查找的名字
result = find_NP_number(name_to_search, number_name_dict)
print(f"{name_to_search}的NP号码是：{result}")
"""
def find_NP_number(name, number_name_dict):
    for CCNP_number, (NP_number, staff_name) in number_name_dict.items():
        if staff_name == name:
            return NP_number
    return "不明"



root = tk.Tk()
root.withdraw()
root.lift()  # 提升窗口
root.attributes('-topmost', True)  # 置顶窗口

pdf_path = filedialog.askopenfilename(title="PDFデータをお選びください。",filetypes=[("pdfファイル","*.pdf")],defaultextension=".pdf")
root.destroy()  # 关闭主窗口
print(pdf_path)


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
        print(name)
        print(name_now)
        print(work_start_hour)

        NP__number = find_NP_number(name,number_name_dict)

        # **如果上班时间早于时间段开始时间**
        try:
            if work_start_hour == "00" and arrival_hour == 23:
                arrival_hour = -1

            if arrival_hour < int(work_start_hour):
                new_hour = int(work_start_hour)
        except ValueError:
            final_mention_list.append((name,NP__number,record,"勤務区分エラーです。手動で確認し、出勤時間を修正してください。"))
            continue

            # **如果是 23 点，改为 00:00**


            # **修改上班时间**
        record[2] = f"{new_hour:02d}:00"

    # **合并相同名字的记录**
    if name in merged_data:
        merged_data[name] = np.vstack((merged_data[name], filtered_records))
    else:
        merged_data[name] = np.array(filtered_records, dtype=object)

# **转换为 NumPy 数组**
structured_array = np.array([(name, NP__number, records) for name, records in merged_data.items()], dtype=object)

# **打印结果**
print(structured_array)
list_len = len(structured_array)
"""
n=2
for i in structured_array:
    arr_name_now = i[0]
    number_master_sheet.cell(row=n,column=3).value = arr_name_now
    n += 1
"""



output_xlsx = "CCMacro_2025_02test.xlsm"
wb.save(output_xlsx)


if final_mention_list:
    with open(f"CozyRecordMention.txt", "w", encoding="utf-8") as f:
        for sublist in final_mention_list:
            f.write(" ".join(map(str, sublist)) + "\n")  # 用空格分隔元素，并换行
    system(f"CozyRecordMention.txt")
