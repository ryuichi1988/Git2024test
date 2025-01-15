from paddleocr import PaddleOCR, draw_ocr
import re

# 初始化OCR
ocr = PaddleOCR(use_angle_cls=False, lang="ch")  # 设置OCR模型为中文
img_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example1.jpg'  # 设置图片路径
result = ocr.ocr(img_path, cls=False)  # 执行OCR识别

# 用于处理日期、出勤、午休和下班时间的函数
def extract_times(data):
    # 用于存储最终结果
    result = []

    # 遍历每一条记录
    for entry in data:
        # 获取文本内容，entry[1][0] 是文本内容
        text = entry[1][0]  # OCR提取的文本内容

        # 获取日期
        date_match = re.match(r"(\d{1,2}/\d{1,2})", text)  # 使用正则提取日期
        date = date_match.group(1) if date_match else ""

        # 初始化时间变量
        attendance = ''
        lunch_start = ''
        lunch_end = ''
        leave = ''

        # 提取出勤时间
        attendance_match = re.search(r'(\d{1,2}:\d{2})', text)  # 查找出勤时间
        if attendance_match:
            attendance = attendance_match.group(1)

        # 提取午休开始和结束时间
        lunch_match = re.findall(r'(\d{1,2}:\d{2})(当日)?', text)  # 查找午休时间段
        if lunch_match:
            lunch_times = ', '.join([match[0] for match in lunch_match])  # 合并所有的午休时间
            lunch_times = lunch_times.replace('当日', '').strip()

            # 智能拆分午休开始和结束时间
            if ',' in lunch_times:
                lunch_start, lunch_end = lunch_times.split(',')
            else:
                lunch_start = lunch_times  # 只有一个时间则作为开始时间

        # 提取下班时间
        leave_match = re.search(r'(\d{1,2}:\d{2})', text)  # 查找下班时间
        if leave_match:
            leave = leave_match.group(1)

        # 添加结果
        if date:  # 只有在存在日期的情况下才添加结果
            result.append({
                "Date": date,
                "Attendance": attendance,
                "Lunch_start": lunch_start,
                "Lunch_end": lunch_end,
                "Leave": leave
            })

    return result

# 提取时间数据
extracted_times = extract_times(result)

# 输出结果
for entry in extracted_times:
    print(f"Date: {entry['Date']}")
    print(f"  Attendance: {entry['Attendance']}")
    print(f"  Lunch_start: {entry['Lunch_start']}")
    print(f"  Lunch_end: {entry['Lunch_end']}")
    print(f"  Leave: {entry['Leave']}\n")
