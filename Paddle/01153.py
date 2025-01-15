from paddleocr import PaddleOCR
import re

# 初始化OCR
ocr = PaddleOCR(use_angle_cls=False, lang="ch")  # 设置OCR模型为中文
img_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example1.jpg'  # 设置图片路径
result = ocr.ocr(img_path, cls=False)  # 执行OCR识别


# 用于处理时间的函数
def extract_times(data):
    # 用于存储最终结果
    result = []

    # 遍历每一条记录
    attendance = ''
    lunch_start = ''
    lunch_end = ''
    leave = ''
    date = ''

    for entry in data:
        text = entry[1][0]
        print(text)# OCR提取的文本内容
        coordinates = entry[0][0]  # 坐标信息

        # 根据X坐标范围分类
        print(entry[0])
        x_min, x_max = coordinates[0][0], coordinates[2][0]  # 获取X轴的最小值和最大值

        if 365 <= x_min <= 475:  # 出勤时间
            attendance = text.replace('当日', '').strip()
        elif 635 <= x_min <= 760:  # 午休开始时间
            lunch_start = text.replace('当日', '').strip()
        elif 720 <= x_min <= 840:  # 午休结束时间
            lunch_end = text.replace('当日', '').strip()
        elif 500 <= x_min <= 610:  # 下班时间
            leave = text.replace('当日', '').strip()

        # 根据日期标记
        date_match = re.match(r"(\d{1,2}/\d{1,2})", text)  # 查找日期
        if date_match:
            date = date_match.group(1)

        # 将结果添加到列表
        if date:
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
