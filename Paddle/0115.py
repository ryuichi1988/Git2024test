from paddleocr import PaddleOCR, draw_ocr
# Paddleocr supports Chinese, English, French, German, Korean and Japanese
# You can set the parameter `lang` as `ch`, `en`, `french`, `german`, `korean`, `japan`
# to switch the language model in order
ocr = PaddleOCR(use_angle_cls=False, lang="ch") # need to run only once to download and load model into memory
img_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example1.jpg'
result = ocr.ocr(img_path, cls=False)

import re


# 用于处理日期、出勤、午休和下班时间的函数
def extract_times(data):
    # 用于存储最终结果
    result = []

    # 遍历每一条记录
    for entry in data:
        # 获取日期
        date = entry[1].split()[0]  # 提取日期部分

        # 初始化时间变量
        attendance = ''
        lunch_start = ''
        lunch_end = ''
        leave = ''

        # 提取出勤时间
        attendance_match = re.search(r'(\d{1,2}:\d{2})', entry[0][0])  # 从X坐标大约365至475范围查找出勤时间
        if attendance_match:
            attendance = attendance_match.group(1)

        # 提取午休开始和结束时间
        lunch_match = re.findall(r'(\d{1,2}:\d{2})(当日)?', entry[0][2])  # 查找午休时间段（包括连在一起的情况）
        if lunch_match:
            lunch_times = ', '.join([match[0] for match in lunch_match])  # 合并所有的午休时间
            lunch_times = lunch_times.replace('当日', '').strip()

            # 智能拆分午休开始和结束时间
            if ',' in lunch_times:
                lunch_start, lunch_end = lunch_times.split(',')
            else:
                lunch_start = lunch_times  # 只有一个时间则作为开始时间

        # 提取下班时间
        leave_match = re.search(r'(\d{1,2}:\d{2})', entry[0][3])  # 查找下班时间
        if leave_match:
            leave = leave_match.group(1)

        # 添加结果
        result.append({
            "Date": date,
            "Attendance": attendance,
            "Lunch_start": lunch_start,
            "Lunch_end": lunch_end,
            "Leave": leave
        })

    return result


# 测试数据
data = [
    [[79.0, 308.0], ['12/05水', 0.998633623123169]],
    [[212.0, 306.0], ['契外', 0.9996782541275024]],
    [[294.0, 308.0], ['P1', 0.9582325220108032]],
    [[77.0, 327.0], ['12/06末', 0.9048758149147034]],
    [[294.0, 328.0], ['P1', 0.9773316383361816]],
    [[373.0, 327.0], ['当日7:50', 0.9197350144386292]],
    [[506.0, 327.0], ['当日17:03', 0.9505756497383118]],
    [[642.0, 324.0], ['当日11:12', 0.9596339464187622]],
    [[739.0, 325.0], ['当日12:07', 0.9404247403144836]],
    [[993.0, 328.0], ['8:00', 0.9866619110107422]],
    [[1353.0, 327.0], ['当日7：50当日17:03', 0.9270889759063721]]
]

# 提取时间数据
extracted_times = extract_times(data)

# 输出结果
for entry in extracted_times:
    print(f"Date: {entry['Date']}")
    print(f"  Attendance: {entry['Attendance']}")
    print(f"  Lunch_start: {entry['Lunch_start']}")
    print(f"  Lunch_end: {entry['Lunch_end']}")
    print(f"  Leave: {entry['Leave']}\n")

# draw result
from PIL import Image
result = result[0]
image = Image.open(img_path).convert('RGB')
boxes = [line[0] for line in result]
txts = [line[1][0] for line in result]
scores = [line[1][1] for line in result]
im_show = draw_ocr(image, boxes, txts, scores, font_path='/path/to/PaddleOCR/doc/fonts/simfang.ttf')
im_show = Image.fromarray(im_show)
im_show.save('result.jpg')