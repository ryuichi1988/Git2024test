from paddleocr import PaddleOCR, draw_ocr
import re
from PIL import Image

# 全角字符到半角字符的转换表
def full_width_to_half_width(text):
    full_width_chars = "．！？，。＠＃＄％＆＊（）＋－＝｜＜＞＿＾｛｝［］；：’＂＜＞"
    half_width_chars = ".!?,-@#$%&*()+-=_|<>^{}[];:'\"<>"
    trans = str.maketrans(full_width_chars, half_width_chars)
    return text.translate(trans)

# 去掉“当日”字样
def remove_todays_word(text):
    return text.replace('当日', '')

# PaddleOCR初始化
ocr = PaddleOCR(use_angle_cls=False, lang="ch")  # 只需运行一次，下载和加载模型到内存
img_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example1.jpg'
result = ocr.ocr(img_path, cls=False)

# 初始化存储结果的列表
attendance_data = []

# 定义时间分类的 X 坐标范围
time_ranges = {
    "attendance": (365, 475),   # 出勤时间
    "lunch_start": (635, 710),  # 午休开始时间
    "lunch_end": (725, 810),    # 午休结束时间
    "leave": (500, 610)         # 下班时间
}

# Y 坐标误差范围
y_tolerance = 20

# 遍历 OCR 结果
for idx in range(len(result)):
    res = result[idx]
    grouped_lines = {}  # 用于按 Y 坐标分组的字典

    for line in res:
        text = line[1][0]  # 获取文本内容
        text = full_width_to_half_width(text)  # 转换全角字符为半角字符
        text = remove_todays_word(text)  # 去掉“当日”字样
        x_coord = int(line[0][0][0])  # 获取第一个点的 X 坐标
        y_coord = int(line[0][0][1])  # 获取第一个点的 Y 坐标

        # 查找是否有已存在的相近 Y 坐标组
        found_group = None
        for group_y in grouped_lines.keys():
            if abs(group_y - y_coord) <= y_tolerance:
                found_group = group_y
                break

        # 将文本加入对应的分组
        if found_group is not None:
            grouped_lines[found_group].append((text, x_coord))
        else:
            grouped_lines[y_coord] = [(text, x_coord)]

    # 处理分组后的数据
    for group_y, texts in grouped_lines.items():
        date_info = None
        times = {
            "attendance": [],
            "lunch_start": [],
            "lunch_end": [],
            "leave": []
        }
        for text, x_coord in texts:
            # 判断是否为日期信息
            if re.search(r'\d{2}/\d{2}[^\d]*$', text):  # 匹配日期格式如“12/04水”
                date_info = re.sub(r'[^\d/]', '', text)  # 去掉非数字和斜杠部分
            # 判断时间并分类
            elif re.search(r'\d{1,2}:\d{2}', text):  # 匹配时间格式
                for time_type, (x_min, x_max) in time_ranges.items():
                    if x_min <= x_coord <= x_max:
                        times[time_type].append(text)
                        break

        # 处理午休时间连在一起的情况
        if times["lunch_start"]:
            fixed_lunch_start = []
            for entry in times["lunch_start"]:
                # 分割连在一起的午休时间
                split_times = re.findall(r'\d{1,2}:\d{2}', entry)
                if len(split_times) == 2:  # 如果找到两个时间，则分别存入开始和结束
                    fixed_lunch_start.append(split_times[0])
                    times["lunch_end"].append(split_times[1])
                else:
                    fixed_lunch_start.append(entry)
            times["lunch_start"] = fixed_lunch_start

        # 如果没有任何打卡记录，跳过
        if not (times["attendance"] or times["leave"] or times["lunch_start"] or times["lunch_end"]):
            continue

        # 保存记录
        if date_info:
            attendance_data.append({
                "date": date_info,
                "times": times
            })

# 打印结果
for data in attendance_data:
    print(f"Date: {data['date']}")
    for time_type, time_list in data["times"].items():
        print(f"  {time_type.capitalize()}: {', '.join(time_list)}")

# draw result
result = result[0]
image = Image.open(img_path).convert('RGB')
boxes = [line[0] for line in result]
txts = [line[1][0] for line in result]
scores = [line[1][1] for line in result]
im_show = draw_ocr(image, boxes, txts, scores, font_path='/path/to/PaddleOCR/doc/fonts/simfang.ttf')
im_show = Image.fromarray(im_show)
im_show.save('result.jpg')
