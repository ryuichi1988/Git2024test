from paddleocr import PaddleOCR, draw_ocr
from pdf2image import convert_from_path
from PIL import Image
import os
import re

# PDF to Image
pdf_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example4.pdf'
output_folder = "pdf_images"  # 保存 PDF 转换后的图片的目录
os.makedirs(output_folder, exist_ok=True)

images = convert_from_path(pdf_path, dpi=300, output_folder=output_folder)


# 全角字符到半角字符的转换表
def full_width_to_half_width(text):
    full_width_chars = "："
    half_width_chars = ":"
    trans = str.maketrans(full_width_chars, half_width_chars)
    return text.translate(trans)


# 去掉“当日”字样
def remove_todays_word(text):
    text = text.replace('当日', '')
    text = text.replace('前日', '')
    return text


# OCR 初始化
ocr = PaddleOCR(use_angle_cls=False, lang="ch")  # 只需运行一次，下载和加载模型到内存

# 定义时间分类的 X 坐标范围
time_ranges = {
    "attendance": (365, 475),  # 出勤时间
    "lunch_start": (635, 710),  # 午休开始时间
    "lunch_end": (725, 810),  # 午休结束时间
    "leave": (500, 610)  # 下班时间
}
y_tolerance = 50  # Y 坐标误差范围
attendance_data = []  # 初始化存储结果的列表

# 遍历 PDF 转换的图片
for page_num, image in enumerate(images, start=1):
    image_path = os.path.join(output_folder, f"page_{page_num}.png")
    image.save(image_path, "PNG")

    result = ocr.ocr(image_path, cls=True)
    grouped_lines = {}

    # 按 OCR 结果分组
    for line in result[0]:
        text = line[1][0]
        text = full_width_to_half_width(text)
        text = remove_todays_word(text)

        x_coord = int(line[0][0][0])
        y_coord = int(line[0][0][1])

        # 分组逻辑
        found_group = None
        for group_y in grouped_lines.keys():
            if abs(group_y - y_coord) <= y_tolerance:
                found_group = group_y
                break
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
            if re.search(r'\d{2}/\d{2}[^\d]*$', text):  # 日期信息
                date_info = re.sub(r'[^\d/]', '', text)
            elif re.search(r'\d{1,2}:\d{2}', text):  # 时间分类
                for time_type, (x_min, x_max) in time_ranges.items():
                    if x_min <= x_coord <= x_max:
                        times[time_type].append(text)
                        break

        # 处理午休时间连在一起的情况
        if times["lunch_start"]:
            fixed_lunch_start = []
            for entry in times["lunch_start"]:
                split_times = re.findall(r'\d{1,2}:\d{2}', entry)
                if len(split_times) == 2:
                    fixed_lunch_start.append(split_times[0])
                    times["lunch_end"].append(split_times[1])
                else:
                    fixed_lunch_start.append(entry)
            times["lunch_start"] = fixed_lunch_start

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

print("Processing completed!")
