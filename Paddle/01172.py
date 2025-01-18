"""
# 定义时间分类的 X 坐标范围
time_ranges = {
    "attendance": (365, 475),  # 出勤时间
    "lunch_start": (635, 710),  # 午休开始时间
    "lunch_end": (725, 810),  # 午休结束时间
    "leave": (500, 610)  # 下班时间
}
y_tolerance = 50  # Y 坐标误差范围
attendance_data = []  # 初始化存储结果的列表


其实不需要如此精确的指定坐标，因为pdf里的信息，每个人是一页，然后从1号至31号的数据都是横向铺开的。类似于
1月1日  8：00   17:00  12；00 13:00
1月2日
1月3日
等等，以此类推

所以，同一行（Y坐标大概一致）的数据，就是同一个 attendance，leave

ChatGPT 说：
ChatGPT
明白了，如果 PDF 中每个人的数据是一页，并且同一天的时间数据是横向排列的，那么可以简化坐标逻辑，只需根据 Y 坐标的分组来区分日期的记录。同一行的数据属于同一个日期，而无需精确指定 X 坐标范围。

以下是调整后的代码版本：
"""


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
y_tolerance = 20  # Y 坐标误差范围
attendance_data = []  # 初始化存储结果的列表

# 遍历 PDF 转换的图片
for page_num, image in enumerate(images, start=1):
    image_path = os.path.join(output_folder, f"page_{page_num}.png")
    image.save(image_path, "PNG")

    result = ocr.ocr(image_path, cls=True)
    grouped_lines = {}

    # 按 Y 坐标分组
    for line in result[0]:
        text = line[1][0]
        text = full_width_to_half_width(text)
        text = remove_todays_word(text)

        y_coord = int(line[0][0][1])

        # 查找是否有已存在的相近 Y 坐标组
        found_group = None
        for group_y in grouped_lines.keys():
            if abs(group_y - y_coord) <= y_tolerance:
                found_group = group_y
                break
        if found_group is not None:
            grouped_lines[found_group].append(text)
        else:
            grouped_lines[y_coord] = [text]

    # 处理分组后的数据
    for group_y, texts in grouped_lines.items():
        # 假设第一列是日期，后面是时间
        date_info = None
        times = []
        for i, text in enumerate(texts):
            if i == 0 and re.search(r'\d{1,2}/\d{1,2}', text):  # 识别日期格式
                date_info = text
            elif re.search(r'\d{1,2}:\d{2}', text):  # 识别时间格式
                times.append(text)

        if date_info:
            attendance_data.append({
                "date": date_info,
                "times": times
            })

# 打印结果
for data in attendance_data:
    print(f"Date: {data['date']}")
    print(f"  Times: {', '.join(data['times'])}")

print("Processing completed!")
