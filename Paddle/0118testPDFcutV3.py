import datetime
import fitz  # PyMuPDF
from paddleocr import PaddleOCR
import re

# 初始化OCR
ocr = PaddleOCR(
    use_angle_cls=False,
    lang='ch',
    table=True,  # 启用表格识别
    table_algorithm='TableAttn',  # 设置表格识别算法
    table_max_len=488,  # 设置表格最大长度
)


# 全角字符到半角字符的转换函数
def full_width_to_half_width(text):
    full_width_chars = "："
    half_width_chars = ":"
    trans = str.maketrans(full_width_chars, half_width_chars)
    return text.translate(trans)


# 去掉“当日”和“前日”字样
def remove_todays_word(text):
    text = text.replace('当日', '')
    text = text.replace('前日', '')
    return text


# 处理日期格式，使用正则智能识别日期并去除后续多余字符
def process_date(text):
    # 匹配日期格式，如“07/14”或“07-14”
    date_pattern = r"\d{2}[/-]\d{2}"
    match = re.match(date_pattern, text)
    if match:
        return match.group()  # 返回匹配到的日期部分
    return text  # 如果没有匹配到日期，返回原文本


# 处理连在一起的时间，如“当日0:57当日1:56”
def split_time_range(text):
    # 匹配“当日”字样并拆分时间
    times = re.findall(r'\d{1,2}:\d{2}', remove_todays_word(text))
    if len(times) == 2:
        return times[0], times[1]
    return None, None


pdf_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example4.pdf'
with fitz.open(pdf_path) as pdf:
    for page_num in range(pdf.page_count):
        page = pdf.load_page(page_num)

        # 获取页面尺寸
        rect = page.rect
        width = rect.width
        height = rect.height

        # 设置 300 DPI 的缩放因子
        zoom_x = 200 / 72  # 300 DPI对应的X轴缩放因子
        zoom_y = 200 / 72  # 300 DPI对应的Y轴缩放因子
        matrix = fitz.Matrix(zoom_x, zoom_y)  # 使用matrix来设置DPI

        # 设置裁剪区域：X坐标从30%到55%，Y坐标从10%到90%
        crop_rect = fitz.Rect(width * 0.0, height * 0.15, width * 0.48, height * 0.75)

        # 裁剪页面
        page.set_cropbox(crop_rect)
        pix = page.get_pixmap(matrix=matrix)

        img_path = "cropped_image.png"
        pix.save(img_path)

        # 使用 OCR 进行识别
        result = ocr.ocr(img_path, cls=False)

        # 收集每行的时间
        times = []  # 用来存储每行的时间信息

        # 打印 OCR 识别结果
        for line in result[0]:
            # 计算文本框的中心点
            text = line[1][0]
            center_pointX = sum([point[0] for point in line[0]]) / 4
            center_pointY = sum([point[1] for point in line[0]]) / 4
            print("中心{},{}:{} {}".format(center_pointX, center_pointY, line[0], line[1]))  # 打印识别的文本

            # 如果是"氏名"字段
            if "氏名" in text:
                namenow = text
                print("发现氏名:", namenow)
                continue

            # 处理日期
            string = process_date(text)
            if string != text:  # 如果日期被提取出来
                print("处理后的日期:", string)
                date = string

            # 处理时间（移除“当日”和“前日”）
            stringt = remove_todays_word(text)
            if stringt != text:
                print("处理后的时间:", stringt)
                time = stringt
                # 存储时间及其X坐标
                times.append((time, center_pointX, center_pointY))

            # 处理连在一起的时间
            start_time, end_time = split_time_range(text)
            if start_time and end_time:
                print("休息时间：", start_time, "到", end_time)
            elif start_time and not end_time:  # 如果只有单一时间，不显示“休息时间”
                print("时间信息:", start_time)

            # 如果已经收集到4个时间
            if len(times) == 4:
                # 按照X坐标从小到大排序
                times.sort(key=lambda x: x[1])  # x[1]是center_pointX

                # 打印排序后的时间
                for t in times:
                    print(f"时间: {t[0]}，坐标X: {t[1]}，坐标Y: {t[2]}")

                # 清空times列表，准备处理下一行的时间信息
                times = []
