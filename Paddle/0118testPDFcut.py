import fitz  # PyMuPDF
from paddleocr import PaddleOCR
import re

ocr = PaddleOCR(
    use_angle_cls=False,
    lang='ch',
    table=True,  # 启用表格识别
    table_algorithm='TableAttn',  # 设置表格识别算法
    table_max_len=488,  # 设置表格最大长度
) # 初始化OCR

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

        # 打印 OCR 识别结果
        for line in result[0]:
            # 计算文本框的中心点
            text = line[1][0]
            center_pointX = sum([point[0] for point in line[0]]) / 4
            center_pointY = sum([point[1] for point in line[0]]) / 4
            print("中心{},{}:{} {}".format(center_pointX,center_pointY,line[0],line[1]))  # 打印识别的文本

            if "氏名" in text:
                namenow = text

