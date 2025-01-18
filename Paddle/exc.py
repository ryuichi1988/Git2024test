from paddleocr import PaddleOCR
from pdf2image import convert_from_path
import os

# 初始化 PaddleOCR
ocr = PaddleOCR(use_angle_cls=False, lang="ch")  # 支持多语言，可以指定 lang="en" 或其他语言

# PDF 文件路径
pdf_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example4.pdf'
output_folder = "pdf_images"  # 保存 PDF 转换后的图片的目录
os.makedirs(output_folder, exist_ok=True)

# 将 PDF 转换为图像
images = convert_from_path(pdf_path, dpi=300, output_folder=output_folder)

# 遍历转换的图片并进行 OCR
for page_num, image in enumerate(images, start=1):
    # 保存图像文件（可选，用于调试）
    image_path = os.path.join(output_folder, f"page_{page_num}.png")
    image.save(image_path, "PNG")

    # OCR 处理
    result = ocr.ocr(image_path, cls=False)

    # 输出 OCR 结果
    print(f"Page {page_num} OCR Results:")
    for line in result[0]:  # 遍历识别结果
        print("{}   {}".format(line[0],line[1]))
