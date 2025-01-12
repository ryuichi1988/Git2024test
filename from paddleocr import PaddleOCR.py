from paddleocr import PaddleOCR

# 初始化 PaddleOCR
ocr = PaddleOCR(det=True, rec=True, table=True, lang='en')

# 图片路径
image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\test1128.jpg'

# 表格识别
results = ocr.ocr(image_path, cls=True)

# 打印结果
print(results)
