import cv2
import pytesseract

# 设置 Tesseract OCR 路径
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\timaz\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# 加载图像
image_path = r'C:\1\1.jpg'
image = cv2.imread(image_path)

# 预处理图像（可选）
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
_, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

# 定义模板区域（以像素为单位的矩形框: (x1, y1, x2, y2)）
template_regions = {
    "field_1":(30, 95, 559, 208),
    "field_2":(30, 95, 559, 208),
}

# 提取每个区域的文字
extracted_data = {}
for field_name, (x1, y1, x2, y2) in template_regions.items():
    region = thresh[y1:y2, x1:x2]  # 裁剪区域
    text = pytesseract.image_to_string(region, lang='jpn')  # 可以指定语言
    extracted_data[field_name] = text.strip()

# 打印提取的结果
for field, value in extracted_data.items():
    print(f"{field}: {value}")
