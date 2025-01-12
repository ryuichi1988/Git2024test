import cv2
import pytesseract
import pandas as pd

# 配置 Tesseract-OCR 的路径
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\timaz\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'


# 加载图片
image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\test1128.jpg'
image = cv2.imread(image_path)

# 预处理图片（灰度化和二值化）
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
_, binary_image = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY)
sharpen_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
sharp_image = cv2.morphologyEx(binary_image, cv2.MORPH_CLOSE, sharpen_kernel)
binary_image = cv2.adaptiveThreshold(
    gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2
)
denoised_image = cv2.fastNlMeansDenoising(gray, None, 30, 7, 21)

# OCR 识别
custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789:'
text = pytesseract.image_to_string(binary_image, config=custom_config)

contours, _ = cv2.findContours(binary_image, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    if w > 10 and h > 10:  # 忽略小区域
        cell_image = binary_image[y:y+h, x:x+w]
        cell_text = pytesseract.image_to_string(cell_image, config=custom_config)
        print(f'单元格内容: {cell_text}')




# 将文本按行拆分
lines = text.split('\n')
data = [line.split() for line in lines if line.strip()]  # 去除空行并按空格拆分

# 转为 Pandas DataFrame
df = pd.DataFrame(data)

# 保存到 Excel
excel_path = 'outputtest1213.xlsx'
df.to_excel(excel_path, index=False, header=False)
print(f'识别结果已保存到 {excel_path}')