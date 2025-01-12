from paddleocr import PaddleOCR
import pandas as pd

import cv2

# 加载图片
image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\test1128.jpg'
image = cv2.imread(image_path)

# 转为灰度图
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# 二值化处理（增强表格线条）
_, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)

# 形态学操作（突出表格线条）
kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 1))
horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 15))
vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

# 合并线条
table_lines = cv2.add(horizontal_lines, vertical_lines)

# 保存增强后的图片
enhanced_image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\enhanced_table.jpg'
cv2.imwrite(enhanced_image_path, table_lines)

print(f"增强后的表格图片已保存为: {enhanced_image_path}")


# 初始化 PaddleOCR
ocr = PaddleOCR(det=True, rec=True, table=True, lang='en')

# 表格识别
results = ocr.ocr(image_path, cls=True)

for box in results[0]:
    print(f"检测到的文本框：{box[0]}, 文本：{box[1][0]}, 置信度：{box[1][1]}")


# 检查结果是否包含表格数据
if len(results) > 1 and len(results[1]) > 0:
    table_results = results[1]  # 获取表格识别结果
    data = []

    # 解析每个单元格
    for table in table_results:
        cells = table.get('html', {}).get('cells', [])
        for cell in cells:
            row = cell.get('row', 0)
            col = cell.get('col', 0)
            text = cell.get('text', '')
            data.append((row, col, text))

    # 转换为 Pandas DataFrame
    df = pd.DataFrame(data, columns=['Row', 'Column', 'Text'])
    df_pivot = df.pivot(index='Row', columns='Column', values='Text')

    # 保存为 Excel 文件
    output_path = r'C:\Users\timaz\Documents\PythonFile\PD2\output_table.xlsx'
    df_pivot.to_excel(output_path, index=False)
    
    print(f"表格内容已提取并保存为 {output_path}")
else:
    print("未检测到表格内容")
