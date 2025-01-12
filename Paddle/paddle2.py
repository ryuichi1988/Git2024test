from paddleocr import PaddleOCR, draw_ocr
from PIL import Image, ImageDraw, ImageFont
import cv2
import matplotlib.pyplot as plt
import numpy as np
 
# Paddleocr目前支持的多语言语种可以通过修改lang参数进行切换
# 例如`ch`, `en`, `fr`, `german`, `korean`, `japan`
ocr = PaddleOCR(use_angle_cls=True, lang="ch")  # need to run only once to download and load model into memory
img_path = "general_ocr_001.jpg"
result = ocr.ocr(img_path, cls=True)
for idx in range(len(result)):
    res = result[idx]
    for line in res:
        print(line)
 
 
result = result[0]
boxleftup = [line[0][0] for line in result]
boxleftup = [[int(x) for x in row] for row in boxleftup]
boxrightbotm=[line[0][2] for line in result]
boxrightbotm = [[int(x) for x in row] for row in boxrightbotm]
txts = [line[1][0] for line in result]
scores = [line[1][1] for line in result]
 
merged_list = []
for i in range(len(result)):
    merged_list.append([boxleftup[i], boxrightbotm[i], txts[i]])
print(merged_list)
 
image = Image.open(img_path).convert('RGB')
image = np.asarray(image)
 
# 将 NumPy 图像转换为 PIL 图像
pil_img = Image.fromarray(image)
 
# 创建绘图对象
draw = ImageDraw.Draw(pil_img)
font = ImageFont.truetype("simfang.ttf", 30)
for (start, end, text) in merged_list:
    draw.rectangle([start[0], start[1], end[0], end[1]], outline="red", width=2)
    draw.text((start[0], start[1] - 30), text, fill=(0, 0, 0), font=font)  # 坐标、文本内容、颜色、字体
 
plt.axis('off')
plt.imshow(pil_img)
plt.savefig('output1.jpg', transparent=True, dpi=500)
plt.show()
