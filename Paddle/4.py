import cv2
from paddleocr import PaddleOCR, draw_ocr
from PIL import Image
import matplotlib.pyplot as plt

# 加载图像路径
img_path = "general_ocr_001.jpg"

# 读取原始图像
image = Image.open(img_path)

# 如果图像太大，可以在进行OCR之前进行尺寸调整
max_size = 2048  # 最大尺寸限制
if max(image.size) > max_size:
    image.thumbnail((max_size, max_size))  # 调整图像大小

# 初始化OCR对象
ocr = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False, use_gpu=False)

# 进行OCR识别
result = ocr.ocr(image, cls=False)

# 获取OCR结果
res = result[0]
boxes = [line[0] for line in res]  # 获取文本框
txts = [line[1][0] for line in res]  # 获取文本
scores = [line[1][1] for line in res]  # 获取置信度

# 绘制OCR结果
font_path = 'C:/Windows/Fonts/msyh.ttc'  # 字体路径（确保路径正确）
im_show = draw_ocr(image, boxes, txts, scores, font_path=font_path)

# 保持原始图像的尺寸（避免在绘制时被缩小）
im_show = Image.fromarray(im_show)
im_show = im_show.resize(image.size)  # 使用原图尺寸调整

# 显示图像
plt.figure(figsize=(image.width / 100, image.height / 100), dpi=100)
plt.imshow(im_show)
plt.axis('off')  # 不显示坐标轴
plt.show()

# 可选：保存图像
im_show.save("output_image.jpg")
