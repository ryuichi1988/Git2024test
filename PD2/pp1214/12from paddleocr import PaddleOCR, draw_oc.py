from paddleocr import PaddleOCR, draw_ocr
import matplotlib.pyplot as plt
from PIL import Image

# 初始化 PaddleOCR，并加载推理模型
ocr = PaddleOCR(
    det_model_dir='./inference/det',  # 如果使用检测模型，指定路径
    rec_model_dir=r'C:\Users\timaz\PaddleOCR\inference',     # 推理模型路径
    use_angle_cls=True,              # 是否使用角度分类模型
    use_gpu=False,                   # 如果没有 GPU，设置为 False
    lang='en'                        # 设置语言类型，中文为 'ch'
)

# 要预测的图片路径
image_path = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\images\image1.jpg"

# 执行 OCR
results = ocr.ocr(image_path, cls=True)

# 打印结果
for line in results[0]:
    print(f"Detected text: {line[1][0]}, Confidence: {line[1][1]}")

# 可视化结果
image = Image.open(image_path).convert('RGB')
boxes = [result[0] for result in results[0]]
texts = [result[1][0] for result in results[0]]
scores = [result[1][1] for result in results[0]]

# 绘制识别结果
drawn_image = draw_ocr(image, boxes, texts, scores, font_path='./doc/fonts/simfang.ttf')
plt.imshow(drawn_image)
plt.axis('off')
plt.show()
