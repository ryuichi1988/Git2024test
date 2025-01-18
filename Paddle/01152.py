from paddleocr import PaddleOCR, draw_ocr
from PIL import Image

# Initialize PaddleOCR with Chinese language
ocr = PaddleOCR(use_angle_cls=False, lang="ch")

# Load image and perform OCR
img_path = r'C:\Users\timaz\Documents\PythonFile\pd2\example5.jpg'
result = ocr.ocr(img_path, cls=False)

# Print OCR result
for idx in range(len(result)):
    res = result[idx]
    for line in res:
        print(line)

# Draw OCR results
image = Image.open(img_path).convert('RGB')
boxes = [line[0] for line in result[0]]
txts = [line[1][0] for line in result[0]]
scores = [line[1][1] for line in result[0]]

# Make sure to not resize the image and draw OCR on the original size
original_size = image.size  # 保存原始图像的大小
im_show = draw_ocr(image, boxes, txts, scores, font_path='/path/to/PaddleOCR/doc/fonts/simfang.ttf')
im_show = Image.fromarray(im_show)
print("{}{}".format(original_size,im_show.size))
# 比较大小，确保没有改变

im_show.save('result.jpg')
