import cv2
from paddleocr import PaddleOCR,draw_ocr

img_path = "general_ocr_001.jpg"
frame = cv2.imread(img_path)
ocr = PaddleOCR(use_angle_cls=True, lang='japan', show_log=False, use_gpu=False)
result = ocr.ocr(frame, cls=False)
# print(result)



from PIL import Image
res = result[0]
image = Image.open(img_path).convert('RGB')
boxes = [line[0] for line in res]
txts = [line[1][0] for line in res]
scores = [line[1][1] for line in res]
im_show = draw_ocr(image, boxes, txts, scores, font_path='C:/Windows/Fonts/msyh.ttc') # font_path='./TaipeiSansTCBeta-Regular.ttf'
im_show = Image.fromarray(im_show)
im_show