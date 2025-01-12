from PIL import Image
res = result[0]
image = Image.open(img_path).convert('RGB')
boxes = [line[0] for line in res]
txts = [line[1][0] for line in res]
scores = [line[1][1] for line in res]
im_show = draw_ocr(image, boxes, txts, scores, font_path='./TaipeiSansTCBeta-Regular.ttf')
im_show = Image.fromarray(im_show)
im_show