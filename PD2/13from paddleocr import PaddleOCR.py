from paddleocr import PaddleOCR
ocr = PaddleOCR(rec_model_dir="./inference", use_gpu=False)
results = ocr.ocr(r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\images\image1.jpg")
print(results)