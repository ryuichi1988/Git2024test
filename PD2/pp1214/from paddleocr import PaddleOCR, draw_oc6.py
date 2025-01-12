from paddleocr import PaddleOCR, draw_ocr
import pandas as pd
import os

# 初始化OCR，加载自定义训练好的模型
ocr = PaddleOCR(det_model_dir='inference/',  # 替换为你的模型路径
                rec_model_dir='inference/',
                use_angle_cls=True, 
                lang='en')  # 根据你的训练语言设置

# 定义图片路径
image_folder = r'C:\Users\timaz\Documents\PythonFile\PD2\train_data\images'  # 替换为你的图片文件夹路径
image_files = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(('.jpg', '.png'))]

# 存储识别结果
results = []

for image_path in image_files:
    ocr_result = ocr.ocr(image_path, cls=True)
    
    # 提取识别的文本
    for line in ocr_result[0]:
        text = line[1][0]
        results.append({'image': os.path.basename(image_path), 'time': text})

# 保存到Excel文件
df = pd.DataFrame(results)
output_path = 'recognized_times.xlsx'
df.to_excel(output_path, index=False)

print(f"识别结果已保存到 {output_path}")
