from PIL import Image
import json
import os

# 原始图片路径
image_path = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\images\image1.jpg"

# 标注文件路径
label_file_path = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\Label_paddle\label.txt"

# 输出分割图片的保存路径
output_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\cropped_images"
os.makedirs(output_dir, exist_ok=True)

# 打开图片
image = Image.open(image_path)

# 读取标注文件
with open(label_file_path, "r", encoding="utf-8") as f:
    line = f.readline()  # 假设标注文件只有一行
    parts = line.split("\t")
    annotations = json.loads(parts[1])  # 获取标注信息

# 遍历标注信息，裁剪图片# 遍历标注信息，裁剪图片
for idx, annotation in enumerate(annotations):
    points = annotation["points"]
    transcription = annotation["transcription"]

    # 替换非法字符（如 ":"）
    sanitized_transcription = transcription.replace(":", "_")

    # 获取裁剪区域 (left, upper, right, lower)
    left = points[0][0]
    upper = points[0][1]
    right = points[2][0]
    lower = points[2][1]

    # 裁剪图片
    cropped_image = image.crop((left, upper, right, lower))

    # 保存裁剪后的图片
    cropped_image_path = os.path.join(output_dir, f"crop_{idx}_{sanitized_transcription}.jpg")
    cropped_image.save(cropped_image_path)
    print(f"Saved cropped image: {cropped_image_path}")
