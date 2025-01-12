import os
import json

# 输入裁剪后的图片文件夹路径
cropped_images_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_images"

# 原始标注文件路径
original_label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\Label_paddle\label.txt"

# 输出新的标注文件路径
new_label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_labels.txt"

# 读取原始标注文件
with open(original_label_file, "r", encoding="utf-8") as f:
    line = f.readline()  # 假设标注文件只有一行
    parts = line.split("\t")
    annotations = json.loads(parts[1])  # 解析标注信息

# 遍历裁剪后的图片文件夹，生成新的标注文件
with open(new_label_file, "w", encoding="utf-8") as new_label:
    for idx, annotation in enumerate(annotations):
        transcription = annotation["transcription"]

        # 构造对应裁剪后的图片文件名
        image_filename = f"crop_{idx}_{transcription.replace(':', '_')}.jpg"  # 替换非法字符

        # 确认图片存在
        image_path = os.path.join(cropped_images_dir, image_filename)
        if os.path.exists(image_path):
            # 写入新标注文件
            new_label.write(f"{image_filename}\t{transcription}\n")
            print(f"Added label for {image_filename}")

print(f"新标注文件已保存到: {new_label_file}")
