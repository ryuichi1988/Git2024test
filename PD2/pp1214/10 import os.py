import os

# 图片文件夹路径
cropped_images_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_images"

# 新标注文件路径
new_label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_labels.txt"

# 检查每个标注文件是否有对应图片
with open(new_label_file, "r", encoding="utf-8") as f:
    for line in f:
        image_filename, transcription = line.strip().split("\t")
        image_path = os.path.join(cropped_images_dir, image_filename)
        if not os.path.exists(image_path):
            print(f"图片缺失: {image_filename}")
