import os
import random
import shutil

# 原始图片路径和标注文件路径
image_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_images"
label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_labels.txt"

# 输出路径
train_image_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data_split\train\images"
val_image_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data_split\val\images"
train_label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data_split\train\labels.txt"
val_label_file = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data_split\val\labels.txt"

# 创建输出目录
os.makedirs(train_image_dir, exist_ok=True)
os.makedirs(val_image_dir, exist_ok=True)

# 读取标注文件
with open(label_file, "r", encoding="utf-8") as f:
    lines = f.readlines()

# 打乱标注行并按 80%:20% 划分
random.seed(42)  # 保证结果可复现
random.shuffle(lines)
split_idx = int(len(lines) * 0.8)
train_lines = lines[:split_idx]
val_lines = lines[split_idx:]

# 复制图片并生成新的标注文件
def copy_files_and_write_labels(lines, image_dir, output_image_dir, output_label_file):
    with open(output_label_file, "w", encoding="utf-8") as label_out:
        for line in lines:
            # 提取图片文件名
            image_filename, transcription = line.strip().split("\t")
            src_image_path = os.path.join(image_dir, image_filename)
            dest_image_path = os.path.join(output_image_dir, image_filename)
            
            # 复制图片
            if os.path.exists(src_image_path):
                shutil.copy(src_image_path, dest_image_path)
                label_out.write(line)
            else:
                print(f"Warning: Image not found: {src_image_path}")

# 处理训练集
copy_files_and_write_labels(train_lines, image_dir, train_image_dir, train_label_file)

# 处理验证集
copy_files_and_write_labels(val_lines, image_dir, val_image_dir, val_label_file)

print(f"数据集划分完成！")
print(f"训练集图片和标注文件保存在: {os.path.dirname(train_image_dir)}")
print(f"验证集图片和标注文件保存在: {os.path.dirname(val_image_dir)}")
