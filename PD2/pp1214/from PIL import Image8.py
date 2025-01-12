from PIL import Image, ImageOps
import os

# 输入裁剪图片的文件夹路径
input_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\cropped_images"
output_dir = r"C:\Users\timaz\Documents\PythonFile\PD2\train_data\resized_images"
os.makedirs(output_dir, exist_ok=True)

# 目标尺寸 (高度 32，宽度 64)
target_height = 32
target_width = 64

# 遍历所有裁剪后的图片
for file_name in os.listdir(input_dir):
    if file_name.endswith(".jpg"):
        image_path = os.path.join(input_dir, file_name)
        image = Image.open(image_path)

        # 调整大小
        resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)

        # 保存调整后的图片
        resized_image_path = os.path.join(output_dir, file_name)
        resized_image.save(resized_image_path)
        print(f"Resized and saved: {resized_image_path}")
