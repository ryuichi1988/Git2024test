import cv2
import easyocr
import openpyxl

# 图像预处理
def preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    _, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return img_bin

# OCR识别
def extract_times(image_path):
    reader = easyocr.Reader(['en'])
    processed_img = preprocess_image(image_path)
    results = reader.readtext(processed_img)
    times = [res[1] for res in results if ":" in res[1]]
    return times

# 写入Excel
def write_to_excel(times, excel_path, start_row=10, column="I"):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    col_idx = openpyxl.utils.column_index_from_string(column)
    for i, time in enumerate(times, start=start_row):
        sheet.cell(row=i, column=col_idx).value = time
    output_path = "Updated_File.xlsx"
    wb.save(output_path)
    return output_path

# 主程序
image_path = "test1128.jpg"
excel_path = "1128.xlsx"
times = extract_times(image_path)
output_file = write_to_excel(times, excel_path)
print(f"时间已填写至：{output_file}")
