import pytesseract
from PIL import Image
import cv2
import openpyxl
import re

# Tesseractのパスを設定
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\timaz\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def preprocess_image(image_path):
    image = cv2.imread(image_path)
    if image is None:
        raise FileNotFoundError(f"Image not found: {image_path}")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    processed_image = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
    )
    return processed_image

def extract_text(image_path):
    processed_image = preprocess_image(image_path)
    text = pytesseract.image_to_string(processed_image, lang='jpn')
    return text

def parse_raw_text(raw_text):
    lines = raw_text.split('\n')
    data = {}
    for line in lines:
        if ':' in line:  # 時間が含まれる行を探す
            name, time = line.split()[:2]  # 名前と時間を分離（仮の形式）
            data[name] = time
    return data

def update_excel(excel_path, data, output_path):
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    for name, time in data.items():
        for row in sheet.iter_rows(min_row=2):
            if row[0].value == name:
                row[1].value = time
    workbook.save(output_path)

# 実行
image_path = r'C:\Users\timaz\Documents\LINEfiles\1128.jpg'
excel_path = r'C:\Users\timaz\Documents\LINEfiles\1128.xlsx'
output_path = r'C:\Users\timaz\Documents\LINEfiles\11282.xlsx'

try:
    raw_text = extract_text(image_path)
    print("OCR Result:", raw_text)
    data = parse_raw_text(raw_text)
    update_excel(excel_path, data, output_path)
    print("Excel updated successfully.")
except Exception as e:
    print("Error:", str(e))
