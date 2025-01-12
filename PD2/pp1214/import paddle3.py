from paddleocr import PaddleOCR
import re
import pandas as pd

# Initialize PaddleOCR
ocr = PaddleOCR(use_angle_cls=True, lang='japan')  # Japanese language support

# Image path
image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\example.jpg'

# OCR Recognition
results = ocr.ocr(image_path, cls=True)

# Extracting time information (HH:MM and HH:MM:SS formats)
time_patterns = [
    r'\b\d{2}:\d{2}\b',          # HH:MM format
    r'\b\d{2}:\d{2}:\d{2}\b',    # HH:MM:SS format
]

# Collect extracted times
extracted_times = []
for line in results[0]:
    text = line[1][0]  # Recognized text
    for pattern in time_patterns:
        matches = re.findall(pattern, text)
        if matches:
            extracted_times.extend(matches)

# Save to DataFrame for Excel output
df_times = pd.DataFrame(extracted_times, columns=["Extracted Times"])

# Save results to Excel file
output_path = r'extracted_times.xlsx'
df_times.to_excel(output_path, index=False)

print(f"Extracted times saved to: {output_path}")
