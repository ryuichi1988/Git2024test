from paddleocr import PaddleOCR
import pandas as pd

# Initialize PaddleOCR
ocr = PaddleOCR(use_angle_cls=True, lang='en')

# Path to the image
image_path = r'C:\Users\timaz\Documents\PythonFile\PD2\train_data\images\image1.jpg' # Replace with your actual image path

# Perform OCR on the image
ocr_result = ocr.ocr(image_path, cls=True)

# Extract recognized text
extracted_text = [line[1][0] for line in ocr_result[0]]

# Create a DataFrame from the extracted text
df = pd.DataFrame(extracted_text, columns=["Time"])

# Save the DataFrame to an Excel file
output_path = 'extracted_times.xlsx'
df.to_excel(output_path, index=False)

print(f"Extracted times saved to {output_path}")
