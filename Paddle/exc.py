import pdfplumber

pdf_path = "cc7.pdf"

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        words = page.extract_words()  # 提取单词及坐标信息

        lines = {}  # 存储按 Y 坐标分组的数据
        y_threshold = 5  # 允许的 Y 轴误差范围

        for word in words:
            x0, y0, x1, y1 = word["x0"], word["top"], word["x1"], word["bottom"]
            text = word["text"]

            # 按 Y 坐标分行，避免误合并
            found = False
            for y_key in lines.keys():
                if abs(y_key - y0) < y_threshold:  # 误差在 5px 内，认为同一行
                    lines[y_key].append((x0, text))
                    found = True
                    break

            if not found:
                lines[y0] = [(x0, text)]

        # 按 Y 轴排序
        sorted_lines = sorted(lines.items(), key=lambda item: item[0])

        for y, words in sorted_lines:
            words.sort()  # 按 X 坐标排序，保持正确顺序
            line_text = " ".join([word[1] for word in words])
            print("line:", line_text)  # 输出按行排序后的文本
