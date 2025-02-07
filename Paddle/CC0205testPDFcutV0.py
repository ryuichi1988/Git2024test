import re
import fitz
import openpyxl
from paddleocr import PaddleOCR
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

MAX_X_DISTANCE = 300

nowtime = datetime.now()
if nowtime.day < 10:
    nowMonth = nowtime.month - 1
    if nowMonth == 0:  # 处理跨年的情况
        nowMonth = 12
else:
    nowMonth = nowtime.month
nowMonth = str(nowMonth).zfill(2)
print("処理中:{}月".format(nowMonth))

root = tk.Tk()
root.withdraw()

pdf_path = filedialog.askopenfilename()
print(pdf_path)

# ============================ 辅助函数 ============================#


def write_to_excel(ws, name, date, time_list):
    """
    立即写入 Excel,一行データ = (名字, 日期, 時間リスト)
    """
    if not time_list:
        print(f"⚠️ 時間リストが空です: {name} - {date}")
        return  # 時間データがない場合は書き込まない

    # から日(DD)を日付から抽出する、例 "02/01" -> 1
    match = re.match(r"\d{2}/(\d{2})", date)
    if not match:
        print(f"⚠️ 日付形式が間違っています: {date}")
        return  # 日付形式が間違っている場合はスキップ

    day = int(match.group(1))
    row_idx = 8 + day  # Excel 行インデックス、例 1日 -> 9行, 2日 -> 10行, ..., 31日 -> 39行

    # X 座標でソート
    time_list.sort(key=lambda x: x[1])

    # 時間文字列を抽出し、ゼロパディングする
    time_strs = [zero_pad_time(x[0]) for x in time_list]

    # Excel に書き込む
    ws[f"A{row_idx}"] = date  # A列 -> 日付
    ws[f"B{row_idx}"] = name  # B列 -> 名前

    # タイムスタンプを Excel に書き込む
    for idx, tm_str in enumerate(time_strs):
        col = chr(ord('C') + idx)  # 列名を計算
        try:
            tobj = datetime.strptime(tm_str, "%H:%M").time()
            ws[f"{col}{row_idx}"] = tobj
        except ValueError:
            ws[f"{col}{row_idx}"] = tm_str  # 変換に失敗した場合は文字列を書き込む

    output_xlsx = "PD2Macro_2025_01.xlsm"
    wb.save(output_xlsx)
    print(f"📄 已写入 Excel: {name} - {date} - {time_strs}")


def full_width_to_half_width(text: str) -> str:
    """将全角字符转换为半角字符 (ここでは主にコロン「：」から「:」への変換を示しています)。"""
    full_width_chars = "："
    half_width_chars = ":"
    trans = str.maketrans(full_width_chars, half_width_chars)
    return text.translate(trans)


def process_date(text: str) -> str:
    """
    '1日'、'14日' などの日付形式をマッチし、'MM/DD' 形式の日付を返します。
    ここで MM は現在の月（2桁）、DD は抽出した日付（2桁）です。
    """
    date_pattern = r"(\d{1,2})日"  # 1-2桁の数字に後続する '日' をマッチ
    match = re.match(date_pattern, text)
    if match:
        day = match.group(1).zfill(2)  # 日付を2桁にゼロパディング
        return f"{nowMonth}/{day}"     # 'MM/DD' 形式を返す
    return text


def parse_times(text: str) -> list:
    """
    テキストから時間を抽出（複数の時間と単一の時間の両方をサポート）、時間文字列のリスト (0~2 個) を返します：
      - "当日0:24当日1:23" -> ["0:24", "1:23"]
      - "前日21:47当日8:06" -> ["21:47", "8:06"]
      - "当日8:06" -> ["8:06"]
      - "関係ないテキスト" -> []
    """
    text = full_width_to_half_width(text)
    pattern = r'(?:当日|前日)?\d{1,2}:\d{2}'
    items = re.findall(pattern, text)
    if len(items) >= 2:
        t1 = re.sub(r'(当日|前日)', '', items[0])
        t2 = re.sub(r'(当日|前日)', '', items[1])
        return [t1, t2]
    elif len(items) == 1:
        t = re.sub(r'(当日|前日)', '', items[0])
        return [t]
    else:
        return []


def zero_pad_time(time_str):
    """
    "6:06" を "06:06" に、"0:15" を "00:15" に変換します。'H:MM' または 'HH:MM' でない場合は元の文字列を返します。
    """
    pattern = r'^(\d{1,2}):(\d{2})$'
    match = re.match(pattern, time_str)
    if match:
        hour, minute = match.groups()
        hour = hour.zfill(2)
        return f"{hour}:{minute}"
    return time_str


# ============================ 主ロジック ============================#

ocr = PaddleOCR(
    use_angle_cls=False,
    lang='japan',
)

wb = openpyxl.load_workbook("PDtestM.xlsm", keep_vba=True)
source = wb["99999　ニッセープロダクツ"]
number_master_sheet = wb['Number_Master']
number_name_dict = {}
for row in range(1, number_master_sheet.max_row + 1):
    PD_number = number_master_sheet.cell(row=row, column=1).value
    staff_name = number_master_sheet.cell(row=row, column=2).value
    NP_number = number_master_sheet.cell(row=row, column=3).value
    temp = {PD_number: (NP_number, staff_name)}
    number_name_dict.update(temp)

name_now = None
current_date = None



with fitz.open(pdf_path) as pdf:
    for page_num in range(pdf.page_count):


        # PDF ページを開く
        page = pdf.load_page(page_num)

        matrix = fitz.Matrix(200 / 72, 200 / 72)
        width, height = page.rect.width, page.rect.height
        crop_rect = fitz.Rect(width * 0.0, height * 0.0, width * 0.6, height * 1)
        page.set_cropbox(crop_rect)

        pix = page.get_pixmap(matrix=matrix)
        img_path = f"output\\PDFToPNG_PAGE_{page_num + 1}.png"
        pix.save(img_path)

        # OCR 認識
        print(f"\n開始OCR Page{page_num + 1}")
        result = ocr.ocr(img_path, cls=False)

        page_rows_data = []  # このページのすべての行データを保存 [(name, date, [times...])]
        # ✅ 毎ページ開始時にのみ初期化、`for` ループ内でクリアしない
        current_line_times = []
        current_line_refY = None
        THRESHOLD_Y = 5  # 許容される上下の誤差

        for i, line in enumerate(result[0]):
            text = line[1][0]
            coords = line[0]
            center_pointX = sum(pt[0] for pt in coords) / 4
            center_pointY = sum(pt[1] for pt in coords) / 4

            print(f"Page {page_num + 1} line {i}: center=({center_pointX:.1f}, {center_pointY:.1f}), text={text}")

            # ✅ `current_line_refY` が空の場合のみ初期化
            if current_line_refY is None:
                current_line_refY = center_pointY
            else:
                # Y 座標の変化がしきい値を超えた場合、改行と判断
                if abs(center_pointY - current_line_refY) > THRESHOLD_Y:
                    if current_line_times:
                        current_line_times.sort(key=lambda x: x[1])  # X 座標でソート
                        page_rows_data.append((name_now, current_date, current_line_times[:]))
                        print(f"改行、前の行は確定: {current_line_times}")
                        print(f"完整信息:", name_now, current_date, current_line_times[:])
                        current_line_times.clear()  # ✅ ここでクリア

                    current_line_refY = center_pointY  # Y 参照値を更新

            # ✅ 名前を解析
            if "NP" in text:
                tmp = text.strip()
                name_now = tmp if tmp else "未識別氏名"

                # 名前の連結処理
                j = i + 1
                while j < len(result[0]):
                    next_text = result[0][j][1][0]
                    next_coords = result[0][j][0]
                    next_centerX = sum(pt[0] for pt in next_coords) / 4
                    next_centerY = sum(pt[1] for pt in next_coords) / 4

                    if abs(next_centerY - center_pointY) < THRESHOLD_Y and next_centerX <= 300:
                        name_now += next_text.strip()
                        j += 1
                    else:
                        break

                try:
                    testnumber = number_name_dict[name_now][0]
                    testname = number_name_dict[name_now][1]
                except KeyError:
                    testnumber = "不明"
                    testname = name_now

                print(f"--- Page {page_num + 1} 収集: {page_rows_data}")
                print(f"現在の current_line_times の値: {current_line_times}")

                # ✅ 前の人のデータを確定
                if page_rows_data:
                    page_rows_data.append((name_now, current_date, current_line_times[:]))
                    print(f"✅ {name_now} のデータを確定: {current_line_times}")
                    write_to_excel(source, name_now, current_date, current_line_times)
                    current_line_times.clear()

                continue  # 次のテキストへ

            else:
                try:
                    name_now = tmp
                except NameError:
                    name_now = "先頭未定"

            # ✅ 日付を解析
            dtmp = process_date(text)
            if dtmp != text:
                current_date = dtmp
                continue

            # ✅ 時間を解析
            parsed = parse_times(text)
            if not parsed:
                continue  # 時間を解析できなかった場合はスキップ

            # ✅ 時間データの処理
            for tm in parsed:
                if current_line_refY is None:
                    current_line_refY = center_pointY
                    current_line_times.append((tm, center_pointX, center_pointY))
                else:
                    if abs(center_pointY - current_line_refY) > THRESHOLD_Y:
                        if current_line_times:
                            page_rows_data.append((name_now, current_date, current_line_times[:]))
                            print(f"改行検出、前の行の時間: {current_line_times}")

                        current_line_times.clear()
                        current_line_refY = center_pointY
                        current_line_times.append((tm, center_pointX, center_pointY))
                    else:
                        current_line_times.append((tm, center_pointX, center_pointY))

        # ✅ 最終行のデータを処理
        if current_line_times:
            page_rows_data.append((name_now, current_date, current_line_times[:]))
            current_line_times.clear()

        print(f"--- Page {page_num + 1} 収集: {page_rows_data}")

        # ✅ Excel に書き込み
        ws = wb.copy_worksheet(source)
        if page_rows_data:
            ws.title = page_rows_data[-1][0] or "結果"
        else:
            ws.title = "結果"

print("全部処理完了。")