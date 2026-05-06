"""
De-identification tool for Taiwan-style ID numbers.
Reads an Excel file produced by taiwan_id_generator.py, adds a
去識別化身分證號碼 column (masking digits 4–7 with ****), and saves a new file.

Usage: python3 taiwan_id_deidentify.py <input.xlsx> [output.xlsx]
"""

import sys
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill


def deidentify(id_str: str) -> str:
    """
    Mask positions 4–7 (0-indexed 3–6) of a Taiwan ID string.
    Example: A123456789 -> A12****789
    """
    if len(id_str) != 10:
        return id_str
    return id_str[:3] + "****" + id_str[7:]


def process_excel(input_path: str, output_path: str) -> None:
    try:
        src = load_workbook(input_path)
    except FileNotFoundError:
        print(f"錯誤：找不到檔案 '{input_path}'")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤：無法開啟檔案 – {e}")
        sys.exit(1)

    ws_src = src.active
    rows = list(ws_src.iter_rows(values_only=True))

    if not rows:
        print("錯誤：來源檔案沒有資料")
        sys.exit(1)

    header = list(rows[0])
    data = rows[1:]

    # Locate the ID column
    id_col_index = None
    for i, col_name in enumerate(header):
        if col_name == "身分證號碼":
            id_col_index = i
            break

    if id_col_index is None:
        print("錯誤：找不到「身分證號碼」欄位，請確認輸入檔案格式正確")
        sys.exit(1)

    wb = Workbook()
    ws = wb.active
    ws.title = "去識別化身分證號碼"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    new_col_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    center = Alignment(horizontal="center", vertical="center")
    alt_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    alt_new_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    new_header = header + ["去識別化身分證號碼"]
    col_widths = [12] * len(header) + [22]

    for col_idx, (title, width) in enumerate(zip(new_header, col_widths), start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font = header_font
        cell.fill = new_col_fill if col_idx == len(new_header) else header_fill
        cell.alignment = center

    for row_idx, row in enumerate(data, start=2):
        is_alt = row_idx % 2 == 0
        id_value = row[id_col_index] if id_col_index < len(row) else ""
        masked = deidentify(str(id_value)) if id_value else ""

        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = center
            if is_alt:
                cell.fill = alt_fill

        new_col_idx = len(row) + 1
        cell = ws.cell(row=row_idx, column=new_col_idx, value=masked)
        cell.alignment = center
        if is_alt:
            cell.fill = alt_new_fill

    ws.freeze_panes = "A2"
    wb.save(output_path)
    print(f"完成：共處理 {len(data)} 筆，已儲存至 {output_path}")
    print(f"去識別化規則：保留前 3 碼與後 3 碼，中間 4 碼遮蔽為 ****")
    print(f"範例：A123456789  →  A12****789")


def main():
    if len(sys.argv) < 2:
        print("用法: python3 taiwan_id_deidentify.py <輸入檔案.xlsx> [輸出檔案.xlsx]")
        print("例如: python3 taiwan_id_deidentify.py taiwan_ids.xlsx deidentified.xlsx")
        sys.exit(1)

    input_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base = input_path.removesuffix(".xlsx")
        output_path = f"{base}_deidentified.xlsx"

    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"

    process_excel(input_path, output_path)


if __name__ == "__main__":
    main()
