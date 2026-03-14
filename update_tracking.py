import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime

# ================= CẤU HÌNH =================
EXCEL_FILE = r"C:\Users\LINH\Desktop\copy_tracking\ocr11032026.xlsx"  # Tên file excel local của bạn
CREDENTIALS_FILE = 'credentials.json'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1-A-4S3h6GwFEIPKAzueFvAqPKAzueFvAqkksdjjkaakajt7TwpJOHk/edit?gid=0#gid=0' #REAL
TAB_NAME = 'Trang tính1'
OUTPUT_TXT_FILE = 'bao_cao_dien_tracking.txt'


# ============================================

def clean_string(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip().lower()


def clean_price(val):
    """Giữ lại toàn bộ các con số, loại bỏ ¥, dấu phẩy, chấm..."""
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r'[^\d]', '', str(val))


def extract_id_from_url(url):
    """Hàm trích xuất mã ID (bắt đầu bằng 'm' và theo sau là số) từ link URL"""
    url_str = str(url).strip()
    match = re.search(r'(m\d+)', url_str, re.IGNORECASE)
    if match:
        return match.group(1).lower()
    return ""


def main():
    print("1. Đang đọc và xử lý dữ liệu từ file Excel local...")
    try:
        df_excel = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"Lỗi đọc file Excel: {e}")
        return

    # Lấy dữ liệu Excel đưa vào Dictionary để tra cứu
    # Cấu trúc: { 'm44746501488': {'gia': '10500', 'tracking': 'TRACK_ABC'} }
    excel_dict = {}
    for index, row in df_excel.iterrows():
        ma_id = clean_string(row.get('Mã ID (m...)'))
        if ma_id:
            excel_dict[ma_id] = {
                'gia': clean_price(row.get('Giá tiền (¥)')),
                'tracking': str(row.get('Mã Tracking')).strip() if not pd.isna(row.get('Mã Tracking')) else ""
            }

    print("2. Đang kết nối tới Google Sheet...")
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)

    try:
        sheet = client.open_by_url(SHEET_URL).worksheet(TAB_NAME)
        sheet_data = sheet.get_all_values()
    except Exception as e:
        print(f"Lỗi kết nối Google Sheet: {e}")
        return

    if not sheet_data:
        print("Google Sheet trống!")
        return

    # Xác định vị trí cột (Header dòng 1 -> index 0)
    headers = [clean_string(h) for h in sheet_data[0]]
    try:
        col_url_idx = headers.index('link /url')
        col_gia_idx = headers.index('đơn giá')
        col_tracking_idx = headers.index('tracking')
    except ValueError as e:
        print(f"Lỗi: Không tìm thấy cột trong Google Sheet. Vui lòng kiểm tra lại tên cột. Chi tiết: {e}")
        return

    print("3. Đang đối chiếu và chuẩn bị dữ liệu cập nhật...")

    cells_to_update = []

    # Biến lưu trữ báo cáo
    log_all = []
    stats = {'success': [], 'err_price': [], 'not_found': [], 'no_tracking_excel': []}

    # Quét qua Google Sheet từ dòng 2 (index 1)
    for row_idx, row_data in enumerate(sheet_data[1:], start=2):
        # Đảm bảo list row_data đủ dài
        while len(row_data) <= max(col_url_idx, col_gia_idx, col_tracking_idx):
            row_data.append("")

        sheet_url = row_data[col_url_idx]
        sheet_gia = clean_price(row_data[col_gia_idx])

        # Trích xuất Mã ID từ URL trên Sheet
        sheet_ma_id = extract_id_from_url(sheet_url)

        if not sheet_ma_id:
            continue  # Bỏ qua nếu dòng này không có link hoặc link không chứa mã m...

        # KIỂM TRA ĐỐI CHIẾU VỚI EXCEL
        if sheet_ma_id in excel_dict:
            excel_record = excel_dict[sheet_ma_id]
            excel_gia = excel_record['gia']
            excel_tracking = excel_record['tracking']

            if sheet_gia == excel_gia:
                if excel_tracking:
                    # Trùng ID + Trùng Giá -> Ghi Tracking vào Gsheet
                    # gspread column index bắt đầu từ 1, nên lấy index + 1
                    cells_to_update.append(gspread.Cell(row=row_idx, col=col_tracking_idx + 1, value=excel_tracking))

                    msg = f"Mã ID: {sheet_ma_id} | OK -> Đã điền Tracking: {excel_tracking}"
                    log_all.append(msg)
                    stats['success'].append(sheet_ma_id)
                else:
                    msg = f"Mã ID: {sheet_ma_id} | Lỗi: File Excel không có sẵn Mã Tracking để điền"
                    log_all.append(msg)
                    stats['no_tracking_excel'].append(sheet_ma_id)
            else:
                msg = f"Mã ID: {sheet_ma_id} | Lỗi Sai Giá (Excel: {excel_gia} vs Sheet: {sheet_gia})"
                log_all.append(msg)
                stats['err_price'].append(sheet_ma_id)

    # Tìm các mã trong Excel không xuất hiện trên Google Sheet
    sheet_extracted_ids = [extract_id_from_url(row[col_url_idx]) for row in sheet_data[1:] if len(row) > col_url_idx]
    for ma_id in excel_dict.keys():
        if ma_id not in sheet_extracted_ids:
            msg = f"Mã ID: {ma_id} | Lỗi: Không tìm thấy link chứa mã này trên Google Sheet"
            log_all.append(msg)
            stats['not_found'].append(ma_id)

    # Cập nhật hàng loạt lên Google Sheet
    if cells_to_update:
        print(f"4. Đang điền Mã Tracking cho {len(cells_to_update)} đơn hàng lên Google Sheet...")
        sheet.update_cells(cells_to_update)
    else:
        print("4. Không có đơn hàng nào hợp lệ để điền Tracking.")

    # Xuất file báo cáo
    print("5. Đang xuất file Báo Cáo TXT...")
    with open(OUTPUT_TXT_FILE, 'w', encoding='utf-8') as f:
        thoi_gian_chay = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"BÁO CÁO TỰ ĐỘNG ĐIỀN TRACKING ({thoi_gian_chay})\n")
        f.write("=" * 60 + "\n\n")

        f.write("PHẦN 1: NHẬT KÝ CHI TIẾT TỪNG MÃ\n")
        f.write("-" * 60 + "\n")
        for log in log_all:
            f.write(log + "\n")

        f.write("\n\nPHẦN 2: THỐNG KÊ TỔNG HỢP\n")
        f.write("-" * 60 + "\n")

        f.write(f"\n[THÀNH CÔNG] Đã điền Tracking ({len(stats['success'])} đơn):\n")
        for m in stats['success']: f.write(f"  + {m}\n")

        f.write(f"\n[LỖI SAI GIÁ] Khớp ID nhưng sai tiền ({len(stats['err_price'])} đơn):\n")
        for m in stats['err_price']: f.write(f"  + {m}\n")

        f.write(
            f"\n[LỖI THIẾU DỮ LIỆU] Khớp ID, khớp giá nhưng Excel bỏ trống cột Tracking ({len(stats['no_tracking_excel'])} đơn):\n")
        for m in stats['no_tracking_excel']: f.write(f"  + {m}\n")

        f.write(
            f"\n[LỖI KHÔNG TÌM THẤY] Mã có trong Excel nhưng không có link trên GSheet ({len(stats['not_found'])} đơn):\n")
        for m in stats['not_found']: f.write(f"  + {m}\n")

    print(f"=> HOÀN TẤT! Đã lưu kết quả vào file '{OUTPUT_TXT_FILE}'")


if __name__ == "__main__":
    main()