from flask import Flask, render_template, request, send_file, jsonify
import easyocr
import pandas as pd
import re
import io
import time  # Thêm thư viện bấm giờ

app = Flask(__name__)

print("Đang khởi động hệ thống AI (EasyOCR)... Vui lòng đợi...")
reader = easyocr.Reader(['en', 'ja'], gpu=True)
print("Sẵn sàng!")

# Biến toàn cục để theo dõi tiến độ xử lý
progress_status = {"current": 0, "total": 0}


@app.route('/')
def index():
    return render_template('index.html')


# API này để màn hình Web liên tục gọi 
@app.route('/progress', methods=['GET'])
def get_progress():
    return jsonify(progress_status)


@app.route('/process_batch', methods=['POST'])
def process_batch():
    global progress_status
    total_pairs = int(request.form.get('total_pairs', 0))

    # Reset tiến độ
    progress_status['total'] = total_pairs
    progress_status['current'] = 0

    results = []

    # BẮT ĐẦU BẤM GIỜ
    start_time = time.time()

    for i in range(total_pairs):
        img1_file = request.files.get(f'pair_{i}_img1')
        img2_file = request.files.get(f'pair_{i}_img2')
        is_notrack = request.form.get(f'pair_{i}_notrack') == 'true'

        row_data = {
            'STT': i + 1,
            'Ảnh Hóa Đơn': img1_file.filename if img1_file else 'Lỗi',
            'Ảnh Tracking': 'notrack' if is_notrack else (img2_file.filename if img2_file else 'Lỗi'),
            'Mã ID (m...)': 'Không tìm thấy',
            'Giá tiền (¥)': 'Không tìm thấy',
            'Mã Tracking': 'notrack' if is_notrack else 'Không tìm thấy',
            'Cảnh báo OCR': 'OK'
        }

        # ==========================================
        # XỬ LÝ ẢNH 1 (HÓA ĐƠN) BẰNG TỌA ĐỘ KHÔNG GIAN
        # ==========================================
        if img1_file:
            img1_bytes = img1_file.read()
            # detail=1 yêu cầu AI trả về [Tọa độ, Chữ, Độ tự tin]
            result_1 = reader.readtext(img1_bytes, detail=1)

            # --- 1. TÌM MÃ ID ---
            # Để lấy nhanh mã ID, ta nối hết text lại như cũ
            full_text_1 = " ".join([item[1] for item in result_1])
            match_id = re.search(r'm\d{11}', full_text_1)
            if match_id:
                row_data['Mã ID (m...)'] = match_id.group(0)
            else:
                row_data['Cảnh báo OCR'] = 'Thiếu mã ID'

            # --- 2. THUẬT TOÁN TỌA ĐỘ TÌM GIÁ TIỀN ---
            anchor_y = -1
            anchor_x_right = -1

            # Bước A: Tìm tọa độ của Mỏ neo "商品代金"
            for bbox, text, prob in result_1:
                if '商品代金' in text.replace(" ", ""):
                    # bbox có dạng: [[x_trái_trên, y_trái_trên], [x_phải_trên, y_phải_trên], ...]
                    anchor_y = (bbox[0][1] + bbox[2][1]) / 2  # Tính tọa độ Y ở giữa chữ
                    anchor_x_right = max(bbox[1][0], bbox[2][0])  # Tọa độ X ngoài cùng bên phải của chữ
                    break

            if anchor_y != -1:
                price_blocks = []

                # Bước B: Tìm tất cả các khối chữ nằm CÙNG DÒNG và BÊN PHẢI mỏ neo
                for bbox, text, prob in result_1:
                    center_y = (bbox[0][1] + bbox[2][1]) / 2
                    center_x = (bbox[0][0] + bbox[1][0]) / 2

                    # Cùng dòng ngang (sai số 25 pixel) và nằm bên phải
                    if abs(center_y - anchor_y) < 25 and center_x > anchor_x_right:
                        # TRỊ BỆNH 1: Sửa lỗi 1,OOO thành 1,000
                        fixed_text = text.replace('O', '0').replace('o', '0').replace('Q', '0')
                        price_blocks.append({
                            'text': fixed_text,
                            'x': bbox[0][0]  # Lưu lại tọa độ X để sắp xếp
                        })

                if price_blocks:
                    # Sắp xếp các khối từ trái sang phải
                    price_blocks.sort(key=lambda b: b['x'])

                    # TRỊ BỆNH 2: Xử lý vụ dính số 1 (Ví dụ AI tách thành 2 khối: ['1', '34,444'])
                    # Nếu khối ngoài cùng bên trái CỰC NGẮN (chỉ là '1', '¥', 'Y') -> XÓA NÓ ĐI
                    first_text = price_blocks[0]['text'].strip()
                    if len(price_blocks) > 1 and (first_text == '1' or first_text in ['¥', '￥', 'Y']):
                        price_blocks = price_blocks[1:]

                    # Nối các khối lại
                    combined_text = "".join([b['text'] for b in price_blocks])

                    # Dọn dẹp ký hiệu thừa và lấy số
                    clean_str = re.sub(r'[¥￥Y\+]', '', combined_text)
                    matches = re.findall(r'[\d,]+', clean_str)

                    for m in matches:
                        clean_m = m.replace(',', '')
                        if clean_m.isdigit() and int(clean_m) >= 300:
                            row_data['Giá tiền (¥)'] = clean_m
                            break

            if row_data['Giá tiền (¥)'] == 'Không tìm thấy':
                row_data['Cảnh báo OCR'] = 'Thiếu giá tiền' if row_data['Cảnh báo OCR'] == 'OK' else row_data[
                                                                                                         'Cảnh báo OCR'] + ' & Thiếu giá tiền'

        # ==========================================
        # XỬ LÝ ẢNH 2 (TRACKING)
        # ==========================================
        if not is_notrack and img2_file:
            img2_bytes = img2_file.read()
            result_2 = reader.readtext(img2_bytes, detail=0)
            full_text_2 = " ".join(result_2)

            match_tracking = re.search(r'\d{4}-?\d{4}-?\d{4}|\d{12}', full_text_2)
            if match_tracking:
                row_data['Mã Tracking'] = match_tracking.group(0).replace('-', '')
            else:
                row_data['Cảnh báo OCR'] = 'Thiếu Tracking' if row_data['Cảnh báo OCR'] == 'OK' else row_data[
                                                                                                         'Cảnh báo OCR'] + ' & Thiếu Tracking'

        results.append(row_data)

        # Cập nhật số lượng ảnh đã xử lý xong để web biết
        progress_status['current'] = i + 1

    # KẾT THÚC BẤM GIỜ VÀ IN RA TERMINAL
    end_time = time.time()
    duration = end_time - start_time
    minutes = int(duration // 60)
    seconds = int(duration % 60)

    print(f"\n=======================================================")
    print(f" [THÀNH CÔNG] Đã xử lý xong {total_pairs} cặp ảnh.")
    if minutes > 0:
        print(f" Thời gian xử lý: {minutes} phút {seconds} giây.")
    else:
        print(f" Thời gian xử lý: {seconds} giây.")
    print(f"=======================================================\n")

    df = pd.DataFrame(results)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ket Qua OCR')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='Ket_qua.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5002, debug=True)