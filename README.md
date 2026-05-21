# ⚡ HỆ THỐNG TẠO THÔNG BÁO TẠM NGỪNG CUNG CẤP ĐIỆN

Ứng dụng web Python dành cho **Tổng công ty Điện lực miền Bắc (EVNNPC)** – tự động tạo Thông báo cắt điện từ file Excel khách hàng và file Word mẫu.

---

## 📦 1. CÀI ĐẶT

### Bước 1 – Cài Python 3.9+
Tải Python từ https://www.python.org/

### Bước 2 – Cài thư viện
Mở terminal/cmd tại thư mục chứa `app.py` rồi chạy:

```bash
pip install -r requirements.txt
```

### Bước 3 – Chạy ứng dụng
```bash
streamlit run app.py
```

Trình duyệt sẽ tự mở tại địa chỉ `http://localhost:8501`.

---

## 📊 2. CẤU TRÚC FILE EXCEL DỮ LIỆU KHÁCH HÀNG

File Excel cần chứa các cột sau (hệ thống tự nhận diện tên cột, **không phân biệt hoa/thường**):

| Cột bắt buộc | Tên gợi ý chấp nhận                                | Ví dụ giá trị         |
|--------------|----------------------------------------------------|----------------------|
| **Lộ đường dây** ✅ | `Lộ`, `Lộ đường dây`, `Lo`, `Feeder`, `Đường dây` | `371E9.12`           |
| **TBA** ✅          | `TBA`, `Trạm biến áp`, `Tên TBA`, `Tram`        | `TBA Đồng Tâm 1`     |
| **Xã/Phường** ✅    | `Xã`, `Phường`, `Xã/Phường`, `Xa`               | `Xã Hòa Bình`        |
| **Thôn/Xóm**        | `Thôn`, `Xóm`, `Thôn/Xóm`                       | `Thôn Đồng Tâm`      |
| **Huyện** (tùy chọn)| `Huyện`, `Quận`                                  | `Huyện Yên Phong`    |
| Số khách hàng (tùy chọn) | `Số KH`                                      | `120`                |

### Ví dụ mẫu Excel:

| STT | Lộ đường dây | TBA              | Thôn         | Xã          | Huyện         |
|-----|--------------|------------------|--------------|-------------|---------------|
| 1   | 371E9.12     | TBA Đồng Tâm 1   | Đồng Tâm     | Hòa Bình    | Yên Phong     |
| 2   | 371E9.12     | TBA Đồng Tâm 1   | Đồng Lực     | Hòa Bình    | Yên Phong     |
| 3   | 371E9.12     | TBA Đồng Tâm 2   | Trung Nghĩa  | Hòa Bình    | Yên Phong     |
| 4   | 371E9.12     | TBA Bắc Sơn      | Bắc Sơn      | Bắc Sơn     | Yên Phong     |
| 5   | 373E1.5      | TBA Tân Lập      | Tân Lập      | Tân Lập     | Yên Phong     |

**Ghi chú:**
- Hệ thống tự **gom các dòng trùng** (thôn, xã, TBA trùng nhau sẽ bị gộp).
- Các thôn cùng xã sẽ được gộp thành: *"các thôn A, B, C – xã X – huyện Y"*.

---

## 📄 3. CẤU TRÚC FILE WORD MẪU (.docx)

File Word mẫu **giữ nguyên định dạng EVN** mà bạn đang sử dụng (logo, font, lề, bảng biểu...).
Chỉ cần chèn các **placeholder** dưới đây vào đúng vị trí:

| Placeholder         | Sẽ được thay bằng                                | Ví dụ                                     |
|---------------------|--------------------------------------------------|-------------------------------------------|
| `{{NGAY_CAT}}`      | Ngày cắt điện                                    | `ngày 25 tháng 05 năm 2026`               |
| `{{GIO_BAT_DAU}}`   | Giờ bắt đầu                                      | `07h00`                                    |
| `{{GIO_KET_THUC}}`  | Giờ kết thúc                                     | `17h00`                                    |
| `{{LO_DUONG_DAY}}`  | Tên lộ đường dây (gom không trùng)              | `371E9.12`                                 |
| `{{DS_TRAM}}`       | Danh sách trạm biến áp                          | `TBA Đồng Tâm 1; TBA Đồng Tâm 2`           |
| `{{DS_THON_XA}}`    | Danh sách khu vực                               | `các thôn Đồng Tâm, Đồng Lực – xã Hòa Bình`|
| `{{LY_DO}}`         | Lý do công tác                                   | `Sửa chữa, bảo dưỡng định kỳ lưới điện`    |

### Ví dụ một đoạn trong template Word mẫu:

```
TỔNG CÔNG TY ĐIỆN LỰC MIỀN BẮC          CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
   CÔNG TY ĐIỆN LỰC ABC                       Độc lập – Tự do – Hạnh phúc

                                            ........., {{NGAY_CAT}}

                                THÔNG BÁO
                       V/v tạm ngừng cung cấp điện

Kính gửi: Quý khách hàng sử dụng điện.

  Thực hiện kế hoạch công tác trên lưới điện nhằm đảm bảo vận hành an toàn,
ổn định và nâng cao chất lượng cung cấp điện, Công ty Điện lực trân trọng
thông báo lịch tạm ngừng cung cấp điện như sau:

1. Thời gian tạm ngừng cấp điện:
   Từ {{GIO_BAT_DAU}} đến {{GIO_KET_THUC}} {{NGAY_CAT}}.

2. Lộ đường dây liên quan: {{LO_DUONG_DAY}}

3. Trạm biến áp bị ảnh hưởng: {{DS_TRAM}}

4. Khu vực mất điện: {{DS_THON_XA}}.

5. Lý do: {{LY_DO}}.

  Công ty Điện lực rất mong Quý khách hàng thông cảm và chủ động bố trí
công việc, sinh hoạt, sản xuất phù hợp trong thời gian nêu trên.

  Mọi thông tin chi tiết, Quý khách hàng vui lòng liên hệ Tổng đài CSKH
19006769 hoặc website cskh.npc.com.vn để được hỗ trợ.

                                                Trân trọng thông báo!

Nơi nhận:                                            GIÁM ĐỐC
- Như kính gửi;
- Lưu: VT, KD.
```

### ⚠️ Lưu ý quan trọng khi soạn template Word

1. **Gõ placeholder liền mạch**: gõ `{{NGAY_CAT}}` trong **một lần** (không Bold/Italic/đổi font giữa các ký tự).
   Nếu Word chia placeholder thành nhiều "run" định dạng khác nhau, có thể bị lỗi không thay được.
   
   ✅ **Cách kiểm tra**: chọn toàn bộ placeholder → chỉ thấy 1 định dạng đồng nhất là OK.

2. **Hoa thường phải đúng**: `{{NGAY_CAT}}` ≠ `{{ngay_cat}}` ≠ `{{Ngay_Cat}}`.

3. **Font Times New Roman**: hệ thống tự đảm bảo font Times New Roman cho toàn bộ document sau khi xuất.

---

## 🚀 4. HƯỚNG DẪN SỬ DỤNG

### Quy trình 4 bước:

1. **Upload file Excel** dữ liệu khách hàng (sidebar bên trái)
2. **Upload file Word mẫu** đã chèn placeholder (sidebar bên trái)
3. **Nhập thông tin** thông báo:
   - 📅 Ngày cắt điện
   - 🕐 Giờ bắt đầu / kết thúc
   - 🔌 Chọn tìm theo **Lộ đường dây** *hoặc* **Trạm biến áp**
   - 📝 Lý do công tác
4. **Bấm nút [⚡ TẠO THÔNG BÁO ⚡]** → Xem preview → **Tải file Word**

### Logic xử lý dữ liệu:

- **Khi nhập tên Lộ** (vd: `371E9.12`):
  - → Tự tìm **tất cả TBA** thuộc lộ này
  - → Tự tìm **tất cả thôn/xã** sử dụng điện từ các TBA đó
  - → Gom danh sách không trùng

- **Khi nhập tên TBA** (vd: `TBA Đồng Tâm`):
  - → Tự tìm **lộ đường dây** chứa TBA này
  - → Tự tìm **các thôn/xã** mà TBA này cấp điện
  - → Gom danh sách không trùng

- **Tìm kiếm chứa từ khóa**: nhập `Đồng Tâm` sẽ tìm cả `TBA Đồng Tâm 1`, `TBA Đồng Tâm 2`...

---

## 🛠️ 5. CẤU TRÚC THƯ MỤC

```
📁 thong-bao-cat-dien/
├── 📄 app.py              # Source code chính
├── 📄 requirements.txt    # Danh sách thư viện
├── 📄 README.md           # File này
├── 📊 data.xlsx           # (Tùy chọn) File Excel mẫu của bạn
└── 📄 template.docx       # (Tùy chọn) File Word mẫu của bạn
```

---

## 📚 6. CÔNG NGHỆ SỬ DỤNG

| Thư viện       | Mục đích                                  |
|----------------|-------------------------------------------|
| `streamlit`    | Tạo giao diện web                        |
| `pandas`       | Đọc và xử lý dữ liệu Excel               |
| `python-docx`  | Đọc/ghi file Word, thay placeholder      |
| `openpyxl`     | Engine đọc file .xlsx                    |
| `xlrd`         | Engine đọc file .xls (cũ)                |

---

## ❓ 7. KHẮC PHỤC SỰ CỐ

| Triệu chứng                              | Giải pháp                                                    |
|------------------------------------------|--------------------------------------------------------------|
| `Thiếu cột bắt buộc trong Excel`         | Đổi tên cột theo gợi ý ở mục 2                              |
| `Placeholder không được thay`            | Gõ lại placeholder trong Word, đảm bảo định dạng đồng nhất  |
| Không tìm thấy dữ liệu                   | Kiểm tra chính tả tên lộ/TBA, có khoảng trắng thừa không   |
| Tiếng Việt bị lỗi font                   | Đảm bảo Excel/Word lưu UTF-8                                |
| Cài đặt báo lỗi `pip`                    | Nâng cấp: `python -m pip install --upgrade pip`             |

---

## 📞 LIÊN HỆ

Mọi góp ý vui lòng phản hồi tới bộ phận phát triển ứng dụng.

> *Phát triển dành cho ngành Điện lực Việt Nam – EVN/EVNNPC*
