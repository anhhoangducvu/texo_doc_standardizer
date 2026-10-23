# 📄 Chuẩn hóa thể thức văn bản – TEXO (v2.0)

App Streamlit đưa văn bản Word (.docx) về đúng thể thức theo **Quy định TEXO** (giấy thường hoặc Letterhead) hoặc **Nghị định 30/2020/NĐ-CP**. Chạy cục bộ, không gửi dữ liệu ra ngoài.

## Bộ quy chuẩn

| | Quy định TEXO | Nghị định 30/2020 |
|---|---|---|
| Khổ giấy | A4 (giữ trang ngang nếu văn bản có) | A4 |
| Lề T/D/T/P | 20/20/30/20 mm · Letterhead 40/30/30/20 mm | 20/20/30/15 mm |
| Phông | Times New Roman (đặt cả bộ phông Unicode của Word) | Times New Roman |
| Cỡ chữ | nội dung 13 · tiêu đề căn giữa in hoa 14 · dòng in hoa khác 12 · bảng 12 | nội dung 14 · dòng in hoa 13 · bảng 13 |
| Giãn dòng | Exactly 17 pt · trong bảng Exactly 15 pt | 1,15 lần · trong bảng 1 lần |
| Cách đoạn | trước 6 pt, sau 3 pt | trước 6 pt |
| Khác | lề trong ô bảng 2 mm | thụt đầu dòng 1,27 cm cho đoạn nội dung |

Thông số nằm trong `PROFILES` của `core/engine.py`; danh sách lỗi gõ nằm trong `SAFE_FIXES`. Muốn đổi, anh Vũ nhắn em để cập nhật và ghi lại vào mục *Nhật ký phiên bản*.

## Điểm mới bản 2.0 (so với bản 1)

- **Giữ căn giữa / căn phải**: bản 1 ép mọi đoạn không có căn giữa trực tiếp thành căn đều, nên tiêu đề căn giữa theo Style bị lệch. Bản 2 đọc căn lề theo cả Style.
- **Bảo vệ hình ảnh**: đoạn có ảnh không bị đặt giãn dòng Exactly, nên ảnh không bị cắt mất.
- **Ngày tháng đúng NĐ 30**: chỉ thêm số 0 cho ngày dưới 10 và tháng 1, 2, và chỉ trong cụm "ngày … tháng … năm …". Bản 1 thêm số 0 cho mọi tháng, kể cả trong "3 tháng 5 ngày", "ngày 1 lần", và đổi ngày của văn bản viện dẫn (18/8/2020 thành 18/08/2020).
- **Bỏ quy tắc dễ sai**:
  - Thêm dấu phẩy sau "Hiện nay / Đến nay" ở giữa câu.
  - Đổi "mươi bốn" thành "mươi tư".
  - Dấu phẩy sau liên từ giờ chỉ thêm khi liên từ đứng đầu câu.
- **Lề ô bảng hợp lý**: không thêm lề cho cột STT ngắn và khối Quốc hiệu căn giữa (bản 1 làm "TT" vỡ thành 2 dòng).
- **Phông đầy đủ**: đặt Times New Roman cho cả Style Normal, header, footer.
- **Nhật ký thay đổi**:
  - Mỗi file ghi lề cũ, số đoạn căn đều, số đoạn giữ căn, số đoạn có ảnh, và từng chỗ sửa chữ (trước → sau).
  - Nhật ký xem trên màn hình hoặc tải Excel; file Excel này cũng có trong gói ZIP.
- **Tùy chọn bật/tắt** từng loại sửa chữ ở thanh bên.
- **Giao diện** đồng bộ bộ app Phòng Kỹ thuật:
  - Đăng nhập có logo TEXO (mật khẩu mặc định `texo2026`, đổi bằng `APP_PASSWORD` trong Secrets).
  - Đầu trang logo kèm tiêu đề, ô KPI, giao diện sáng cố định.

## Chạy

```bash
pip install -r requirements.txt
streamlit run app.py
```

Dòng lệnh (giữ như bản 1):
```bash
python -m core.internal_standard "file.docx" [--letterhead]
python -m core.gov_standard "file.docx"
```

## Cấu trúc

```
app.py                    Giao diện
core/engine.py            Bộ máy chuẩn hóa + bộ quy chuẩn + nhật ký
core/internal_standard.py Hàm cũ apply_texo_internal_standard (gọi engine)
core/gov_standard.py      Hàm cũ apply_nd30_standard (gọi engine)
core/ui.py                Đăng nhập, đầu trang, ô KPI dùng chung bộ app P.KT
assets/logo_texo.png      Logo
.streamlit/config.toml    Giao diện sáng, giới hạn tải 100 MB
```

## Giới hạn

- Chỉ nhận `.docx`. File `.doc` cũ cần mở bằng Word và *Lưu thành* `.docx`.
- App không tự co bảng rộng hơn khổ giấy và không chỉnh ngắt trang. Sau khi chuẩn hóa, mở Word xem lại bảng và khối chữ ký.
- Sửa chữ chỉ làm trong từng đoạn định dạng (run). Lỗi gõ bị tách giữa hai đoạn định dạng sẽ không bắt được.

## Nhật ký phiên bản

- **2.0 (23/10/2026)**: bộ máy mới `core/engine.py`, sửa các lỗi căn lề, ảnh, ngày tháng của bản 1; thêm nhật ký thay đổi, tùy chọn sửa chữ, giao diện đồng bộ.
- **1.0 (03/2026)**: bản đầu.

_Người xây dựng: Hoàng Đức Vũ – Phòng Kỹ thuật._
