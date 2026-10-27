# 🇻🇳 VietExcel FontFixer

> **Công cụ sửa lỗi font VNI-Windows & TCVN3 (ABC) cho file Excel (.xls, .xlsx) sang Unicode chuẩn — Giữ nguyên 100% định dạng, viền ô, căn lề và công thức.**

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()

---

## 📌 Giới thiệu

Bạn có những file Excel cũ (thập niên 2000–2015) dùng font **`VNI-Times`**, **`VNI-Helve`** hoặc **`.VnTime`**, **`.VnArial`** mà khi mở trên máy tính hiện đại hay Google Sheets bị lỗi hiển thị chữ loằng ngoằng (ví dụ: `Nöôùc`, `nhaän`, `khoâng`...)?

Nếu dùng Unikey để chuyển mã, bạn sẽ phải copy từng ô và **bị mất sạch toàn bộ kẻ bảng (borders), màu sắc, chiều cao dòng, độ rộng cột và ô gộp (merged cells)**.

**VietExcel FontFixer** ra đời để giải quyết triệt để vấn đề này:
* 🎯 **Chuyển mã thông minh:** Tự động nhận diện và chuyển từ **VNI-Windows** và **TCVN3 (ABC)** sang **Unicode chuẩn** (UTF-8).
* 🎨 **Bảo toàn 100% định dạng:** Giữ nguyên viền ô kẻ bảng, chiều cao dòng, độ rộng cột, font in đậm/in nghiêng, căn chỉnh lề (trái, phải, giữa, ngắt dòng) và các ô được gộp lại (`merged cells`).
* 🔄 **Tự động đổi font:** Đổi font cũ sang font hiện đại tương ứng (`VNI-Times` / `.VnTime` ➔ `Times New Roman`, `VNI-Helve` / `.VnArial` ➔ `Arial`).
* 📁 **Xử lý hàng loạt:** Chọn nhiều file hoặc chọn cả một thư mục gồm hàng chục file để xử lý cùng lúc.
* 🔒 **Bảo mật tuyệt đối:** Chạy hoàn toàn offline trên máy tính cá nhân, không tải dữ liệu lên bất kỳ máy chủ nào.

---

## 🚀 Tải về sử dụng ngay (Không cần cài Python)

1. Tải file **`VietExcel_FontFixer.exe`** mới nhất tại mục [Releases](https://github.com/tuangato147/VietExcel-FontFixer/releases).
2. Nhấp đúp chuột để mở ứng dụng.
3. Kéo/chọn các file Excel cần sửa ➔ Bấm **"BẮT ĐẦU CHUYỂN ĐỔI"**.
4. File kết quả sẽ được tạo ra với đuôi `_fixed.xlsx` ngay cạnh file gốc.

---

## 🛠️ Chạy từ mã nguồn (Dành cho lập trình viên)

### 1. Cài đặt môi trường
Yêu cầu Python 3.8 trở lên. Cài đặt các thư viện phụ thuộc:
```bash
pip install xlrd openpyxl customtkinter pyinstaller
```

### 2. Khởi chạy giao diện ứng dụng
```bash
python app_gui.py
```

### 3. Đóng gói thành file `.exe` duy nhất
Chạy lệnh sau để tự động tạo file thực thi độc lập:
```bash
python build_exe.py
```
File `.exe` thành phẩm sẽ nằm trong thư mục `dist/VietExcel_FontFixer.exe`.

---

## 📖 Cấu trúc dự án

```text
├── viet_converter.py     # Module bảng mã ánh xạ VNI & TCVN3 sang Unicode + Regex tối ưu
├── excel_processor.py    # Engine đọc & tái tạo file Excel (.xls, .xlsx) giữ nguyên style
├── app_gui.py            # Giao diện đồ họa hiện đại xây dựng bằng CustomTkinter
├── build_exe.py          # Script tự động đóng gói ứng dụng bằng PyInstaller
├── README.md             # Hướng dẫn sử dụng & tài liệu dự án
└── .gitignore            # Cấu hình bỏ qua các file tạm của Git
```

---

## 🤝 Đóng góp (Contributing)

Mọi đóng góp nhằm cải thiện bảng mã hoặc hỗ trợ thêm các định dạng khác (như file Word `.doc`/`.docx`) đều rất được hoan nghênh. Xin vui lòng:
1. Fork dự án
2. Tạo nhánh tính năng (`git checkout -b feature/tinh-nang-moi`)
3. Commit thay đổi (`git commit -m 'Thêm tính năng mới'`)
4. Push lên nhánh (`git push origin feature/tinh-nang-moi`)
5. Tạo Pull Request

---

## 📜 Giấy phép

Phát hành dưới giấy phép [MIT License](LICENSE). Tự do sử dụng cho mục đích cá nhân và thương mại.
