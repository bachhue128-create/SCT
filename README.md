# 🏆 SCT - HỆ THỐNG CẬP NHẬT TỰ ĐỘNG BẢNG VÀNG THI ĐUA SIM FPT 3 TUỔI

Dự án tự động hóa trích xuất dữ liệu thi đua từ file báo cáo Excel nội bộ (`Thi đua Sim`), xếp hạng Top 1 - 5 cho 3 hạng mục và xuất ra ấn phẩm Infographic chất lượng cao (Ultra-HD PNG) chỉ với 1 cú click.

---

## 🚀 Tính Năng Chính
- **Tự động đọc dữ liệu**: Quét và đọc trực tiếp cấu trúc XML của file Excel `.xlsx` mà không cần cài đặt Microsoft Office hoặc thư viện nặng.
- **Tự động xếp hạng**:
  - 👑 **VUA SIM**: Top Shop có sản lượng SIM FPT cao nhất (Cột D).
  - 🎯 **CAO THỦ CHỐT SIM**: Top Shop có tỷ lệ khai thác SIM/Máy cao nhất (Cột E).
  - 📈 **BỨT PHÁ SIM**: Top Shop có tỷ lệ tăng trưởng so với cùng kỳ năm 2025 cao nhất (Cột F).
- **Thiết kế Infographic chuyên nghiệp**: Giao diện chuẩn nhận diện FPT MVNO, bố cục cân đối, màu sắc sinh động, tối ưu hiển thị trên Zalo, Workplace, Email.
- **Tương thích 100% mọi máy tính Windows**: Sử dụng PowerShell và Chrome/Edge headless có sẵn trên Windows.

---

## 📂 Cấu Trúc Thư Mục
```
SCT/
├── CHAY_CAP_NHAT.bat             # [CHÍNH] Nhấp đúp vào đây để cập nhật hàng ngày
├── HUONG_DAN_SU_DUNG.txt         # Hướng dẫn nhanh cho người dùng
├── input_data/                   # Thư mục chứa file báo cáo Excel hàng ngày
│   └── BaoCao_ThiDua_Sim.xlsx    # File báo cáo mẫu
├── output/                       # Thư mục xuất kết quả
│   ├── KetQua_ThiDua_Sim.png     # Ảnh Infographic duy nhất rõ nét nhất (Ultra-HD)
│   └── KetQua_ThiDua_Sim.html    # Trang web Infographic tương tác
├── template/                     # Mẫu thiết kế & tài nguyên hình ảnh
│   ├── assets/                   # Banner Ultra-HD & Mascot SIM FPT
│   └── infographic_template.html # File giao diện HTML/CSS
└── scripts/                      # Bộ mã xử lý tự động hóa
    ├── update_infographic.ps1    # Engine chính bằng PowerShell
    └── update_infographic.py     # Script dự phòng bằng Python
```

---

## 🛠️ Hướng Dẫn Sử Dụng
1. Copy file Excel báo cáo mới vào thư mục `input_data/`.
2. Nhấp đúp vào file `CHAY_CAP_NHAT.bat`.
3. Hệ thống sẽ tự động tính toán, xuất ảnh và mở file `KetQua_ThiDua_Sim.png` lên màn hình.
