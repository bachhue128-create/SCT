# 🏆 SCT - HỆ THỐNG CẬP NHẬT TỰ ĐỘNG BẢNG VÀNG THI ĐUA SIM FPT 3 TUỔI

Hệ thống tự động hóa trích xuất dữ liệu thi đua từ file báo cáo Excel nội bộ (`Thi đua Sim`), xếp hạng Top 1 - 5 cho 3 hạng mục và xuất ra ấn phẩm Infographic chất lượng cao (Ultra-HD PNG) chỉ với 1 cú click.

---

## ⚡ Các File Điều Khiển Chính (Thư mục gốc)
| File | Chức năng | Hướng dẫn |
| :--- | :--- | :--- |
| 🚀 **`CHAY_CAP_NHAT.bat`** | **Cập nhật số liệu & Tạo ảnh Infographic** | Nhấp đúp chuột sau khi thả file Excel mới vào `input_data/` |
| ⬆️ **`KEO_LEN.bat`** | **Đẩy code & dữ liệu lên GitHub** | Nhấp đúp chuột để đồng bộ thay đổi từ máy này lên GitHub |
| ⬇️ **`KEO_VE.bat`** | **Kéo code & dữ liệu mới nhất từ GitHub** | Nhấp đúp chuột để tải bản mới nhất từ GitHub về máy hiện tại |

---

## 🔄 Hướng Dẫn Sử Dụng Giữa 2 Máy Tính (Máy 1 & Máy 2)

### 1. Cài đặt lần đầu trên Máy thứ 2:
1. Mở trình duyệt trên Máy 2, truy cập: [https://github.com/bachhue128-create/SCT](https://github.com/bachhue128-create/SCT)
2. Bấm nút **Code** màu xanh $\rightarrow$ Chọn **Download ZIP**.
3. Giải nén file ZIP ra màn hình Desktop (đặt tên thư mục là `SCT` hoặc `THI_DUA_SIM_FPT`).
4. Xong! Toàn bộ file chạy và dữ liệu đã sẵn sàng.

### 2. Quy trình làm việc và đồng bộ hàng ngày:
```mermaid
flowchart LR
    A["💻 Máy 1<br>(Cập nhật Excel & Ảnh)"] -->|"Bấm KEO_LEN.bat"| B["🌐 GitHub Repository<br>(bachhue128-create/SCT)"]
    B -->|"Bấm KEO_VE.bat"| C["💻 Máy 2<br>(Nhận dữ liệu mới nhất)"]
    C -->|"Bấm KEO_LEN.bat"| B
    B -->|"Bấm KEO_VE.bat"| A
```

- Khi bạn vừa cập nhật số liệu trên Máy 1 xong: Nhấp đúp **`KEO_LEN.bat`** trên Máy 1.
- Khi sang Máy 2 ngồi làm việc: Nhấp đúp **`KEO_VE.bat`** trên Máy 2 là toàn bộ số liệu và ảnh mới nhất sẽ tự động được kéo về!
*(Không cần cài đặt Git hay phần mềm lập trình nào, mọi máy tính Windows đều chạy được trực tiếp)*

---

## 📂 Cấu Trúc Dự Án
```
SCT/
├── CHAY_CAP_NHAT.bat             # [CHÍNH] Nhấp đúp vào đây để cập nhật hàng ngày
├── KEO_LEN.bat                   # [ĐẨY LÊN] Đồng bộ code/dữ liệu lên GitHub
├── KEO_VE.bat                    # [KÉO VỀ] Tải bản cập nhật mới nhất từ GitHub về máy
├── HUONG_DAN_SU_DUNG.txt         # Hướng dẫn nhanh cho người dùng
├── README.md                     # Tài liệu dự án
├── input_data/                   # Thư mục chứa file báo cáo Excel hàng ngày
│   └── BaoCao_ThiDua_Sim.xlsx    # File báo cáo mẫu
├── output/                       # Thư mục xuất kết quả
│   ├── KetQua_ThiDua_Sim.png     # Ảnh Infographic duy nhất rõ nét nhất (Ultra-HD)
│   └── KetQua_ThiDua_Sim.html    # Trang web Infographic tương tác
├── template/                     # Mẫu thiết kế & tài nguyên hình ảnh Ultra-HD
└── scripts/                      # Bộ mã tự động hóa PowerShell
    ├── update_infographic.ps1    # Engine xử lý Excel & render headless
    ├── keo_len.ps1               # Engine đẩy code lên GitHub qua API
    └── keo_ve.ps1                # Engine kéo code từ GitHub về máy
```

---

## 🎯 3 Hạng Mục Thi Đua
1. 👑 **VUA SIM**: Top Shop bán nhiều SIM FPT nhất (Cột D).
2. 🎯 **CAO THỦ CHỐT SIM**: Top Shop có tỷ lệ khai thác SIM/Máy cao nhất (Cột E).
3. 📈 **BỨT PHÁ SIM**: Top Shop có tỷ lệ tăng trưởng so với cùng kỳ năm 2025 cao nhất (Cột F).
