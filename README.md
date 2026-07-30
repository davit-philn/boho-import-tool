# BOHO Import BOM / THDM

Công cụ desktop giúp đọc file Excel BOM (Bill of Materials) và THDM (Tổng Hợp Định Mức Vật Tư), xử lý dữ liệu và import trực tiếp vào SQL Server (hệ thống Bravo ERP).

---

## Tính năng chính

- Đọc file Excel BOM nhiều sheet (A, B1, B2, C, D, E1, E2, F, G, H, I, J)
- Đọc sheet TH VT (Tổng Hợp Vật Tư) — bao gồm expand Mục theo CHI TIẾT GIAO CT/TP
- Preview SQL INSERT trước khi import
- Import trực tiếp vào SQL Server qua kết nối pyodbc
- Mapping cột Excel → SQL cấu hình qua file `CK_Mapping_v5.xlsx` (không hardcode)
- Giao diện Dark / Light theme
- Xuất file `.sql` để kiểm tra offline
- Build thành file `.exe` Portable không cần cài Python

---

## Yêu cầu

- Python 3.10+
- SQL Server với Bravo ERP (cần VPN nếu kết nối remote)
- Windows 10/11

Cài thư viện:

```bash
pip install -r requirements.txt
```

Các thư viện bắt buộc: `pandas`, `customtkinter`, `openpyxl`, `pyodbc`, `rapidfuzz`  
Tuỳ chọn: `tksheet` (bảng nhanh hơn), `msoffcrypto` (mở file Excel có mật khẩu)

---

## Cấu hình kết nối DB

Tạo file `Tools/config/db_config.json` (file này **không** được commit lên git):

```json
{
  "server": "IP_SERVER",
  "database": "TEN_DATABASE",
  "username": "username",
  "password": "password"
}
```

---

## Cách chạy

```bash
cd Tools
python main.py
```

---

## Cấu trúc thư mục

```
Tools/
├── main.py                   # Entry point
├── build.py                  # Script build .exe (PyInstaller)
├── requirements.txt (ở root)
├── config/
│   └── db_config.json        # Credentials DB — KHÔNG commit
├── services/
│   ├── bom_parser.py         # Parser Excel BOM + THDM
│   ├── mapping_loader.py     # Đọc CK_Mapping_v5.xlsx
│   ├── validators.py         # Kiểm tra dữ liệu
│   └── utils.py              # Tiện ích chung
└── views/
    ├── main_window.py        # Giao diện chính
    └── widgets.py            # Widget tuỳ chỉnh

Mapping/
└── CK_Mapping_v5.xlsx        # Cấu hình map cột Excel → SQL

Structure/
└── *.xlsx                    # Schema tham chiếu các bảng SQL

Doc/
└── *.docx                    # Tài liệu thiết kế và hướng dẫn sử dụng
```

---

## Build thành .exe

Yêu cầu: PyInstaller đã cài (`pip install pyinstaller`)

```bash
cd Tools
python build.py
```

Output:
- `dist/BOHO_IMPORT_BOM_THDM/` — chạy trực tiếp
- `installer_output/BOHO_ImportBOM_THDM_v2.1_Portable.zip` — bản Portable
- `installer_output/BOHO_ImportBOM_THDM_Setup_v2.1.exe` — bản Installer (cần Inno Setup)

---

## Lưu ý bảo mật

Các file sau chứa thông tin nhạy cảm và **đã được loại khỏi git** (`.gitignore`):

- `Tools/config/db_config.json` — thông tin kết nối DB
- `Other/acc.txt` — thông tin tài khoản nội bộ

---

## Phiên bản

| Version | Ngày | Ghi chú |
|---------|------|---------|
| v2.1 | 2026-07-29 | Thêm THDM expand Mục, preview SQL, Dark/Light theme |
| v2.0 | 2026-06 | Import BOM đa section, mapping-driven |
