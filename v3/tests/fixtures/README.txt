tests/fixtures/ — Excel fixture files và snapshots

KHÔNG commit file .xlsx lên git (chứa dữ liệu thực tế).
File .json và .txt snapshot được commit để track regressions.

File cần đặt vào thư mục này:
─────────────────────────────────────────────────────────────────────────────
  bom_valid.xlsx     File BOM Excel chuẩn — phải pass validate layer 1 (0 error)
  bom_invalid.xlsx   File BOM Excel lỗi  — phải trigger ≥1 validation error

File do pytest tự tạo (lần đầu chạy):
─────────────────────────────────────────────────────────────────────────────
  row_counts.json    Snapshot số dòng per-section của bom_valid.xlsx
  warning_count.txt  Snapshot số warning của bom_valid.xlsx

Nếu file fixture thay đổi intentionally:
  → Xóa row_counts.json và warning_count.txt rồi chạy lại để regenerate.
