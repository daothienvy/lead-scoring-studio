---
name: monthly-operations-report
description: >-
  Automate monthly ERP operational data processing, data cleaning, Net Salary calculations,
  anomaly audit checks, manager file splitting (2-sheet Excel workbooks), Master Dashboard generation,
  and executive operations reporting organized by date subfolders (reports/D-M-YYYY/).
  Use when the user requests processing monthly operations data, generating manager payroll/KPI files,
  handling ERP exports, or creating executive operational performance reports.
category: workflow
keywords: [operations, erp, payroll, monthly-report, excel, dashboard, manager-split, anomaly-audit, pdca]
metadata:
  course: "Agentic AI with Google Antigravity"
  version: "1.0.0"
---

# Monthly Operations & Executive Reporting Skill

Kỹ năng tự động hóa toàn diện quy trình xử lý dữ liệu vận hành hàng tháng từ hệ thống ERP cho **Operation Analyst**. Kỹ năng này thay thế 100% thao tác thủ công: làm sạch dữ liệu, lọc phòng ban, tách file Excel cho từng Manager, kiểm toán số liệu bất thường, vẽ biểu đồ dashboard và soạn thảo báo cáo quản trị cấp cao.

---

## 1. Nguyên Tắc Cốt Lõi (Core Principles)

1. **Chuẩn hóa đầu vào (Single Source of Truth):** Nhận file export ERP trực tiếp qua link Google Sheets công khai hoặc file CSV/Excel cục bộ. Luôn tạo snapshot dữ liệu gốc vào `sample-data/` trước khi xử lý.
2. **Bảo mật & Tách quyền phân tán:** Mỗi Manager chỉ nhận 1 file Excel duy nhất chứa danh sách nhân sự trực thuộc, chia thành 2 Sheet chuyên nghiệp (Tổng quan KPI & Chi tiết nhân sự có công thức tính động).
3. **Kiểm toán dữ liệu chủ động (Data Quality & Anomaly Audit):** Tự động phát hiện vi phạm phạt cao, thưởng vượt khung, hoặc ca bất thường kép (Dual Anomaly) để báo cáo Ban Giám đốc trước khi giải ngân.
4. **Tổ chức báo cáo theo ngày (`outputs/reports/<D-M-YYYY>/`):** Tránh lưu trữ lộn xộn, mọi báo cáo tổng hợp và dashboard phải được lưu vào thư mục ngày tương ứng (ví dụ: `outputs/reports/9-9-2026/`).
5. **Tuân thủ quy tắc Workspace:** Ghi nhận chu trình cải tiến vào `docs/pdca-log.md` và `docs/lesson-checkpoint-log.md`.

---

## 2. Quy Trình Thực Hiện Từng Bước (Step-by-Step Runbook)

### Bước 1: Tiếp nhận và xác thực dữ liệu nguồn
- Kiểm tra nguồn dữ liệu từ người dùng (URL Google Sheet hoặc đường dẫn file `.csv`/`.xlsx`).
- Định dạng bắt buộc gồm 8 cột chuẩn:
  `Employee_ID`, `Employee_Name`, `Manager`, `Department`, `Base_Salary`, `Bonus`, `Penalty`, `Month`.
- Tham khảo chi tiết tại [references/data-schema.md](./references/data-schema.md).

### Bước 2: Chạy công cụ tự động hóa `process_operations.py`
Sử dụng script Python được đóng gói sẵn trong workspace:

```bash
# Lệnh chạy mặc định với dữ liệu Google Sheet hiện tại:
python3 process_operations.py

# Hoặc truyền URL/File tùy biến cùng ngày báo cáo:
python3 process_operations.py --input "<URL_HOAC_FILE_PATH>" --month "2026-03" --date "9-9-2026"
```

**Các tham số tùy chọn nâng cao:**
- `--threshold-penalty`: Ngưỡng cảnh báo phạt cao (mặc định: `1500000` VNĐ).
- `--threshold-bonus`: Ngưỡng cảnh báo thưởng xuất sắc (mặc định: `4000000` VNĐ).
- `--date`: Chỉ định tên thư mục ngày lưu trữ trong `reports/` (mặc định: ngày hiện tại `D-M-YYYY`).

### Bước 3: Kiểm tra các sản phẩm đầu ra (Outputs Verification)
Đảm bảo tất cả tệp sau được tạo lập thành công và đầy đủ dữ liệu:

1. **Bản sao lưu raw:** `sample-data/ERP_Operations_Raw_<month>.csv`
2. **File Excel cho 4 Manager:** `outputs/managers/<month>/Manager_<X>_Operations_<month>.xlsx`
   - Kiểm tra Sheet 1: Thẻ KPI, bảng phân bổ theo phòng ban, danh sách nhân sự cần lưu ý.
   - Kiểm tra Sheet 2: Danh sách chi tiết nhân viên, công thức `=E+F-G`, format tiền tệ `#,##0 "₫"`, công thức dòng Total `=SUM(...)`.
3. **Master Executive Dashboard:** `outputs/reports/<D-M-YYYY>/Operations_Master_Dashboard_<month>.xlsx`
   - Sheet 1: Dashboard KPI cards, Pivot Phòng ban & Quản lý, 2 biểu đồ cột động OpenPyXL.
   - Sheet 2: Dữ liệu sạch toàn bộ nhân sự.
   - Sheet 3: Danh sách kiểm toán bất thường chi tiết (Audit Log).
4. **Báo cáo điều hành Executive Report:** `outputs/reports/<D-M-YYYY>/Operations_Executive_Report_<month>.md`
   - Báo cáo phân tích chuyên sâu cho COO/Ban Giám đốc.

### Bước 4: Đánh giá kiểm toán và các phát hiện vận hành (Insights Extraction)
- Tra cứu danh mục bất thường theo các tiêu chí tại [references/audit-rules.md](./references/audit-rules.md).
- Rà soát các chỉ số:
  - Tỷ lệ quỹ thưởng (`Total_Bonus / Total_Base * 100`) — ngưỡng chuẩn < 15%.
  - Tỷ lệ phạt vi phạm (`Total_Penalty / Total_Base * 100`) — phòng ban nào cao bất thường?
  - Phân bổ quản lý (Span of Control) — phát hiện nếu có Manager gánh trên 30% nhân sự toàn công ty.

### Bước 5: Cập nhật tài liệu Workspace (PDCA Compliance)
- Mở `docs/pdca-log.md` và ghi nhận một entry PDCA mới (Plan - Do - Check - Act).
- Mở `docs/lesson-checkpoint-log.md` để đánh dấu các tác vụ đã hoàn thành.

---

## 3. Cấu Trúc Thư Mục Tài Liệu Đi Kèm

| Loại tài liệu | Đường dẫn liên kết | Mục đích sử dụng |
|:---|:---|:---|
| **Kịch bản chính** | [`process_operations.py`](file:///Users/ThienTuCorp/Desktop/Agenetic%20AI%202026/my-workspace/process_operations.py) | Script CLI tự động hóa toàn bộ luồng xử lý |
| **Data Schema** | [references/data-schema.md](./references/data-schema.md) | Đặc tả cấu trúc cột dữ liệu ERP và công thức tính |
| **Audit Rules** | [references/audit-rules.md](./references/audit-rules.md) | Quy chuẩn kiểm toán bất thường và phân loại mức độ rủi ro |

---

## 4. Xử Lý Tình Huống Ngoại Lệ (Troubleshooting)

- **Lỗi chứng chỉ SSL khi tải Google Sheet trên Mac:** Script đã tích hợp sẵn `ssl._create_unverified_context()`. Nếu cần tải thủ công, dùng lệnh: `curl -L "<sheet_csv_url>" -o sample-data/temp.csv`.
- **Cột tiền bị định dạng Text có dấu phẩy hoặc ký tự lạ:** Hàm `clean_and_enrich_data` sử dụng `pd.to_numeric(errors='coerce')` tự động làm sạch các giá trị không hợp lệ.
- **Có thêm phòng ban hoặc Manager mới:** Script tự động gom nhóm linh hoạt theo giá trị duy nhất trong cột `Manager` và `Department`, không bị giới hạn cố định 4 người.
