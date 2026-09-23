# ERP Operations Data Schema & Calculation Rules

Tài liệu định nghĩa cấu trúc dữ liệu thô (Input Schema), các trường phái sinh (Enriched Fields) và quy tắc định dạng số học áp dụng cho toàn bộ luồng xử lý vận hành.

---

## 1. Dữ Liệu Nguồn (Raw ERP Input)

File dữ liệu xuất từ ERP phải đảm bảo chứa tối thiểu 8 trường thông tin sau:

| Tên Cột | Kiểu Dữ Liệu | Ví Dụ | Mô Tả Nghiệp Vụ | Ràng Buộc Kiểm Tra |
|:---|:---:|:---:|:---|:---|
| `Employee_ID` | String | `E001` | Mã định danh nhân sự | Bắt buộc, không trùng lặp |
| `Employee_Name` | String | `Employee_1` | Họ và tên nhân sự | Tự động trim khoảng trắng thừa |
| `Manager` | String | `Manager_A` | Quản lý trực tiếp phụ trách | Dùng để phân tách file riêng |
| `Department` | String | `Operations` | Phòng ban (HR, Finance, Ops, Sales) | Dùng để phân bổ chi phí |
| `Base_Salary` | Numeric | `29672904` | Lương cơ bản hàng tháng (VNĐ) | Phải > 0, định dạng số nguyên |
| `Bonus` | Numeric | `4117307` | Tiền thưởng hiệu suất/dự án (VNĐ) | >= 0 |
| `Penalty` | Numeric | `939890` | Tiền phạt vi phạm quy chế (VNĐ) | >= 0 |
| `Month` | String | `2026-03` | Kỳ tính lương (định dạng YYYY-MM) | Dùng để đặt tên thư mục & file |

---

## 2. Trường Phái Sinh (Enriched Fields)

Hệ thống tự động tính toán và bổ sung các cột sau vào bảng dữ liệu sạch:

### 2.1. Lương Thực Nhận (`Net_Salary`)
- **Công thức kế toán:**
  $$\text{Net\_Salary} = \text{Base\_Salary} + \text{Bonus} - \text{Penalty}$$
- **Trong Excel:** Sử dụng công thức tham chiếu ô động `=E{row}+F{row}-G{row}` để người xem có thể kiểm tra và tính toán lại khi có điều chỉnh.
- **Ràng buộc:** `Net_Salary` phải luôn dương (> 0). Trường hợp <= 0 sẽ bị gắn cờ cảnh báo nghiêm trọng.

### 2.2. Nhãn Kiểm Toán (`Anomaly_Flag` & `Audit_Note`)
- `Normal`: Dữ liệu hợp lệ trong ngưỡng an toàn.
- `High_Penalty`: Phạt vi phạm vượt trần (>= 1.5M VNĐ hoặc >= 10% Lương cơ bản).
- `High_Bonus`: Thưởng hiệu suất vượt trội (>= 4.0M VNĐ).
- `Dual_Anomaly`: Vừa thưởng cao vừa phạt cao (cần rà soát đặc biệt).
- `Negative_Net`: Net Salary âm hoặc bằng 0.

---

## 3. Quy Chuẩn Trình Bày & Định Dạng Số (Formatting Standards)

- **Đơn vị tiền tệ:** Format chuẩn kế toán VNĐ: `#,##0 "₫"` (ví dụ: `29,672,904 ₫`).
- **Tỷ lệ phần trăm:** Format `0.0%` (ví dụ: `13.2%`).
- **Số lượng / Số đếm:** Format `#,##0` (ví dụ: `200`).
- **Font chữ:** `Segoe UI` hoặc `Calibri`, kích cỡ tiêu đề `14pt bold`, bảng dữ liệu `9pt - 10pt`.
- **Màu sắc chủ đạo:**
  - Header: Navy Executive `#1B365D` (chữ trắng `#FFFFFF`).
  - Sub-header: Medium Slate Blue `#2B6CB0`.
  - Thẻ KPI: Nền `#F8FAFC`, Viền `#CBD5E1`.
  - Cảnh báo phạt cao: Nền đỏ nhạt `#FEE2E2`, chữ đỏ đậm `#991B1B`.
  - Cảnh báo thưởng cao: Nền xanh nhạt `#DCFCE7`, chữ xanh đậm `#166534`.
