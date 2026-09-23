# Operations Data Quality & Anomaly Audit Rules

Tài liệu hướng dẫn quy tắc kiểm toán dữ liệu, nhận diện bất thường và hướng dẫn hành động xử lý trước khi phát hành file cho Manager và giải ngân lương.

---

## 1. Danh Mục Các Loại Bất Thường (Anomaly Matrix)

| Loại Bất Thường | Điều Kiện Kích Hoạt | Mức Độ Rủi Ro | Hành Động Yêu Cầu |
|:---|:---|:---:|:---|
| **Dual Anomaly (Bất thường kép)** | `Penalty >= 1,500,000` **VÀ** `Bonus >= 4,000,000` | 🔴 **High (Cao)** | Yêu cầu Manager trực tiếp xác minh bằng văn bản giải trình. Kiểm tra khả năng nhập sai mã nhân viên hoặc xung đột đánh giá KPI. |
| **High Penalty (Phạt vượt trần)** | `Penalty >= 1,500,000` HOẶC `(Penalty / Base) >= 10%` | 🟡 **Medium (Vừa)** | Quản lý bộ phận đối chiếu biên bản vi phạm quy chế hoặc sự cố vận hành; đào tạo lại nhân sự nếu lỗi lặp lại. |
| **High Bonus (Thưởng vượt khung)** | `Bonus >= 4,000,000` | 🟢 **Low (Thấp)** | Xác nhận phê duyệt từ Giám đốc khối (COO/Director) nhằm đảm bảo ngân sách và tính công bằng nội bộ. |
| **Critical Error (Lương âm/0)** | `Net_Salary <= 0` | 🔴 **High (Cao)** | Tạm dừng phê duyệt dòng lương; điều chỉnh lại mức phạt theo quy định Luật Lao động (không khấu trừ quá trần cho phép). |

---

## 2. Tiêu Chí Đánh Giá Cấp Độ Doanh Nghiệp (Macro Indicators)

1. **Tỷ Lệ Quỹ Thưởng / Quỹ Lương Cơ Bản:**
   - Ngưỡng tối ưu: **10% — 15%**.
   - Dưới 10%: Động lực tăng trưởng và ghi nhận hiệu suất thấp.
   - Trên 20%: Nguy cơ vượt trần ngân sách chi phí nhân sự.

2. **Tỷ Lệ Tiền Phạt / Quỹ Lương Cơ Bản:**
   - Ngưỡng an toàn: **< 3%**.
   - Cảnh báo: **>= 5%** (cho thấy quy trình vận hành có lỗ hổng hoặc quy định phạt chưa hợp lý, gây ảnh hưởng tinh thần nhân sự).

3. **Phân Bổ Tải Quản Lý (Span of Control):**
   - Tiêu chuẩn: Mỗi Manager nên phụ trách từ **15 — 35 nhân sự**.
   - Nếu vượt quá 50 nhân sự/Manager: Tiềm ẩn nguy cơ quá tải phê duyệt, đánh giá hời hợt và chậm trễ xử lý sự cố.

---

## 3. Quy Trình Phê Duyệt Trước Khi Phát Hành

```text
[Pipeline chạy tự động]
         ↓
[Xuất Sheet 3: Kiểm Toán & Bất Thường]
         ↓
[Operation Analyst lọc danh sách Severity = High & Medium]
         ↓
[Gửi file đính kèm & Phiếu rà soát cho từng Manager]
         ↓
[Manager xác nhận / điều chỉnh trong vòng 48 giờ]
         ↓
[Chuyển bộ phận Kế toán & HR giải ngân lương]
```
