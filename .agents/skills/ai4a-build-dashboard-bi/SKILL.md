---
name: ai4a:build-dashboard-BI
description: "Build modern, state-of-the-art BI & Analytics Dashboards adhering to high-end UI design standards: Glassmorphism, Dark Mode, KPI card layouts, vibrant high-contrast color palettes, and interactive number-counter animations. Use when designing executive dashboards, business intelligence UI, sales/operations performance monitors, or interactive Web BI apps."
user-invocable: true
when_to_use: "Use when creating or updating BI dashboards, data visualization interfaces, executive KPI monitoring pages, or Web Analytics applications."
category: design-system
keywords: [dashboard, bi, analytics, glassmorphism, dark-mode, kpi, counter-up, ui-ux, chartjs, apexcharts, ai4a]
argument-hint: "[topic or data-source] [--html] [--react] [--standalone] [--dark] [--brand-colors]"
metadata:
  author: "MT Đức Thuận"
  brand: "AI4A"
  course: "Agentic AI with Google Antigravity"
  version: "1.0.0"
---

# AI4A: Build Dashboard BI

> **Đóng gói & phát triển bởi MT Đức Thuận**  
> *Dành tặng học viên chương trình Agentic AI with Google Antigravity (AI4A)*

Kỹ năng này đóng gói bộ tiêu chuẩn thiết kế Dashboard Business Intelligence (BI) hiện đại cấp doanh nghiệp, giúp biến các bộ dữ liệu thô (ERP, CRM, Sales, HR, Finance, Operations) thành các bảng điều khiển trực quan (Executive Dashboard) đẳng cấp với 5 trụ cột UI/UX hiện đại: **Glassmorphism**, **Dark Mode**, **KPI Layout**, **Phối màu tương phản Neon**, và **Hiệu ứng nhảy số (Counter-up Animation)**.

---

## 💎 5 Trụ Cột Thiết Kế Dashboard Hiện Đại

### 1. Glassmorphism (Hiệu ứng Kính mờ Hữu cơ)
Tạo độ sâu thị giác và cảm giác hiệu ứng sang trọng bằng kính mờ xuyên thấu:
- **Backdrop Blur:** `backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px);`
- **Translucent Background:** `background: rgba(15, 23, 42, 0.65);` (Dark Slate) hoặc `rgba(255, 255, 255, 0.03);`
- **Subtle Glass Border:** `border: 1px solid rgba(255, 255, 255, 0.1);` (viền mỏng tương phản ánh kim)
- **Soft Glow Box Shadow:** `box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37), inset 0 0 0 1px rgba(255, 255, 255, 0.05);`

### 2. Sleek Dark Mode (Nền Tối Sang Trọng & Tương Phản Cao)
- **Canvas Base:** Nền tối dải Gradient mượt mờ sang trọng: `background: radial-gradient(circle at top left, #0f172a, #0b0f19, #05070f);`
- **Typography:** Sử dụng phông chữ hiện đại từ Google Fonts (`Outfit`, `Inter`, hoặc `Plus Jakarta Sans`).
- **Phân cấp Văn bản (Text Hierarchy):**
  - Title/Value chính: `#F8FAFC` (Pure white-gray)
  - Subtitle/Secondary text: `#94A3B8` (Slate gray)
  - Muted labels/Icons: `#64748B` hoặc các gam màu Accent phát sáng.

### 3. KPI Layout & Grid Bố Cục Chuyên Nghiệp
- **Top Bar (Header & Filter Bar):** Tiêu đề Dashboard, Badge Thời gian thực, Bộ lọc Ngày/Khu vực/Phòng ban và Nút Export.
- **Row 1 (KPI Cards Grid):** 4 thẻ KPI summary chuẩn (Doanh thu, Lợi nhuận, Đơn hàng/Khách hàng, Tỷ lệ tăng trưởng). Mỗi thẻ có:
  - Icon phát sáng mờ (`background: rgba(..., 0.15)`)
  - Nhãn KPI & Chỉ số chính với **Hiệu ứng Nhảy số**
  - Trend Badge (+12.5% MoM màu Emerald / -3.2% màu Rose)
- **Row 2 (Biểu đồ Xu hướng & Phân bổ Main Charts):**
  - Khối chính (65%): Area / Line Chart xu hướng thời gian với dải Gradient mờ dưới đường biểu đồ.
  - Khối phụ (35%): Donut / Bar Chart phân bổ theo danh mục/khu vực.
- **Row 3 (Bảng Dữ liệu Detail & Bảng Xếp hạng):** Top Sản phẩm / Top Nhân viên / Chi tiết giao dịch với thanh cuộn mượt và hover highlight.

### 4. Phối Màu Tương Phản High-Contrast Neon Accents
Dùng các tông màu Neon rực rỡ tương phản cao trên nền tối để thu hút thị giác:
- **Tăng trưởng / Lợi nhuận / Doanh thu:** Emerald Neon (`#10B981`) / Cyan Glow (`#06B6D4`)
- **Chi phí / Cảnh báo / Giảm:** Coral Rose (`#F43F5E`) / Amber Orange (`#F59E0B`)
- **Khối lượng / Đơn hàng / Vận hành:** Electric Blue (`#3B82F6`) / Indigo (`#6366F1`)
- **Chỉ số Phụ / Mục tiêu:** Violet Purple (`#8B5CF6`) / Neon Magenta (`#D946EF`)

### 5. Hiệu Ứng Nhảy Số (Counter-up / Number Ticker Animation)
Chạy hiệu ứng tăng số mượt mà từ `0` tới giá trị đích khi trang tải xong hoặc khi thay đổi bộ lọc:
- Sử dụng thuật toán `requestAnimationFrame` đảm bảo mượt 60fps.
- Tự động định dạng tiền tệ VNĐ (ví dụ `1.250.000.000 ₫` hoặc `1.25 Tỷ ₫`), USD (`$1,250,000`), Phần trăm (`+18.5%`), và Số nguyên (`4,850`).

---

## 🛠️ Quy Trình Triển Khai 5 Bước (Runbook)

### Bước 1: Phân tích Dữ liệu Đầu vào & Chỉ số KPI
- Tiếp nhận file dữ liệu (CSV, Excel, JSON hoặc ERP Export).
- Xác định 4 chỉ số KPI quan trọng nhất (Top-line Metrics) và các trục phân tích (Dimensions: Thời gian, Địa lý, Danh mục).

### Bước 2: Khởi tạo Cấu trúc Layout HTML5 / CSS Glassmorphism
- Áp dụng bộ thẻ CSS utility Glassmorphism chuẩn vào `index.css` hoặc thẻ `<style>`.
- Đảm bảo tính Responsive (`CSS Grid`, `Flexbox`, `clamp()`).

### Bước 3: Tích hợp Thư viện Biểu đồ (Chart.js / ApexCharts)
- Khởi tạo các chart instance với Theme Dark Glass:
  - Tắt đường lưới ngang dọc thô cứng (`grid: { color: 'rgba(255, 255, 255, 0.05)' }`).
  - Sử dụng Gradient Fill mượt cho Line/Area Chart.
  - Thiết lập Tooltip Glassmorphism tùy chỉnh.

### Bước 4: Nhúng JavaScript Counter-up & Tương tác
- Gắn script đếm số tự động vào các phần tử `.counter-value`.
- Thêm hiệu ứng Hover Glass tilt / Glow hiệu ứng khi di chuột qua thẻ KPI.

### Bước 5: Kiểm Thử & Đóng Gói Output
- Kiểm tra hiển thị trên màn hình Desktop (1920x1080) và Laptop (1366x768).
- Xuất file Dashboard HTML hoàn chỉnh vào thư mục `outputs/dashboards/<topic>_dashboard.html`.

---

## 💻 Code Boilerplate Mẫu Chuẩn (Ready-to-Use Snippet)

### 1. Hàm JavaScript Hiệu Ứng Nhảy Số (Number Ticker Animation)

```javascript
/**
 * Animate numbers smoothly from 0 to target value
 * @param {HTMLElement} el - Element to display value
 * @param {number} target - Destination number
 * @param {string} prefix - Symbol prefix (e.g. '$', '+')
 * @param {string} suffix - Symbol suffix (e.g. ' ₫', '%')
 * @param {number} duration - Animation time in ms (default 1500ms)
 */
function animateCounter(el, target, prefix = '', suffix = '', duration = 1500) {
  const startTime = performance.now();
  const startVal = 0;
  
  function update(currentTime) {
    const elapsed = currentTime - startTime;
    const progress = Math.min(elapsed / duration, 1);
    
    // Easing function: easeOutCubic
    const easeProgress = 1 - Math.pow(1 - progress, 3);
    const currentVal = Math.floor(startVal + (target - startVal) * easeProgress);
    
    // Format number with locale separators
    const formatted = currentVal.toLocaleString('vi-VN');
    el.textContent = `${prefix}${formatted}${suffix}`;
    
    if (progress < 1) {
      requestAnimationFrame(update);
    } else {
      el.textContent = `${prefix}${target.toLocaleString('vi-VN')}${suffix}`;
    }
  }
  
  requestAnimationFrame(update);
}

// Tự động kích hoạt cho tất cả phần tử có class .counter-value khi trang load
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.counter-value').forEach(card => {
    const target = parseFloat(card.getAttribute('data-target') || '0');
    const prefix = card.getAttribute('data-prefix') || '';
    const suffix = card.getAttribute('data-suffix') || '';
    animateCounter(card, target, prefix, suffix, 1600);
  });
});
```

### 2. Bộ CSS Glassmorphism & Dark Mode Style Sheet Mẫu

```css
:root {
  --bg-dark: #0b0f19;
  --bg-card: rgba(15, 23, 42, 0.65);
  --border-glass: rgba(255, 255, 255, 0.1);
  --border-glass-hover: rgba(255, 255, 255, 0.25);
  --text-main: #f8fafc;
  --text-sub: #94a3b8;
  
  --accent-emerald: #10b981;
  --accent-cyan: #06b6d4;
  --accent-blue: #3b82f6;
  --accent-purple: #8b5cf6;
  --accent-rose: #f43f5e;
  --accent-amber: #f59e0b;
}

body {
  margin: 0;
  padding: 24px;
  background: radial-gradient(circle at top left, #0f172a, #0b0f19, #05070f);
  color: var(--text-main);
  font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
  min-height: 100vh;
}

/* Thẻ Glassmorphism Chuẩn */
.glass-card {
  background: var(--bg-card);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border: 1px solid var(--border-glass);
  border-radius: 16px;
  padding: 20px;
  box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37), inset 0 0 0 1px rgba(255, 255, 255, 0.05);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

.glass-card:hover {
  border-color: var(--border-glass-hover);
  transform: translateY(-3px);
  box-shadow: 0 12px 40px 0 rgba(0, 0, 0, 0.5), 0 0 20px rgba(6, 182, 212, 0.15);
}

/* Lưới KPI Grid */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
  gap: 20px;
  margin-bottom: 24px;
}

.kpi-card {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}

.kpi-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
}

.kpi-title {
  font-size: 0.875rem;
  color: var(--text-sub);
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

.kpi-icon {
  width: 40px;
  height: 40px;
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.2rem;
}

.kpi-value {
  font-size: 1.85rem;
  font-weight: 700;
  color: var(--text-main);
  letter-spacing: -0.02em;
}

.trend-badge {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 8px;
  border-radius: 20px;
  font-size: 0.75rem;
  font-weight: 600;
}

.trend-up {
  background: rgba(16, 185, 129, 0.15);
  color: var(--accent-emerald);
  border: 1px solid rgba(16, 185, 129, 0.3);
}

.trend-down {
  background: rgba(244, 63, 94, 0.15);
  color: var(--accent-rose);
  border: 1px solid rgba(244, 63, 94, 0.3);
}
```

---

## 📚 Tài Liệu Đi Kèm (Bundled References)

| Loại tài liệu | Đường dẫn liên kết | Mục đích |
|---|---|---|
| **Design Tokens & Theme** | [references/design-system-guide.md](references/design-system-guide.md) | Chi tiết bảng màu HSL/HEX, Font Typography scale và Chart.js Dark Glass config |

---

## 🛡️ Quy Tắc & Giới Hạn (Guardrails)

1. **Tuyệt đối không dùng màu đỏ/xanh thô trùng lặp:** Luôn dùng bảng màu tương phản Neon đã được tuyển chọn (`#10B981`, `#06B6D4`, `#F43F5E`, `#8B5CF6`).
2. **Kiểm tra độ tương phản Chữ:** Text trên nền kính tối luôn đạt chuẩn WCAG AA (độ tương phản > 4.5:1).
3. **Đơn tệp tự chứa (Self-contained HTML):** Khi tạo báo cáo dashboard HTML, ưu tiên gói tất cả CSS/JS trong 1 file HTML duy nhất để người dùng dễ dàng mở trực tiếp trên trình duyệt mà không bị lỗi cross-origin hay thiếu tài nguyên.
