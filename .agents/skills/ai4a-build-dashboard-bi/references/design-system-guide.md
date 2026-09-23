# Dashboard BI Design System & Tokens Guide

Tài liệu hướng dẫn quy chuẩn Design System & Tokens cho Dashboard BI cao cấp.

---

## 1. Color Palette (Bảng màu Neon Tương Phản)

### Background Base (Tông Nền Tối Sâu)
- **Deep Space Mesh:** `radial-gradient(ellipse at 20% -20%, #1e1b4b 0%, #0f172a 50%, #020617 100%)`
- **Dark Obsidian:** `#0B0F19`
- **Glass Panel Surface:** `rgba(15, 23, 42, 0.65)` hoặc `rgba(30, 41, 59, 0.5)`

### Neon Accent Colors (Màu Tương Phản Cực Cao)
| Vai trò chỉ số | Tên màu | HEX | RGB | Gradient Fill |
|---|---|---|---|---|
| **Revenue / Sales** | Cyan Glow | `#06B6D4` | `rgb(6, 182, 212)` | `linear-gradient(180deg, rgba(6, 182, 212, 0.4) 0%, rgba(6, 182, 212, 0.0) 100%)` |
| **Profit / Growth** | Emerald Neon | `#10B981` | `rgb(16, 185, 129)` | `linear-gradient(180deg, rgba(16, 185, 129, 0.4) 0%, rgba(16, 185, 129, 0.0) 100%)` |
| **Operations / Orders** | Electric Blue | `#3B82F6` | `rgb(59, 130, 246)` | `linear-gradient(180deg, rgba(59, 130, 246, 0.4) 0%, rgba(59, 130, 246, 0.0) 100%)` |
| **KPI Target / Special** | Violet Glow | `#8B5CF6` | `rgb(139, 92, 246)` | `linear-gradient(180deg, rgba(139, 92, 246, 0.4) 0%, rgba(139, 92, 246, 0.0) 100%)` |
| **Costs / Penalties** | Coral Rose | `#F43F5E` | `rgb(244, 63, 94)` | `linear-gradient(180deg, rgba(244, 63, 94, 0.4) 0%, rgba(244, 63, 94, 0.0) 100%)` |
| **Alerts / Pending** | Amber Warning | `#F59E0B` | `rgb(245, 158, 11)` | `linear-gradient(180deg, rgba(245, 158, 11, 0.4) 0%, rgba(245, 158, 11, 0.0) 100%)` |

---

## 2. Glassmorphic Surface Specs

```css
/* Multi-layer Glass Card Spec */
.glass-panel {
  background: rgba(15, 23, 42, 0.65);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-top: 1px solid rgba(255, 255, 255, 0.18); /* Highlight viền trên */
  border-left: 1px solid rgba(255, 255, 255, 0.14);
  border-radius: 16px;
  box-shadow: 
    0 8px 32px 0 rgba(0, 0, 0, 0.36),
    inset 0 1px 0 0 rgba(255, 255, 255, 0.1);
}
```

---

## 3. Chart.js Dark Glass Config Theme

```javascript
// Cấu hình mặc định toàn cục cho Chart.js trên Dark Glass UI
Chart.defaults.color = '#94A3B8';
Chart.defaults.font.family = "'Plus Jakarta Sans', 'Inter', sans-serif";

const glassChartOptions = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: {
      labels: {
        color: '#F8FAFC',
        font: { size: 12, weight: '600' },
        usePointStyle: true,
        padding: 20
      }
    },
    tooltip: {
      backgroundColor: 'rgba(15, 23, 42, 0.85)',
      borderColor: 'rgba(255, 255, 255, 0.15)',
      borderWidth: 1,
      titleColor: '#F8FAFC',
      bodyColor: '#CBD5E1',
      padding: 12,
      cornerRadius: 10,
      backdropFilter: 'blur(8px)'
    }
  },
  scales: {
    x: {
      grid: { color: 'rgba(255, 255, 255, 0.04)', drawBorder: false },
      ticks: { color: '#64748B' }
    },
    y: {
      grid: { color: 'rgba(255, 255, 255, 0.04)', drawBorder: false },
      ticks: { color: '#64748B' }
    }
  }
};
```
