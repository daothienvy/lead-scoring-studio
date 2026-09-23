import json
import os
import openpyxl

def generate_static_dashboard():
    excel_path = 'sample-data/ngan_sach_phong_ban.xlsx'
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Cannot find {excel_path}")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() for h in rows[0]]
    
    data = []
    for r in rows[1:]:
        row_dict = {}
        for h, v in zip(headers, r):
            if hasattr(v, 'isoformat'):
                row_dict[h] = v.isoformat()[:10]
            elif isinstance(v, (int, float)):
                row_dict[h] = v
            else:
                row_dict[h] = str(v) if v is not None else ""
        data.append(row_dict)

    data_json = json.dumps(data, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Báo Cáo Ngân Sách Beta Solutions 2026 — Snapshot Tĩnh</title>
  <meta name="description" content="Báo cáo Dashboard Quản lý Ngân sách Beta Solutions 2026. File snapshot tĩnh tự đứng, đầy đủ tương tác và dữ liệu." />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    :root {{
      --brand-budget:  #7C3AED;
      --brand-expense: #06B6D4;
      --brand-over:    #F43F5E;
      --budget-glow:   rgba(124, 58, 237, 0.35);
      --expense-glow:  rgba(6, 182, 212, 0.35);
      --over-glow:     rgba(244, 63, 94, 0.35);
      --bg-dark:       #0B0F19;
      --bg-card:       rgba(15, 23, 42, 0.75);
      --border-glass:  rgba(255, 255, 255, 0.08);
      --border-hover:  rgba(255, 255, 255, 0.18);
      --text-main:     #F8FAFC;
      --text-sub:      #94A3B8;
      --text-muted:    #64748B;
      --radius-lg:     18px;
      --radius-md:     12px;
      --radius-sm:     8px;
      --transition:    all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'Plus Jakarta Sans', sans-serif;
      background: radial-gradient(ellipse 120% 80% at 15% 10%, #150D2E 0%, #080C16 50%, #030710 100%);
      color: var(--text-main);
      min-height: 100vh;
      padding: 24px;
      overflow-x: hidden;
    }}
    .wrap {{ max-width: 1560px; margin: 0 auto; position: relative; z-index: 1; }}
    .glass {{
      background: var(--bg-card);
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      border: 1px solid var(--border-glass);
      border-radius: var(--radius-lg);
      box-shadow: 0 10px 36px rgba(0, 0, 0, 0.45);
      transition: var(--transition);
    }}
    .glass:hover {{
      border-color: var(--border-hover);
      box-shadow: 0 16px 48px rgba(0, 0, 0, 0.55);
    }}
    
    /* HEADER */
    .header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 20px 28px;
      margin-bottom: 22px;
      flex-wrap: wrap;
      gap: 16px;
    }}
    .header-left {{ display: flex; align-items: center; gap: 16px; }}
    .brand-logo {{
      width: 50px;
      height: 50px;
      border-radius: 14px;
      background: linear-gradient(135deg, var(--brand-budget), #4C1D95);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 24px;
      box-shadow: 0 0 24px var(--budget-glow);
    }}
    .brand-info h1 {{
      font-size: 1.28rem;
      font-weight: 800;
      letter-spacing: -0.02em;
    }}
    .brand-info p {{
      font-size: 0.78rem;
      color: var(--text-sub);
      margin-top: 3px;
    }}
    .header-right {{
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }}
    .badge-static {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 6px 14px;
      border-radius: 20px;
      background: rgba(124, 58, 237, 0.16);
      border: 1px solid rgba(124, 58, 237, 0.35);
      color: #C4B5FD;
      font-size: 0.75rem;
      font-weight: 700;
    }}
    .badge-date {{
      font-size: 0.75rem;
      color: var(--text-muted);
    }}

    /* FILTER BAR */
    .filter-bar {{
      padding: 16px 24px;
      margin-bottom: 22px;
      display: flex;
      align-items: center;
      gap: 18px;
      flex-wrap: wrap;
    }}
    .filter-label {{
      font-size: 0.72rem;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .btn-group {{ display: flex; gap: 8px; flex-wrap: wrap; }}
    .btn-filter {{
      padding: 7px 16px;
      background: transparent;
      border: 1px solid var(--border-glass);
      border-radius: var(--radius-sm);
      color: var(--text-sub);
      font-family: inherit;
      font-size: 0.82rem;
      font-weight: 600;
      cursor: pointer;
      transition: var(--transition);
    }}
    .btn-filter:hover {{
      border-color: var(--brand-budget);
      color: var(--text-main);
      background: rgba(124, 58, 237, 0.12);
    }}
    .btn-filter.active {{
      background: rgba(124, 58, 237, 0.25);
      border-color: var(--brand-budget);
      color: #DDD6FE;
      box-shadow: 0 0 14px var(--budget-glow);
    }}
    .dept-select {{
      padding: 8px 14px;
      background: #111827;
      border: 1px solid var(--border-glass);
      border-radius: var(--radius-sm);
      color: var(--text-main);
      font-family: inherit;
      font-size: 0.82rem;
      font-weight: 500;
      cursor: pointer;
      outline: none;
      transition: var(--transition);
      min-width: 220px;
    }}
    .dept-select:hover, .dept-select:focus {{
      border-color: var(--brand-expense);
      box-shadow: 0 0 0 2px rgba(6, 182, 212, 0.2);
    }}

    /* KPI CARDS */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 20px;
      margin-bottom: 22px;
    }}
    @media (max-width: 960px) {{
      .kpi-grid {{ grid-template-columns: 1fr; }}
    }}
    .kpi-card {{
      padding: 24px;
      display: flex;
      flex-direction: column;
      gap: 14px;
      position: relative;
      overflow: hidden;
    }}
    .kpi-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 3px;
    }}
    .card-budget::before {{ background: linear-gradient(90deg, var(--brand-budget), transparent); }}
    .card-expense::before {{ background: linear-gradient(90deg, var(--brand-expense), transparent); }}
    .card-over::before {{ background: linear-gradient(90deg, var(--brand-over), transparent); }}

    .kpi-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .kpi-title {{
      font-size: 0.74rem;
      font-weight: 700;
      color: var(--text-sub);
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .kpi-icon {{
      width: 44px;
      height: 44px;
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
    }}
    .card-budget .kpi-icon {{ background: rgba(124, 58, 237, 0.16); box-shadow: 0 0 16px var(--budget-glow); }}
    .card-expense .kpi-icon {{ background: rgba(6, 182, 212, 0.16); box-shadow: 0 0 16px var(--expense-glow); }}
    .card-over .kpi-icon {{ background: rgba(244, 63, 94, 0.16); box-shadow: 0 0 16px var(--over-glow); }}

    .kpi-number {{
      font-size: 2.25rem;
      font-weight: 800;
      letter-spacing: -0.03em;
      line-height: 1;
    }}
    .card-budget .kpi-number {{ color: #DDD6FE; }}
    .card-expense .kpi-number {{ color: #A5F3FC; }}
    .card-over .kpi-number {{ color: #FECDD3; }}

    .kpi-unit {{
      font-size: 0.74rem;
      color: var(--text-muted);
      margin-top: 4px;
    }}
    .kpi-footer {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding-top: 14px;
      border-top: 1px solid var(--border-glass);
      font-size: 0.74rem;
    }}
    .pill-budget {{ background: rgba(124,58,237,0.18); color: #C4B5FD; padding: 3px 10px; border-radius: 12px; font-weight: 700; }}
    .pill-expense {{ background: rgba(6,182,212,0.18); color: #67E8F9; padding: 3px 10px; border-radius: 12px; font-weight: 700; }}
    .pill-over {{ background: rgba(244,63,94,0.18); color: #FDA4AF; padding: 3px 10px; border-radius: 12px; font-weight: 700; }}

    /* CHARTS */
    .chart-grid {{
      display: grid;
      grid-template-columns: 1.6fr 1fr;
      gap: 20px;
      margin-bottom: 22px;
    }}
    @media (max-width: 1024px) {{
      .chart-grid {{ grid-template-columns: 1fr; }}
    }}
    .chart-card {{ padding: 24px; }}
    .chart-head {{ margin-bottom: 18px; }}
    .chart-head h3 {{ font-size: 0.95rem; font-weight: 700; color: var(--text-main); }}
    .chart-head p {{ font-size: 0.72rem; color: var(--text-muted); margin-top: 2px; }}
    .chart-wrap {{ height: 300px; position: relative; }}

    /* TABLE */
    .table-card {{ padding: 24px; margin-bottom: 22px; }}
    .table-head {{ margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; }}
    .table-head h3 {{ font-size: 0.95rem; font-weight: 700; color: var(--text-main); }}
    .table-head p {{ font-size: 0.72rem; color: var(--text-muted); margin-top: 2px; }}
    .table-container {{ overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.78rem; }}
    th {{
      padding: 12px 14px;
      text-align: left;
      font-size: 0.68rem;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
      border-bottom: 1px solid rgba(255, 255, 255, 0.08);
      white-space: nowrap;
    }}
    td {{
      padding: 12px 14px;
      color: var(--text-sub);
      border-bottom: 1px solid rgba(255, 255, 255, 0.03);
      white-space: nowrap;
    }}
    tr:hover td {{ background: rgba(255, 255, 255, 0.03); color: var(--text-main); }}
    .badge-status {{
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 3px 9px;
      border-radius: 12px;
      font-size: 0.7rem;
      font-weight: 700;
    }}
    .status-over {{
      background: rgba(244, 63, 94, 0.16);
      color: #FDA4AF;
      border: 1px solid rgba(244, 63, 94, 0.35);
    }}

    /* FOOTER */
    .footer {{
      text-align: center;
      padding: 20px;
      font-size: 0.75rem;
      color: var(--text-muted);
    }}
    .footer b {{ color: var(--brand-budget); }}
  </style>
</head>
<body>
  <div class="wrap">
    <!-- HEADER -->
    <header class="glass header">
      <div class="header-left">
        <div class="brand-logo">💜</div>
        <div class="brand-info">
          <h1>Beta Solutions — Dashboard Quản Lý Ngân Sách 2026</h1>
          <p>Trưởng phòng Tài chính kiêm Fullstack Developer · Báo cáo Snapshot Tĩnh</p>
        </div>
      </div>
      <div class="header-right">
        <div class="badge-static">📁 Báo Cáo Snapshot Tĩnh (Offline Ready)</div>
        <div class="badge-date" id="txtDate">Cập nhật: 2026-09-23</div>
      </div>
    </header>

    <!-- FILTER BAR -->
    <section class="glass filter-bar">
      <span class="filter-label">Bộ Lọc Quý:</span>
      <div class="btn-group">
        <button class="btn-filter active" onclick="setQuarter('ALL', this)">Tất cả</button>
        <button class="btn-filter" onclick="setQuarter('Q1', this)">Quý 1</button>
        <button class="btn-filter" onclick="setQuarter('Q2', this)">Quý 2</button>
        <button class="btn-filter" onclick="setQuarter('Q3', this)">Quý 3</button>
        <button class="btn-filter" onclick="setQuarter('Q4', this)">Quý 4</button>
      </div>

      <div style="width: 1px; height: 28px; background: var(--border-glass); margin: 0 4px;"></div>

      <span class="filter-label">Phòng Ban:</span>
      <select class="dept-select" id="selDept" onchange="setDept(this.value)">
        <option value="ALL">🏢 Tất cả phòng ban</option>
      </select>
    </section>

    <!-- KPI CARDS -->
    <section class="kpi-grid">
      <!-- KPI 1: Budget -->
      <div class="glass kpi-card card-budget">
        <div class="kpi-header">
          <span class="kpi-title">Tổng Ngân Sách</span>
          <div class="kpi-icon">💰</div>
        </div>
        <div>
          <div class="kpi-number" id="kpiBudget">0.00</div>
          <div class="kpi-unit">Tỷ VNĐ (Kế hoạch năm 2026)</div>
        </div>
        <div class="kpi-footer">
          <span style="color: var(--text-muted);" id="kpiBudgetCount">96 giao dịch</span>
          <span class="pill-budget">Màu #7C3AED</span>
        </div>
      </div>

      <!-- KPI 2: Expense -->
      <div class="glass kpi-card card-expense">
        <div class="kpi-header">
          <span class="kpi-title">Tổng Chi Tiêu Thực Tế</span>
          <div class="kpi-icon">📊</div>
        </div>
        <div>
          <div class="kpi-number" id="kpiExpense">0.00</div>
          <div class="kpi-unit" id="kpiExpenseUnit">Tỷ VNĐ · Tỷ lệ sử dụng: 0%</div>
        </div>
        <div class="kpi-footer">
          <span style="color: var(--text-muted);" id="kpiExpenseStatus">Tiến độ chi tiêu</span>
          <span class="pill-expense">Màu #06B6D4</span>
        </div>
      </div>

      <!-- KPI 3: Over Budget -->
      <div class="glass kpi-card card-over">
        <div class="kpi-header">
          <span class="kpi-title">Giao Dịch Vượt Ngân Sách</span>
          <div class="kpi-icon">🚨</div>
        </div>
        <div>
          <div class="kpi-number" id="kpiOver">0</div>
          <div class="kpi-unit" id="kpiOverUnit">0 giao dịch · 0.0% tổng GD</div>
        </div>
        <div class="kpi-footer">
          <span style="color: var(--text-muted);" id="kpiOverSub">Cảnh báo rủi ro chi</span>
          <span class="pill-over">Màu #F43F5E</span>
        </div>
      </div>
    </section>

    <!-- CHARTS -->
    <section class="chart-grid">
      <!-- Bar Chart -->
      <div class="glass chart-card">
        <div class="chart-head">
          <h3>📊 So Sánh Ngân Sách vs Chi Tiêu theo Phòng Ban</h3>
          <p>Cột ghép: Ngân Sách (Tím #7C3AED) vs Chi Tiêu (Xanh #06B6D4) · Đơn vị: Triệu VNĐ</p>
        </div>
        <div class="chart-wrap">
          <canvas id="barChart"></canvas>
        </div>
      </div>

      <!-- Donut Chart -->
      <div class="glass chart-card">
        <div class="chart-head">
          <h3>🍩 Tỷ Trọng Chi Tiêu theo Phòng Ban</h3>
          <p>% Phân bổ chi tiêu thực tế giữa các phòng ban</p>
        </div>
        <div class="chart-wrap">
          <canvas id="donutChart"></canvas>
        </div>
      </div>
    </section>

    <!-- DETAIL TABLE -->
    <section class="glass table-card">
      <div class="table-head">
        <div>
          <h3>🚨 Danh Sách Các Giao Dịch Vượt Ngân Sách</h3>
          <p id="tableSub">Sắp xếp theo thứ tự mức vượt ngân sách giảm dần</p>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Mã GD</th>
              <th>Ngày</th>
              <th>Quý</th>
              <th>Phòng Ban</th>
              <th>Hạng Mục</th>
              <th>Ngân Sách (Tr.đ)</th>
              <th>Chi Tiêu (Tr.đ)</th>
              <th>Vượt (Tr.đ)</th>
              <th>% Sử Dụng</th>
              <th>Trạng Thái</th>
            </tr>
          </thead>
          <tbody id="tblBody">
            <!-- Injected by JavaScript -->
          </tbody>
        </table>
      </div>
    </section>

    <!-- FOOTER -->
    <footer class="footer">
      Bản quyền © 2026 <b>Beta Solutions</b>. Dashboard Quản lý Ngân sách theo chuẩn Brand Guideline (Tím #7C3AED, Xanh #06B6D4, Hồng #F43F5E).
    </footer>
  </div>

  <script>
    const RAW_DATA = {data_json};

    const DEPT_SHORT = {{
      "Công nghệ thông tin (IT)": "IT",
      "Marketing & Truyền thông": "Marketing",
      "Kinh doanh (Sales)": "Kinh doanh",
      "Nhân sự (HR)": "Nhân sự",
      "Tài chính - Kế toán": "Tài chính",
      "Vận hành (Operations)": "Vận hành"
    }};

    let activeQuarter = 'ALL';
    let activeDept = 'ALL';
    let chartBar = null;
    let chartDonut = null;

    // Number counter animation
    function animateCounter(elementId, startVal, endVal, duration = 800, isFloat = true) {{
      const el = document.getElementById(elementId);
      if (!el) return;
      const startTime = performance.now();
      function update(currentTime) {{
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const easeOut = 1 - Math.pow(1 - progress, 3);
        const currentVal = startVal + (endVal - startVal) * easeOut;
        el.textContent = isFloat ? currentVal.toFixed(2) : Math.round(currentVal).toString();
        if (progress < 1) {{
          requestAnimationFrame(update);
        }} else {{
          el.textContent = isFloat ? endVal.toFixed(2) : endVal.toString();
        }}
      }}
      requestAnimationFrame(update);
    }}

    function fmtTr(num) {{
      return (num / 1e6).toLocaleString('vi-VN', {{ maximumFractionDigits: 1 }});
    }}

    function getFilteredData() {{
      return RAW_DATA.filter(r => {{
        const qMatch = activeQuarter === 'ALL' || r.Quy === activeQuarter;
        const dMatch = activeDept === 'ALL' || r.Phong_Ban === activeDept;
        return qMatch && dMatch;
      }});
    }}

    function updateKPIs(data) {{
      const totalBudget = data.reduce((s, r) => s + (r.Ngan_Sach_Du_Kien || 0), 0);
      const totalExpense = data.reduce((s, r) => s + (r.Chi_Tieu_Thuc_Te || 0), 0);
      const overList = data.filter(r => r.Trang_Thai === "Vượt ngân sách");
      const overCount = overList.length;

      const budgetBillion = totalBudget / 1e9;
      const expenseBillion = totalExpense / 1e9;
      const usagePct = totalBudget > 0 ? ((totalExpense / totalBudget) * 100).toFixed(1) : 0;
      const overPct = data.length > 0 ? ((overCount / data.length) * 100).toFixed(1) : 0;

      animateCounter('kpiBudget', 0, budgetBillion, 700, true);
      animateCounter('kpiExpense', 0, expenseBillion, 700, true);
      animateCounter('kpiOver', 0, overCount, 700, false);

      document.getElementById('kpiBudgetCount').textContent = data.length + ' giao dịch';
      document.getElementById('kpiExpenseUnit').textContent = 'Tỷ VNĐ · Tỷ lệ sử dụng: ' + usagePct + '%';
      document.getElementById('kpiOverUnit').textContent = overCount + ' GD vượt · ' + overPct + '% tổng số GD';
    }}

    function updateCharts(data) {{
      const depts = [...new Set(RAW_DATA.map(r => r.Phong_Ban))].sort();
      const budgetMap = {{}}, expenseMap = {{}};
      depts.forEach(d => {{ budgetMap[d] = 0; expenseMap[d] = 0; }});

      data.forEach(r => {{
        budgetMap[r.Phong_Ban] = (budgetMap[r.Phong_Ban] || 0) + (r.Ngan_Sach_Du_Kien || 0);
        expenseMap[r.Phong_Ban] = (expenseMap[r.Phong_Ban] || 0) + (r.Chi_Tieu_Thuc_Te || 0);
      }});

      const labels = depts.map(d => DEPT_SHORT[d] || d);
      const budgetVals = depts.map(d => Math.round(budgetMap[d] / 1e6));
      const expenseVals = depts.map(d => Math.round(expenseMap[d] / 1e6));

      // Bar Chart
      if (chartBar) chartBar.destroy();
      const ctxBar = document.getElementById('barChart').getContext('2d');
      chartBar = new Chart(ctxBar, {{
        type: 'bar',
        data: {{
          labels: labels,
          datasets: [
            {{
              label: 'Ngân Sách',
              data: budgetVals,
              backgroundColor: 'rgba(124, 58, 237, 0.85)',
              borderColor: '#7C3AED',
              borderWidth: 1.5,
              borderRadius: 6
            }},
            {{
              label: 'Chi Tiêu Thực Tế',
              data: expenseVals,
              backgroundColor: 'rgba(6, 182, 212, 0.80)',
              borderColor: '#06B6D4',
              borderWidth: 1.5,
              borderRadius: 6
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ labels: {{ color: '#94A3B8', font: {{ family: 'Plus Jakarta Sans', size: 12 }} }} }},
            tooltip: {{
              backgroundColor: 'rgba(11, 15, 25, 0.95)',
              borderColor: 'rgba(255, 255, 255, 0.1)',
              borderWidth: 1,
              callbacks: {{
                label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString('vi-VN') + ' Tr.đ'
              }}
            }}
          }},
          scales: {{
            x: {{
              ticks: {{ color: '#64748B', font: {{ family: 'Plus Jakarta Sans' }} }},
              grid: {{ color: 'rgba(255, 255, 255, 0.04)' }}
            }},
            y: {{
              ticks: {{
                color: '#64748B',
                font: {{ family: 'Plus Jakarta Sans' }},
                callback: v => v.toLocaleString('vi-VN')
              }},
              grid: {{ color: 'rgba(255, 255, 255, 0.05)' }}
            }}
          }}
        }}
      }});

      // Donut Chart
      if (chartDonut) chartDonut.destroy();
      const ctxDonut = document.getElementById('donutChart').getContext('2d');
      const donutColors = [
        'rgba(124, 58, 237, 0.9)',
        'rgba(6, 182, 212, 0.9)',
        'rgba(244, 63, 94, 0.9)',
        'rgba(16, 185, 129, 0.9)',
        'rgba(245, 158, 11, 0.9)',
        'rgba(99, 102, 241, 0.9)'
      ];

      chartDonut = new Chart(ctxDonut, {{
        type: 'doughnut',
        data: {{
          labels: labels,
          datasets: [{{
            data: expenseVals,
            backgroundColor: donutColors,
            borderColor: 'rgba(11, 15, 25, 0.8)',
            borderWidth: 2,
            hoverOffset: 8
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          cutout: '62%',
          plugins: {{
            legend: {{
              position: 'bottom',
              labels: {{
                color: '#94A3B8',
                font: {{ family: 'Plus Jakarta Sans', size: 11 }},
                padding: 10,
                usePointStyle: true
              }}
            }},
            tooltip: {{
              backgroundColor: 'rgba(11, 15, 25, 0.95)',
              borderColor: 'rgba(255, 255, 255, 0.1)',
              borderWidth: 1,
              callbacks: {{
                label: function(c) {{
                  const total = c.dataset.data.reduce((a, b) => a + b, 0);
                  const pct = total > 0 ? ((c.parsed / total) * 100).toFixed(1) : 0;
                  return ' ' + c.parsed.toLocaleString('vi-VN') + ' Tr.đ (' + pct + '%)';
                }}
              }}
            }}
          }}
        }}
      }});
    }}

    function updateTable(data) {{
      const overList = data.filter(r => r.Trang_Thai === "Vượt ngân sách")
                           .sort((a, b) => a.Chenh_Lech - b.Chenh_Lech);
      const tbody = document.getElementById('tblBody');
      document.getElementById('tableSub').textContent = overList.length + ' giao dịch vượt ngân sách được phát hiện';

      if (overList.length === 0) {{
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--text-muted)">Không có giao dịch vượt ngân sách trong bộ lọc này.</td></tr>';
        return;
      }}

      tbody.innerHTML = overList.map(r => `
        <tr>
          <td style="font-family: monospace; font-size: 0.72rem; color: #94A3B8;">${{r.Ma_Giao_Dich}}</td>
          <td>${{String(r.Ngay || '').slice(0, 10)}}</td>
          <td style="color: #DDD6FE; font-weight: 700;">${{r.Quy}}</td>
          <td style="color: #CBD5E1;">${{DEPT_SHORT[r.Phong_Ban] || r.Phong_Ban}}</td>
          <td style="color: #F8FAFC;">${{r.Hang_Muc}}</td>
          <td style="color: #C4B5FD; font-weight: 600;">${{fmtTr(r.Ngan_Sach_Du_Kien)}}</td>
          <td style="color: #67E8F9; font-weight: 600;">${{fmtTr(r.Chi_Tieu_Thuc_Te)}}</td>
          <td style="color: #FDA4AF; font-weight: 700;">+${{fmtTr(Math.abs(r.Chenh_Lech))}}</td>
          <td style="color: #FDA4AF; font-weight: 700;">${{r.Phan_Tram_Su_Dung}}%</td>
          <td><span class="badge-status status-over">🔴 Vượt NS</span></td>
        </tr>
      `).join('');
    }}

    function renderAll() {{
      const data = getFilteredData();
      updateKPIs(data);
      updateCharts(data);
      updateTable(data);
    }}

    function setQuarter(q, btn) {{
      activeQuarter = q;
      document.querySelectorAll('.btn-filter').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderAll();
    }}

    function setDept(d) {{
      activeDept = d;
      renderAll();
    }}

    // Init
    document.addEventListener('DOMContentLoaded', () => {{
      const depts = [...new Set(RAW_DATA.map(r => r.Phong_Ban))].sort();
      const sel = document.getElementById('selDept');
      depts.forEach(d => {{
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        sel.appendChild(opt);
      }});
      renderAll();
    }});
  </script>
</body>
</html>
"""

    os.makedirs('outputs/reports', exist_ok=True)
    out_file = 'outputs/reports/BetaSolutions_NganSach_Dashboard_Snapshot.html'
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✅ Generated static dashboard successfully: {out_file} ({len(html_content)} bytes)")

if __name__ == '__main__':
    generate_static_dashboard()
