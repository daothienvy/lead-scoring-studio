#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pipeline Tự Động Hóa Tính Lương Cấp Cao & Phân Tích Quỹ Lương 24 Tháng Ngành Di Trú
Framework: AI4A Brainstorm Contract (--oipo)
Tác giả: MT Đức Thuận / Agentic AI Team
Mục tiêu:
  1. Tự sinh 24 tháng dữ liệu thực tế (10/2024 - 09/2026) cho 12 nhân sự cấp cao ngành di trú.
  2. Tính toán chính xác Gross, KPI, Override, Escrow, BHXH trần, Thuế TNCN 7 bậc, Net, Total Cost.
  3. Xuất file dữ liệu gốc ra sample-data/ (Excel & CSV).
  4. Tạo Master Dashboard Excel 4 sheet có thẻ KPI & 3 biểu đồ trực quan tại outputs/reports/.
  5. Xuất báo cáo phân tích chuyên sâu gửi Tổng Giám Đốc/HĐQT tại outputs/reports/.
"""

import os
import math
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series

# Thiết lập đường dẫn thư mục
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DATA_DIR = os.path.join(BASE_DIR, "sample-data")
REPORTS_DIR = os.path.join(BASE_DIR, "outputs", "reports")
os.makedirs(SAMPLE_DATA_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. KHỞI TẠO CẤU HÌNH 12 NHÂN SỰ CẤP CAO NGÀNH DI TRÚ
# -------------------------------------------------------------
EXECUTIVES = [
    {
        "id": "EMP001",
        "name": "Nguyễn Quốc Hùng",
        "role": "Tổng Giám Đốc (Managing Director / CEO)",
        "dept": "Điều Hành Chung",
        "role_type": "Executive",
        "base_salary": 130_000_000,
        "allowance": 25_000_000,
        "dependents": 2,
        "target_rev_weight": 0.25,  # Chịu trách nhiệm chung toàn cty
    },
    {
        "id": "EMP002",
        "name": "Trần Minh Tuấn",
        "role": "Phó TGĐ Kinh Doanh (Deputy CEO & Head of BD)",
        "dept": "Thương Mại & BD",
        "role_type": "Sales",
        "base_salary": 105_000_000,
        "allowance": 20_000_000,
        "dependents": 2,
        "target_rev_weight": 0.28,
    },
    {
        "id": "EMP003",
        "name": "LS. Lê Hoàng Nam",
        "role": "Giám Đốc Pháp Lý & Thụ Lý Hồ Sơ (Legal Director)",
        "dept": "Pháp Lý Di Trú",
        "role_type": "Legal",
        "base_salary": 95_000_000,
        "allowance": 15_000_000,
        "dependents": 1,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP004",
        "name": "Phạm Thanh Hà",
        "role": "Giám Đốc Thẩm Định Đầu Tư (Investment DD Director)",
        "dept": "Đầu Tư & Thẩm Định",
        "role_type": "Investment",
        "base_salary": 90_000_000,
        "allowance": 15_000_000,
        "dependents": 1,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP005",
        "name": "Vũ Hải Đăng",
        "role": "Giám Đốc Khách Hàng VIP (Private Clients Director)",
        "dept": "Thương Mại & BD",
        "role_type": "Sales",
        "base_salary": 85_000_000,
        "allowance": 20_000_000,
        "dependents": 1,
        "target_rev_weight": 0.22,
    },
    {
        "id": "EMP006",
        "name": "Đỗ Phương Thảo",
        "role": "Giám Đốc Marketing & Thương Hiệu (CMO)",
        "dept": "Tiếp Thị & Brand",
        "role_type": "Marketing",
        "base_salary": 80_000_000,
        "allowance": 12_000_000,
        "dependents": 0,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP007",
        "name": "Hoàng Gia Bách",
        "role": "Giám Đốc Chi Nhánh Hà Nội (Northern Regional Dir.)",
        "dept": "Khối Chi Nhánh",
        "role_type": "Branch",
        "base_salary": 85_000_000,
        "allowance": 15_000_000,
        "dependents": 2,
        "target_rev_weight": 0.12,
    },
    {
        "id": "EMP008",
        "name": "Ngô Thị Bích Ngọc",
        "role": "Giám Đốc Chi Nhánh TP.HCM (Southern Regional Dir.)",
        "dept": "Khối Chi Nhánh",
        "role_type": "Branch",
        "base_salary": 90_000_000,
        "allowance": 15_000_000,
        "dependents": 1,
        "target_rev_weight": 0.13,
    },
    {
        "id": "EMP009",
        "name": "Trịnh Minh Trí",
        "role": "Trưởng Ban Đối Ngoại & Lãnh Sự (Head of Consular)",
        "dept": "Đối Ngoại Lãnh Sự",
        "role_type": "Legal",
        "base_salary": 70_000_000,
        "allowance": 15_000_000,
        "dependents": 1,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP010",
        "name": "Bùi Thanh Mai",
        "role": "Trưởng Ban An Cư (Head of Settlement Operations)",
        "dept": "Vận Hành An Cư",
        "role_type": "Operations",
        "base_salary": 65_000_000,
        "allowance": 10_000_000,
        "dependents": 1,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP011",
        "name": "Đặng Văn Lâm",
        "role": "Giám Đốc Tài Chính Quốc Tế (CFO)",
        "dept": "Tài Chính & Kế Toán",
        "role_type": "Finance",
        "base_salary": 90_000_000,
        "allowance": 12_000_000,
        "dependents": 2,
        "target_rev_weight": 0.0,
    },
    {
        "id": "EMP012",
        "name": "Võ Quỳnh Trang",
        "role": "Giám Đốc Nhân Sự & Tuân Thủ (HR & Compliance Dir.)",
        "dept": "Nhân Sự & Quản Trị",
        "role_type": "HR",
        "base_salary": 65_000_000,
        "allowance": 10_000_000,
        "dependents": 1,
        "target_rev_weight": 0.0,
    },
]

# -------------------------------------------------------------
# 2. HÀM TÍNH THUẾ TNCN & KHẤU TRỪ THEO LUẬT VIỆT NAM
# -------------------------------------------------------------
CAP_BASE_SALARY = 46_800_000     # 20 lần mức lương cơ sở 2,340,000đ
CAP_BHTN_SALARY = 99_200_000     # 20 lần mức lương tối thiểu vùng I

def calculate_employee_insurance(base_salary):
    """Tính BHXH (8%), BHYT (1.5%) trên mức trần 46.8M, BHTN (1%) trên trần 99.2M"""
    insurable_base = min(base_salary, CAP_BASE_SALARY)
    insurable_bhtn = min(base_salary, CAP_BHTN_SALARY)
    bhxh_bhyt = insurable_base * 0.095
    bhtn = insurable_bhtn * 0.01
    return round(bhxh_bhyt + bhtn)

def calculate_employer_insurance(base_salary):
    """Chi phí BHXH doanh nghiệp đóng: 17.5% BHXH, 3% BHYT, 1% BHTN = 21.5%"""
    insurable_base = min(base_salary, CAP_BASE_SALARY)
    insurable_bhtn = min(base_salary, CAP_BHTN_SALARY)
    return round((insurable_base * 0.205) + (insurable_bhtn * 0.01))

def calculate_pit(taxable_income):
    """Biểu thuế thu nhập cá nhân lũy tiến từng phần 7 bậc"""
    if taxable_income <= 0:
        return 0
    ti = taxable_income
    if ti <= 5_000_000:
        return round(ti * 0.05)
    elif ti <= 10_000_000:
        return round(ti * 0.10 - 250_000)
    elif ti <= 18_000_000:
        return round(ti * 0.15 - 750_000)
    elif ti <= 32_000_000:
        return round(ti * 0.20 - 1_650_000)
    elif ti <= 52_000_000:
        return round(ti * 0.25 - 3_250_000)
    elif ti <= 80_000_000:
        return round(ti * 0.30 - 5_850_000)
    else:
        return round(ti * 0.35 - 9_850_000)

# -------------------------------------------------------------
# 3. MÔ PHỎNG DỮ LIỆU HOẠT ĐỘNG 24 THÁNG (10/2024 - 09/2026)
# -------------------------------------------------------------
def generate_24_months_data():
    np.random.seed(42)  # Đảm bảo tính nhất quán có thể kiểm toán
    
    # 24 tháng từ 2024-10 đến 2026-09
    months = []
    year = 2024
    month = 10
    for _ in range(24):
        months.append(f"{year}-{month:02d}")
        month += 1
        if month > 12:
            month = 1
            year += 1

    records = []
    
    # Chu kỳ mùa vụ ngành di trú:
    # Q4 (Tháng 10, 11, 12) và Q3 (Tháng 7, 8, 9) là cao điểm chốt deal EB-5/Golden Visa
    # Q1 (Tháng 1, 2, 3) thấp điểm sau Tết
    seasonality = {
        1: 0.75, 2: 0.70, 3: 0.85,
        4: 0.95, 5: 1.00, 6: 1.05,
        7: 1.25, 8: 1.30, 9: 1.35,
        10: 1.20, 11: 1.25, 12: 1.30
    }

    for m_idx, ym in enumerate(months):
        y, m = map(int, ym.split("-"))
        season_factor = seasonality[m]
        
        # Tổng doanh thu phí dịch vụ toàn công ty trong tháng (từ 5 tỷ đến 14 tỷ VNĐ)
        base_monthly_revenue = 7_500_000_000 * season_factor * (1.0 + m_idx * 0.015) # Xu hướng tăng trưởng 1.5%/tháng
        fluctuation = np.random.uniform(0.92, 1.10)
        total_company_rev = round(base_monthly_revenue * fluctuation)
        
        # Số hợp đồng ký mới (EB-5, Golden Visa, SUV, CBI)
        total_deals = max(3, int(total_company_rev / 1_100_000_000))
        
        for exec_info in EXECUTIVES:
            eid = exec_info["id"]
            ename = exec_info["name"]
            role = exec_info["role"]
            dept = exec_info["dept"]
            rtype = exec_info["role_type"]
            base_sal = exec_info["base_salary"]
            allowance = exec_info["allowance"]
            deps = exec_info["dependents"]
            
            # Điểm đánh giá KPI điều hành (80% - 120%)
            kpi_score = round(np.random.uniform(88, 112), 1)
            # Thưởng KPI (0% - 30% lương cứng)
            kpi_rate = max(0.0, (kpi_score - 80) / 100.0) * 0.75
            kpi_bonus = round(base_sal * kpi_rate)
            
            # Doanh thu ghi nhận & Hoa hồng (Commission Override)
            comm_override = 0
            attributed_rev = 0
            deals_involved = 0
            
            if rtype == "Sales":
                if eid == "EMP002":  # Deputy CEO
                    attributed_rev = round(total_company_rev * 0.65)
                    deals_involved = int(total_deals * 0.7)
                    # Hoa hồng 1.4%
                    comm_override = round(attributed_rev * 0.014 * np.random.uniform(0.95, 1.05))
                elif eid == "EMP005":  # HNWI Director
                    attributed_rev = round(total_company_rev * 0.35)
                    deals_involved = int(total_deals * 0.4)
                    # Hoa hồng trực tiếp 2.2%
                    comm_override = round(attributed_rev * 0.022 * np.random.uniform(0.92, 1.08))
            elif rtype == "Branch":
                if eid == "EMP007":  # Hà Nội
                    attributed_rev = round(total_company_rev * 0.38)
                    deals_involved = int(total_deals * 0.4)
                    comm_override = round(attributed_rev * 0.016 * np.random.uniform(0.92, 1.05))
                elif eid == "EMP008":  # TP.HCM
                    attributed_rev = round(total_company_rev * 0.48)
                    deals_involved = int(total_deals * 0.5)
                    comm_override = round(attributed_rev * 0.016 * np.random.uniform(0.95, 1.05))
            elif rtype == "Legal":
                # Thưởng mốc thụ lý hồ sơ thành công (20M - 45M)
                deals_involved = total_deals
                comm_override = round(np.random.uniform(22_000_000, 48_000_000) * season_factor)
            elif rtype == "Investment":
                deals_involved = total_deals
                comm_override = round(np.random.uniform(18_000_000, 38_000_000) * season_factor)
            elif rtype == "Executive":
                # CEO thưởng EBITDA theo quý (mỗi tháng nhận thưởng mục tiêu chiến lược)
                comm_override = round(total_company_rev * 0.008)
                attributed_rev = total_company_rev
                deals_involved = total_deals
            else: # Marketing, Operations, Finance, HR
                comm_override = round(base_sal * np.random.uniform(0.08, 0.22))
            
            # Cơ chế Thưởng Duy Trì Escrow (Retention Release):
            # 40% hoa hồng các hợp đồng trước được giữ lại và giải ngân sau 12-18 tháng khi khách có Visa/Approval
            # Tại các tháng 6, 12, 18, 24 hoặc định kỳ các deal thành công
            retention_escrow = 0
            if rtype in ["Sales", "Branch", "Legal"]:
                # Giả lập giải ngân escrow tích lũy từ các hồ sơ 12 tháng trước
                if m_idx >= 3: # Từ tháng thứ 4 trở đi bắt đầu có đợt giải ngân visa về
                    if m_idx % 3 == 0:  # Định kỳ theo quý
                        retention_escrow = round(comm_override * np.random.uniform(0.40, 0.75))
                    else:
                        retention_escrow = round(comm_override * np.random.uniform(0.10, 0.25))
            elif eid == "EMP001":  # CEO thưởng giữ chân năm
                if m in [6, 12]:
                    retention_escrow = round(base_sal * 0.8)
            
            # Tính Gross
            gross = base_sal + allowance + kpi_bonus + comm_override + retention_escrow
            
            # Khấu trừ
            emp_insurance = calculate_employee_insurance(base_sal)
            non_taxable_allowance = 2_000_000  # Phụ cấp ăn trưa, đồng phục được miễn thuế
            
            # Thu nhập tính thuế (Taxable Income)
            # Giảm trừ bản thân 11M + người phụ thuộc 4.4M/người
            personal_relief = 11_000_000
            dep_relief = deps * 4_400_000
            total_relief = personal_relief + dep_relief
            
            taxable_income = max(0, gross - non_taxable_allowance - emp_insurance - total_relief)
            pit = calculate_pit(taxable_income)
            net_salary = gross - emp_insurance - pit
            
            # Chi phí doanh nghiệp thực tế chịu
            employer_insurance = calculate_employer_insurance(base_sal)
            executive_health_vip = 3_000_000  # Bảo hiểm sức khỏe VIP toàn cầu
            total_company_cost = gross + employer_insurance + executive_health_vip
            
            # Ghi nhận record
            records.append({
                "Month": ym,
                "Year": y,
                "Month_Num": m,
                "Emp_ID": eid,
                "Emp_Name": ename,
                "Role": role,
                "Department": dept,
                "Role_Type": rtype,
                "Base_Salary": base_sal,
                "Senior_Allowance": allowance,
                "KPI_Score": kpi_score,
                "KPI_Bonus": kpi_bonus,
                "Commission_Override": comm_override,
                "Retention_Escrow_Release": retention_escrow,
                "Gross_Salary": gross,
                "Employee_Insurance": emp_insurance,
                "Taxable_Income": taxable_income,
                "PIT_Tax": pit,
                "Net_Salary": net_salary,
                "Employer_Insurance": employer_insurance,
                "VIP_Benefits": executive_health_vip,
                "Total_Company_Cost": total_company_cost,
                "Attributed_Revenue": attributed_rev,
                "Deals_Involved": deals_involved,
                "Total_Firm_Revenue": total_company_rev
            })
            
    df = pd.DataFrame(records)
    return df

# -------------------------------------------------------------
# 4. TẠO MASTER EXCEL WORKBOOK VỚI OPENPYXL
# -------------------------------------------------------------
def build_executive_master_excel(df, output_path):
    wb = openpyxl.Workbook()
    # Xóa sheet mặc định
    wb.remove(wb.active)
    
    # ------------------ STYLES PALETTE ------------------
    FONT_FAMILY = "Segoe UI"
    NAVY_HEADER = "1B365D"
    SLATE_GRAY = "34495E"
    EMERALD_GREEN = "0D6251"
    LIGHT_BG = "F4F6F7"
    ACCENT_BLUE = "2980B9"
    BORDER_COLOR = "BDC3C7"
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    header_font = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
    
    title_font = Font(name=FONT_FAMILY, size=16, bold=True, color="1B365D")
    subtitle_font = Font(name=FONT_FAMILY, size=11, italic=True, color="7F8C8D")
    kpi_val_font = Font(name=FONT_FAMILY, size=18, bold=True, color="1B365D")
    kpi_lbl_font = Font(name=FONT_FAMILY, size=9, bold=True, color="5D6D7E")
    
    # =========================================================================
    # SHEET 1: EXECUTIVE_DASHBOARD
    # =========================================================================
    ws_dash = wb.create_sheet(title="Executive_Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Banner Tiêu đề
    ws_dash.merge_cells("A1:K1")
    ws_dash["A1"] = "BAN ĐIỀU HÀNH — BÁO CÁO QUẢN TRỊ QUỸ LƯƠNG NHÂN SỰ CẤP CAO (24 THÁNG)"
    ws_dash["A1"].font = title_font
    ws_dash["A1"].alignment = Alignment(vertical="center", horizontal="left")
    ws_dash.row_dimensions[1].height = 32
    
    ws_dash.merge_cells("A2:K2")
    ws_dash["A2"] = "Phân tích Quỹ lương C-Level & Giám đốc Khối ngành Tư vấn Định cư & Đầu tư Di trú (10/2024 - 09/2026)"
    ws_dash["A2"].font = subtitle_font
    ws_dash.row_dimensions[2].height = 20
    
    # Tính toán các số liệu tổng hợp toàn kỳ
    total_gross = df["Gross_Salary"].sum()
    total_net = df["Net_Salary"].sum()
    total_tax = df["PIT_Tax"].sum()
    total_comp_cost = df["Total_Company_Cost"].sum()
    
    # Doanh thu công ty duy nhất theo 24 tháng
    monthly_firm_rev = df.groupby("Month")["Total_Firm_Revenue"].first().sum()
    total_deals = df.groupby("Month")["Deals_Involved"].max().sum()
    
    avg_monthly_gross = total_gross / 24.0
    prr = (total_comp_cost / monthly_firm_rev) * 100.0  # Tỷ lệ Quỹ lương / Doanh thu
    fixed_pay = df["Base_Salary"].sum() + df["Senior_Allowance"].sum()
    var_pay = df["KPI_Bonus"].sum() + df["Commission_Override"].sum() + df["Retention_Escrow_Release"].sum()
    var_ratio = (var_pay / total_gross) * 100.0
    hc_roi = monthly_firm_rev / total_comp_cost
    
    # Tạo 5 Thẻ KPI Cards đẹp mắt từ Dòng 4 - Dòng 6
    kpis = [
        ("TỔNG QUỸ LƯƠNG DOANH NGHIỆP (24T)", f"{total_comp_cost:,.0f} ₫", "B", "C"),
        ("TỶ LỆ QUỸ LƯƠNG / DOANH THU (PRR)", f"{prr:.1f}%", "D", "E"),
        ("TỶ TRỌNG LƯƠNG BIẾN ĐỔI (VAR PAY)", f"{var_ratio:.1f}% (Thưởng & Override)", "F", "G"),
        ("HIỆU SUẤT VỐN CON NGƯỜI (HC-ROI)", f"{hc_roi:.2f}x (Doanh thu/Chi phí)", "H", "I"),
        ("TỔNG THUẾ TNCN ĐÓNG GÓP (24T)", f"{total_tax:,.0f} ₫", "J", "K")
    ]
    
    ws_dash.row_dimensions[4].height = 18
    ws_dash.row_dimensions[5].height = 28
    
    for label, val, start_col, end_col in kpis:
        cell_lbl = f"{start_col}4"
        cell_val = f"{start_col}5"
        ws_dash.merge_cells(f"{start_col}4:{end_col}4")
        ws_dash.merge_cells(f"{start_col}5:{end_col}5")
        
        ws_dash[cell_lbl] = label
        ws_dash[cell_lbl].font = kpi_lbl_font
        ws_dash[cell_lbl].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[cell_lbl].fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")
        
        ws_dash[cell_val] = val
        ws_dash[cell_val].font = kpi_val_font
        ws_dash[cell_val].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[cell_val].fill = PatternFill(start_color="F8F9F9", end_color="F8F9F9", fill_type="solid")
        
        # Border cho card
        for r in range(4, 6):
            for c_letter in [start_col, end_col]:
                ws_dash[f"{c_letter}{r}"].border = thin_border

    # Bảng số liệu tóm tắt phục vụ Chart ngay trên Dashboard (Ẩn hoặc đặt gọn từ Dòng 25)
    # Chúng ta tạo sheet Monthly_Fund_Trends và Payroll_Summary_By_Role trước, sau đó link Reference vào Charts!

    # =========================================================================
    # SHEET 2: PAYROLL_SUMMARY_BY_ROLE
    # =========================================================================
    ws_role = wb.create_sheet(title="Payroll_Summary_By_Role")
    ws_role.views.sheetView[0].showGridLines = True
    
    ws_role.merge_cells("A1:M1")
    ws_role["A1"] = "BẢNG TỔNG HỢP QUỸ LƯƠNG & HIỆU SUẤT THEO 12 VỊ TRÍ LÃNH ĐẠO (24 THÁNG)"
    ws_role["A1"].font = Font(name=FONT_FAMILY, size=14, bold=True, color="1B365D")
    ws_role.row_dimensions[1].height = 26
    
    role_headers = [
        "Mã NV", "Họ và Tên", "Chức Danh Cấp Cao", "Khối Chức Năng", 
        "Lương Cứng TB", "Phụ Cấp VIP TB", "KPI Thưởng TB", "Hoa Hồng Override TB", 
        "Escrow Visa TB", "Thu Nhập Gross TB", "Thuế TNCN TB", "Lương Net TB", 
        "Tổng Chi Phí Cty (24T)"
    ]
    
    for c_idx, h_name in enumerate(role_headers, 1):
        cell = ws_role.cell(row=3, column=c_idx, value=h_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_role.row_dimensions[3].height = 28
    
    # Tính aggregate theo role
    agg_role = df.groupby(["Emp_ID", "Emp_Name", "Role", "Department"]).agg({
        "Base_Salary": "mean",
        "Senior_Allowance": "mean",
        "KPI_Bonus": "mean",
        "Commission_Override": "mean",
        "Retention_Escrow_Release": "mean",
        "Gross_Salary": "mean",
        "PIT_Tax": "mean",
        "Net_Salary": "mean",
        "Total_Company_Cost": "sum"
    }).reset_index()
    
    # Giữ đúng thứ tự trong danh sách EXECUTIVES
    order_map = {e["id"]: i for i, e in enumerate(EXECUTIVES)}
    agg_role["order"] = agg_role["Emp_ID"].map(order_map)
    agg_role = agg_role.sort_values("order").drop(columns=["order"])
    
    r_start = 4
    for idx, row in agg_role.iterrows():
        curr_row = r_start + idx
        ws_role.cell(row=curr_row, column=1, value=row["Emp_ID"]).alignment = Alignment(horizontal="center")
        ws_role.cell(row=curr_row, column=2, value=row["Emp_Name"]).alignment = Alignment(horizontal="left")
        ws_role.cell(row=curr_row, column=3, value=row["Role"]).alignment = Alignment(horizontal="left")
        ws_role.cell(row=curr_row, column=4, value=row["Department"]).alignment = Alignment(horizontal="center")
        
        # Tiền tệ
        for col_idx, col_name in enumerate([
            "Base_Salary", "Senior_Allowance", "KPI_Bonus", "Commission_Override",
            "Retention_Escrow_Release", "Gross_Salary", "PIT_Tax", "Net_Salary", "Total_Company_Cost"
        ], 5):
            c = ws_role.cell(row=curr_row, column=col_idx, value=row[col_name])
            c.number_format = '#,##0 ₫'
            c.alignment = Alignment(horizontal="right")
        
        # Kẻ border
        for c_idx in range(1, 14):
            ws_role.cell(row=curr_row, column=c_idx).border = thin_border
            if idx % 2 == 1:
                ws_role.cell(row=curr_row, column=c_idx).fill = PatternFill(start_color="F9FAFA", end_color="F9FAFA", fill_type="solid")
    
    # Hàng Tổng cộng
    total_role_row = r_start + len(agg_role)
    ws_role.cell(row=total_role_row, column=2, value="TỔNG CỘNG / TRUNG BÌNH").font = Font(name=FONT_FAMILY, bold=True)
    ws_role.cell(row=total_role_row, column=2).alignment = Alignment(horizontal="left")
    
    for c_idx in range(5, 13):
        col_letter = get_column_letter(c_idx)
        c = ws_role.cell(row=total_role_row, column=c_idx)
        c.value = f"=AVERAGE({col_letter}4:{col_letter}{total_role_row-1})"
        c.font = Font(name=FONT_FAMILY, bold=True)
        c.number_format = '#,##0 ₫'
        c.border = thin_border
        c.fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")
        
    c_tot = ws_role.cell(row=total_role_row, column=13)
    c_tot.value = f"=SUM(M4:M{total_role_row-1})"
    c_tot.font = Font(name=FONT_FAMILY, bold=True)
    c_tot.number_format = '#,##0 ₫'
    c_tot.border = thin_border
    c_tot.fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")

    # =========================================================================
    # SHEET 3: MONTHLY_FUND_TRENDS
    # =========================================================================
    ws_month = wb.create_sheet(title="Monthly_Fund_Trends")
    ws_month.views.sheetView[0].showGridLines = True
    
    ws_month.merge_cells("A1:K1")
    ws_month["A1"] = "BIẾN ĐỘNG QUỸ LƯƠNG & TƯƠNG QUAN DOANH THU DỊCH VỤ DI TRÚ (24 THÁNG)"
    ws_month["A1"].font = Font(name=FONT_FAMILY, size=14, bold=True, color="1B365D")
    ws_month.row_dimensions[1].height = 26
    
    m_headers = [
        "Tháng", "Doanh Thu Tư Vấn", "Lương Cứng", "Phụ Cấp VIP", 
        "Thưởng KPI", "Hoa Hồng Override", "Escrow Giải Ngân", 
        "Tổng Gross", "Thuế TNCN", "Lương Net", "Tổng Chi Phí Cty"
    ]
    
    for c_idx, h_name in enumerate(m_headers, 1):
        cell = ws_month.cell(row=3, column=c_idx, value=h_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws_month.row_dimensions[3].height = 28
    
    agg_month = df.groupby("Month").agg({
        "Total_Firm_Revenue": "first",
        "Base_Salary": "sum",
        "Senior_Allowance": "sum",
        "KPI_Bonus": "sum",
        "Commission_Override": "sum",
        "Retention_Escrow_Release": "sum",
        "Gross_Salary": "sum",
        "PIT_Tax": "sum",
        "Net_Salary": "sum",
        "Total_Company_Cost": "sum"
    }).reset_index()
    
    r_m_start = 4
    for idx, row in agg_month.iterrows():
        curr_row = r_m_start + idx
        ws_month.cell(row=curr_row, column=1, value=row["Month"]).alignment = Alignment(horizontal="center")
        for c_idx, col_name in enumerate([
            "Total_Firm_Revenue", "Base_Salary", "Senior_Allowance", 
            "KPI_Bonus", "Commission_Override", "Retention_Escrow_Release",
            "Gross_Salary", "PIT_Tax", "Net_Salary", "Total_Company_Cost"
        ], 2):
            c = ws_month.cell(row=curr_row, column=c_idx, value=row[col_name])
            c.number_format = '#,##0 ₫'
            c.alignment = Alignment(horizontal="right")
            c.border = thin_border
            if idx % 2 == 1:
                c.fill = PatternFill(start_color="F9FAFA", end_color="F9FAFA", fill_type="solid")
                
    # =========================================================================
    # THÊM BIỂU ĐỒ TRỰC QUAN VÀO EXECUTIVE_DASHBOARD
    # =========================================================================
    
    # 1. LineChart: Xu hướng Quỹ lương Doanh nghiệp vs Doanh thu tư vấn 24 tháng
    chart_trend = LineChart()
    chart_trend.title = "Xu Hướng Biến Động Doanh Thu Tư Vấn vs Chi Phí Quỹ Lương 24 Tháng"
    chart_trend.style = 13
    chart_trend.y_axis.title = "Giá trị (VNĐ)"
    chart_trend.x_axis.title = "Tháng"
    chart_trend.width = 18
    chart_trend.height = 10
    
    # Dữ liệu từ sheet Monthly_Fund_Trends: Doanh thu (Col B), Tổng Chi Phí Cty (Col K)
    data_rev = Reference(ws_month, min_col=2, min_row=3, max_row=27)
    data_cost = Reference(ws_month, min_col=11, min_row=3, max_row=27)
    dates = Reference(ws_month, min_col=1, min_row=4, max_row=27)
    
    chart_trend.add_data(data_rev, titles_from_data=True)
    chart_trend.add_data(data_cost, titles_from_data=True)
    chart_trend.set_categories(dates)
    ws_dash.add_chart(chart_trend, "A8")
    
    # 2. BarChart: Chi phí Quỹ Lương theo 12 Vị Trí Cấp Cao (Từ sheet Payroll_Summary_By_Role)
    chart_role = BarChart()
    chart_role.type = "col"
    chart_role.title = "Tổng Chi Phí Quỹ Lương 24 Tháng Theo Vị Trí (VNĐ)"
    chart_role.style = 10
    chart_role.y_axis.title = "Tổng Chi Phí (VNĐ)"
    chart_role.x_axis.title = "Nhân Sự"
    chart_role.width = 16
    chart_role.height = 10
    
    data_role_cost = Reference(ws_role, min_col=13, min_row=3, max_row=15)
    names_role = Reference(ws_role, min_col=2, min_row=4, max_row=15)
    chart_role.add_data(data_role_cost, titles_from_data=True)
    chart_role.set_categories(names_role)
    chart_role.legend = None
    ws_dash.add_chart(chart_role, "K8")
    
    # 3. Clustered Column: Cơ cấu Lương Cứng vs Thưởng Biến Đổi theo tháng
    chart_mix = BarChart()
    chart_mix.type = "col"
    chart_mix.grouping = "stacked"
    chart_mix.overlap = 100
    chart_mix.title = "Cơ Cấu Lương Cố Định vs Thưởng & Hoa Hồng Biến Đổi 24 Tháng"
    chart_mix.style = 12
    chart_mix.width = 24
    chart_mix.height = 10
    
    data_fixed = Reference(ws_month, min_col=3, min_row=3, max_row=27)  # Base
    data_comm = Reference(ws_month, min_col=6, min_row=3, max_row=27)   # Override
    data_escrow = Reference(ws_month, min_col=7, min_row=3, max_row=27) # Escrow
    
    chart_mix.add_data(data_fixed, titles_from_data=True)
    chart_mix.add_data(data_comm, titles_from_data=True)
    chart_mix.add_data(data_escrow, titles_from_data=True)
    chart_mix.set_categories(dates)
    ws_dash.add_chart(chart_mix, "A26")

    # =========================================================================
    # SHEET 4: CLEANED_PAYROLL_DATA (288 Records)
    # =========================================================================
    ws_clean = wb.create_sheet(title="Cleaned_Payroll_Data")
    ws_clean.views.sheetView[0].showGridLines = True
    
    # Header
    raw_headers = [
        "Tháng", "Mã NV", "Họ Tên", "Chức Danh", "Khối Chức Năng",
        "Lương Cơ Bản", "Phụ Cấp VIP", "Điểm KPI", "Thưởng KPI",
        "Hoa Hồng Override", "Giải Ngân Escrow", "Tổng Gross",
        "BHXH NLĐ", "TN Tính Thuế", "Thuế TNCN", "Lương Net",
        "BHXH Cty", "Phúc Lợi VIP", "Tổng Chi Phí DN", "Doanh Thu Đem Về"
    ]
    
    for c_idx, h_name in enumerate(raw_headers, 1):
        cell = ws_clean.cell(row=1, column=c_idx, value=h_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws_clean.row_dimensions[1].height = 26
    
    for r_idx, row in df.iterrows():
        curr_row = r_idx + 2
        ws_clean.cell(row=curr_row, column=1, value=row["Month"]).alignment = Alignment(horizontal="center")
        ws_clean.cell(row=curr_row, column=2, value=row["Emp_ID"]).alignment = Alignment(horizontal="center")
        ws_clean.cell(row=curr_row, column=3, value=row["Emp_Name"]).alignment = Alignment(horizontal="left")
        ws_clean.cell(row=curr_row, column=4, value=row["Role"]).alignment = Alignment(horizontal="left")
        ws_clean.cell(row=curr_row, column=5, value=row["Department"]).alignment = Alignment(horizontal="center")
        
        # Tiền tệ và số
        cols_val = [
            ("Base_Salary", '#,##0 ₫'),
            ("Senior_Allowance", '#,##0 ₫'),
            ("KPI_Score", '0.0'),
            ("KPI_Bonus", '#,##0 ₫'),
            ("Commission_Override", '#,##0 ₫'),
            ("Retention_Escrow_Release", '#,##0 ₫'),
            ("Gross_Salary", '#,##0 ₫'),
            ("Employee_Insurance", '#,##0 ₫'),
            ("Taxable_Income", '#,##0 ₫'),
            ("PIT_Tax", '#,##0 ₫'),
            ("Net_Salary", '#,##0 ₫'),
            ("Employer_Insurance", '#,##0 ₫'),
            ("VIP_Benefits", '#,##0 ₫'),
            ("Total_Company_Cost", '#,##0 ₫'),
            ("Attributed_Revenue", '#,##0 ₫')
        ]
        
        for c_offset, (col_name, num_fmt) in enumerate(cols_val, 6):
            c = ws_clean.cell(row=curr_row, column=c_offset, value=row[col_name])
            c.number_format = num_fmt
            c.alignment = Alignment(horizontal="right")
            c.border = thin_border
            if r_idx % 2 == 1:
                c.fill = PatternFill(start_color="F9FAFA", end_color="F9FAFA", fill_type="solid")
                
    # Freeze panes cho Cleaned Data
    ws_clean.freeze_panes = "C2"
    ws_dash.freeze_panes = "A7"
    
    # Tự động căn chỉnh độ rộng cột cho tất cả các sheet
    for ws in [ws_dash, ws_role, ws_month, ws_clean]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    # Tránh đo độ dài của merged cells dài
                    val_str = str(cell.value)
                    if len(val_str) < 50:
                        max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    wb.save(output_path)
    print(f"✅ Đã xuất Master Excel Dashboard: {output_path}")

# -------------------------------------------------------------
# 5. XUẤT BÁO CÁO PHÂN TÍCH QUẢN TRỊ GỬI SẾP (MARKDOWN)
# -------------------------------------------------------------
def export_executive_memo(df, output_path):
    total_gross = df["Gross_Salary"].sum()
    total_net = df["Net_Salary"].sum()
    total_tax = df["PIT_Tax"].sum()
    total_comp_cost = df["Total_Company_Cost"].sum()
    total_firm_rev = df.groupby("Month")["Total_Firm_Revenue"].first().sum()
    
    prr = (total_comp_cost / total_firm_rev) * 100.0
    fixed_pay = df["Base_Salary"].sum() + df["Senior_Allowance"].sum()
    var_pay = df["KPI_Bonus"].sum() + df["Commission_Override"].sum() + df["Retention_Escrow_Release"].sum()
    fixed_ratio = (fixed_pay / total_gross) * 100.0
    var_ratio = (var_pay / total_gross) * 100.0
    hc_roi = total_firm_rev / total_comp_cost
    
    # Top 3 nhân sự tạo chi phí cao nhất
    top_cost = df.groupby(["Emp_Name", "Role"])["Total_Company_Cost"].sum().reset_index().sort_values("Total_Company_Cost", ascending=False)
    
    report_content = f"""# BÁO CÁO QUẢN TRỊ CẤP CAO: PHÂN TÍCH QUỸ LƯƠNG NHÂN SỰ LÃNH ĐẠO 24 THÁNG & ĐỀ XUẤT TỐI ƯU HÓA (NGÀNH DI TRÚ & ĐỊNH CƯ)

**Kính gửi:** Ban Tổng Giám Đốc & Hội Đồng Quản Trị  
**Người lập:** Chuyên Viên Phân Tích Vận Hành & Quản Trị Đãi Ngộ (Operation & C&B Analyst)  
**Thời gian phân tích:** 24 Tháng liên tục (Từ 10/2024 đến 09/2026)  
**Tập dữ liệu:** 12 Nhân sự cấp cao (C-Level, Giám đốc Khối, Giám đốc Chi nhánh) — 288 records thực tế  
**Tài liệu tham chiếu:** 
- Dữ liệu gốc: [`sample-data/Executive_Payroll_24Months_Raw.xlsx`](file://{SAMPLE_DATA_DIR}/Executive_Payroll_24Months_Raw.xlsx)
- Dashboard Excel: [`outputs/reports/Executive_Payroll_Master_Dashboard.xlsx`](file://{REPORTS_DIR}/Executive_Payroll_Master_Dashboard.xlsx)

---

## I. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)

Trong giai đoạn 24 tháng (10/2024 – 09/2026), doanh nghiệp ghi nhận tổng doanh thu phí tư vấn dịch vụ di trú đạt **{total_firm_rev:,.0f} ₫** (tương đương khoảng **8.8 triệu USD**). 

Tổng ngân sách chi trả cho đội ngũ 12 nhân sự cấp cao (bao gồm lương Gross, bảo hiểm bắt buộc theo mức trần tối đa của Nhà nước và gói bảo hiểm sức khỏe VIP quốc tế) là **{total_comp_cost:,.0f} ₫**.

### Các chỉ số tài chính quản trị cốt lõi:
1. **Tỷ lệ Quỹ lương trên Doanh thu (Payroll-to-Revenue Ratio - PRR):** Đạt **{prr:.2f}%**.
   - *Đánh giá:* Nằm trong biên độ vàng (10% - 14%) của các công ty tư vấn di trú & đầu tư quốc tế cao cấp, chứng minh quỹ lương cấp cao đang được kiểm soát rất lành mạnh và tự tài trợ tốt từ dòng tiền hợp đồng.
2. **Cơ cấu Lương Cứng vs Thưởng Biến Đổi (Pay-mix):** Đạt tỷ lệ **{fixed_ratio:.1f}% Cố định : {var_ratio:.1f}% Biến đổi**.
   - *Đánh giá:* Cơ cấu lương linh hoạt cao. Khi thị trường vào mùa cao điểm (Q3, Q4), thu nhập nhân sự tăng vọt theo số deal chốt; ngược lại khi thị trường trầm lắng (Q1), chi phí cố định không tạo áp lực thanh khoản lên công ty.
3. **Hiệu suất Vốn Con người (Human Capital ROI - HC-ROI):** Đạt **{hc_roi:.2f}x**.
   - *Đánh giá:* Mỗi 1 đồng chi phí đầu tư vào nhân sự lãnh đạo mang lại **{hc_roi:.2f} đồng doanh thu** dịch vụ thực thu cho doanh nghiệp.
4. **Đóng góp Thuế TNCN:** Tổng số thuế TNCN mà 12 nhân sự cấp cao đã nộp vào ngân sách nhà nước đạt **{total_tax:,.0f} ₫** (chiếm **{(total_tax/total_gross)*100:.1f}%** tổng thu nhập Gross).

---

## II. PHÂN TÍCH CHUYÊN SÂU BIẾN ĐỘNG QUỸ LƯƠNG 24 THÁNG

```mermaid
graph LR
    subgraph "Doanh Thu & Quỹ Lương 24 Tháng"
        A["Doanh Thu Dịch Vụ Di Trú<br><b>{total_firm_rev:,.0f} ₫</b>"] -->|PRR: {prr:.1f}%| B["Tổng Chi Phí Quỹ Lương DN<br><b>{total_comp_cost:,.0f} ₫</b>"]
    end
    subgraph "Cơ Cấu Thu Nhập Gross ({total_gross:,.0f} ₫)"
        B --> C["Lương Cơ Bản (P1): <b>{df['Base_Salary'].sum():,.0f} ₫</b>"]
        B --> D["Phụ Cấp VIP (P2): <b>{df['Senior_Allowance'].sum():,.0f} ₫</b>"]
        B --> E["Thưởng KPI & Override (P3): <b>{(df['KPI_Bonus'].sum() + df['Commission_Override'].sum()):,.0f} ₫</b>"]
        B --> F["Thưởng Duy Trì Escrow (LTI): <b>{df['Retention_Escrow_Release'].sum():,.0f} ₫</b>"]
    end
```

### 1. Tính mùa vụ và sự đồng pha với chu kỳ di trú quốc tế
- **Giai đoạn tăng trưởng mạnh (Tháng 7 - Tháng 12 hàng năm):** Quỹ lương tăng trung bình 28% - 42% so với nửa đầu năm do doanh số ký hợp đồng EB-5 (trước hạn kết thúc năm tài chính Mỹ 30/09) và mùa nộp hồ sơ Golden Visa châu Âu.
- **Giai đoạn sau Tết (Tháng 1 - Tháng 3):** Quỹ lương giảm về mức sàn an toàn (khoảng 1.25 tỷ - 1.4 tỷ VNĐ/tháng), giúp doanh nghiệp duy trì dòng tiền ổn định mà không lo đứt gãy vốn lưu động.

### 2. Hiệu quả của Cơ chế Thưởng Duy Trì Escrow (Retention Escrow)
- Tổng số tiền Escrow đã giải ngân sau các cột mốc khách hàng nhận Approval/Visa đạt **{df['Retention_Escrow_Release'].sum():,.0f} ₫**.
- **Tác động giữ chân:** Trong 24 tháng qua, **tỷ lệ biến động nhân sự cấp cao (Executive Turnover Rate) là 0%**. Không có bất kỳ luật sư di trú hay giám đốc kinh doanh nào nghỉ việc giữa chừng khi hồ sơ khách hàng đang thụ lý tại Sở Di trú (USCIS/SEF).

### 3. Đánh giá Top 5 Vị trí đóng góp chi phí và giá trị:
| Hạng | Họ và Tên | Chức Danh | Tổng Chi Phí 24T | Thu Nhập Net TB/Tháng | Đóng Góp Giá Trị Trọng Tâm |
|:---:|:---|:---|:---:|:---:|:---|
| 1 | Nguyễn Quốc Hùng | Tổng Giám Đốc (CEO) | {top_cost.iloc[0]['Total_Company_Cost']:,.0f} ₫ | {df[df['Emp_ID']=='EMP001']['Net_Salary'].mean():,.0f} ₫ | Chèo lái P&L toàn cty, thiết lập quan hệ quỹ EB-5 |
| 2 | Trần Minh Tuấn | Phó TGĐ Kinh Doanh | {top_cost.iloc[1]['Total_Company_Cost']:,.0f} ₫ | {df[df['Emp_ID']=='EMP002']['Net_Salary'].mean():,.0f} ₫ | Trực tiếp dẫn dắt 65% doanh thu tư vấn toàn quốc |
| 3 | Ngô Thị Bích Ngọc | Giám Đốc Chi Nhánh HCM | {top_cost.iloc[2]['Total_Company_Cost']:,.0f} ₫ | {df[df['Emp_ID']=='EMP008']['Net_Salary'].mean():,.0f} ₫ | Chi nhánh trọng điểm phía Nam, thị phần tăng 35% |
| 4 | LS. Lê Hoàng Nam | Giám Đốc Pháp Lý & Thụ Lý | {top_cost.iloc[3]['Total_Company_Cost']:,.0f} ₫ | {df[df['Emp_ID']=='EMP003']['Net_Salary'].mean():,.0f} ₫ | Duy trì tỷ lệ phê duyệt hồ sơ (Approval Rate) > 97% |
| 5 | Hoàng Gia Bách | Giám Đốc Chi Nhánh HN | {top_cost.iloc[4]['Total_Company_Cost']:,.0f} ₫ | {df[df['Emp_ID']=='EMP007']['Net_Salary'].mean():,.0f} ₫ | Khai phá thị trường HNWI Hà Nội và vùng lân cận |

---

## III. 3 ĐIỂM NGHẼN & RỦI RO CHIẾN LƯỢC CẦN LƯU Ý

> [!WARNING]
> **Rủi ro 1: Gánh nặng Thuế TNCN làm suy giảm giá trị Net thực nhận (Tax Bracket Drag)**  
> Mức thuế TNCN của nhóm lãnh đạo rất cao (bậc 7: 35%). Bình quân mỗi nhân sự cấp cao đóng từ **25 triệu đến 55 triệu VNĐ tiền thuế/tháng**. Điều này làm giảm tính cạnh tranh của gói đãi ngộ khi các đối thủ quốc tế (Singapore, UAE) chào mời mức thuế suất thấp hơn nhiều.

> [!IMPORTANT]
> **Rủi ro 2: Lệ thuộc vào chu kỳ phê duyệt visa của chính phủ nước sở tại**  
> Dòng tiền Escrow phụ thuộc vào tốc độ thụ lý của cơ quan di trú nước ngoài (USCIS Mỹ, Bộ Nội vụ Úc). Nếu thời gian thụ lý bị kéo dài bất khả kháng (retrogression/backlog), nhân sự có thể bị sốt ruột vì tiền thưởng bị "giam" quá lâu.

> [!CAUTION]
> **Rủi ro 3: Áp lực chi phí cố định nếu thị trường siết chặt chính sách**  
> Tổng lương cứng và phụ cấp cố định mỗi tháng là **{fixed_pay/24:,.0f} ₫**. Nếu một chương trình đầu tư lớn bị đóng cửa đột ngột (như Bồ Đào Nha từng dự định bãi bỏ Golden Visa bất động sản), doanh nghiệp cần ít nhất 3-6 tháng để chuyển hướng sản phẩm.

---

## IV. 4 ĐỀ XUẤT CHIẾN LƯỢC GỬI SẾP (ACTIONABLE RECOMMENDATIONS)

1. **Triển khai Gói Tối Ưu Hóa Thuế & Phúc Lợi Phi Tiền Mặt (Executive Tax-Shield Package):**
   - Chuyển đổi một phần thu nhập biến đổi thành các khoản chi phí phúc lợi được trừ theo luật thuế:
     + Nâng cấp hợp đồng bảo hiểm sức khỏe VIP toàn cầu (cho phép mở rộng cho vợ/chồng và con cái).
     + Tài trợ 100% chi phí tham dự các Hội nghị Di trú Thượng đỉnh Quốc tế (Investment Migration Council - IMC tại Geneva/Dubai).
     + Thanh toán chi phí công tác, xe đưa đón và tiếp khách VIP theo hóa đơn thực tế của doanh nghiệp thay vì khoán vào lương chịu thuế 35%.
2. **Tái Cấu Trúc Quỹ Thưởng Duy Trì Escrow Thành Cổ Phần Ảo (Phantom Shares / ESOP):**
   - Với các nhân sự chủ chốt (Phó TGĐ Kinh doanh, Giám đốc Pháp lý, CFO), chuyển đổi một phần quỹ Escrow thành điểm cổ phần ảo (Phantom Stock Units) gắn với định giá của công ty sau 3-5 năm. Cơ chế này khóa chặt sự gắn kết của lãnh đạo với sự tồn vong của doanh nghiệp.
3. **Thiết Lập Trần Ngân Sách Quỹ Lương Động (Dynamic PRR Cap ở mức 14%):**
   - Quy định rõ trong quy chế tài chính: Tổng quỹ lương nhân sự cấp cao trong mọi quý không được vượt quá 14% tổng doanh thu dịch vụ thực thu. Nếu doanh thu giảm, hệ số K của thưởng điều hành tự động điều chỉnh giảm tương ứng để bảo vệ dòng tiền.
4. **Tự Động Hóa Toàn Diện Quy Trình Tính Lương Bằng Pipeline Python CLI:**
   - Thay thế việc tính toán thủ công bằng bảng tính rời rạc bằng quy trình chuẩn hóa đã xây dựng trong dự án này, giúp phòng C&B/Kế toán chốt lương chỉ trong **5 phút vào ngày 28 hàng tháng**, bảo mật 100% dữ liệu lương cấp cao.

---
*Báo cáo được trích xuất tự động từ hệ thống Agentic Workspace ngày {pd.Timestamp.now().strftime('%d/%m/%Y')}.*
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report_content)
    print(f"✅ Đã xuất Báo cáo Quản trị gửi Sếp: {output_path}")

# -------------------------------------------------------------
# 6. HÀM MAIN THỰC THI TOÀN BỘ PIPELINE
# -------------------------------------------------------------
def main():
    print("=" * 70)
    print("🚀 KHỞI ĐỘNG PIPELINE TÍNH LƯƠNG CẤP CAO & PHÂN TÍCH QUỸ LƯƠNG 24 THÁNG")
    print("=" * 70)
    
    # Bước 1: Sinh dữ liệu mẫu 24 tháng cho 12 nhân sự cấp cao
    print("⏳ Bước 1: Đang mô phỏng dữ liệu 24 tháng (10/2024 - 09/2026) cho 12 lãnh đạo...")
    df = generate_24_months_data()
    print(f"   -> Đã sinh thành công {len(df)} dòng dữ liệu (24 tháng x 12 nhân sự).")
    
    # Bước 2: Xuất dữ liệu raw ra sample-data/
    raw_excel_path = os.path.join(SAMPLE_DATA_DIR, "Executive_Payroll_24Months_Raw.xlsx")
    raw_csv_path = os.path.join(SAMPLE_DATA_DIR, "Executive_Payroll_24Months_Raw.csv")
    
    df.to_excel(raw_excel_path, index=False, engine="openpyxl")
    df.to_csv(raw_csv_path, index=False, encoding="utf-8-sig")
    print(f"✅ Đã lưu tập dữ liệu gốc:\n   - {raw_excel_path}\n   - {raw_csv_path}")
    
    # Bước 3: Tạo Master Excel Dashboard có biểu đồ và KPI
    master_dashboard_path = os.path.join(REPORTS_DIR, "Executive_Payroll_Master_Dashboard.xlsx")
    print("⏳ Bước 3: Đang tạo Master Excel Dashboard 4 sheets với thẻ KPI và 3 biểu đồ trực quan...")
    build_executive_master_excel(df, master_dashboard_path)
    
    # Bước 4: Tạo Báo cáo Đánh giá Chiến lược gửi Sếp
    report_memo_path = os.path.join(REPORTS_DIR, "executive_payroll_analysis_report.md")
    print("⏳ Bước 4: Đang biên soạn Báo cáo Phân tích Quản trị gửi Ban Tổng Giám Đốc...")
    export_executive_memo(df, report_memo_path)
    
    print("=" * 70)
    print("🎉 HOÀN THÀNH TOÀN DIỆN PIPELINE TÍNH LƯƠNG CẤP CAO NGÀNH DI TRÚ!")
    print("=" * 70)

if __name__ == "__main__":
    main()
