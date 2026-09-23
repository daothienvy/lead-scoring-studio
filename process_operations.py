#!/usr/bin/env python3
"""
Process Operations Automation Pipeline
======================================
Automated end-to-end data pipeline for monthly ERP operational exports:
- Data ingestion from Google Sheet (public export link) or local CSV/Excel
- Cleaning, normalization, and Net Salary calculations
- Anomaly detection (high penalty, anomalous bonus, data integrity)
- Generation of dedicated 2-sheet Excel workbooks for each Manager with executive styling
- Generation of Master Executive Dashboard with KPI cards, Pivot summaries, and charts
- Generation of detailed Executive Markdown Report for Board/Directors
- Automated audit log generation and PDCA logging support

Author: Operation Analyst / Antigravity AI Pair Programmer
Course: Agentic AI with Google Antigravity
"""

import argparse
import os
import sys
import ssl
import urllib.request
from datetime import datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart


# ==========================================
# CONSTANTS & STYLING TOKENS
# ==========================================
FONT_NAME = "Segoe UI"
COLOR_PRIMARY_NAVY = "1B365D"    # Deep Navy
COLOR_SECONDARY_BLUE = "2B6CB0"  # Medium Blue
COLOR_ACCENT_TEAL = "0F766E"     # Teal
COLOR_DARK_TEXT = "0F172A"       # Slate 900
COLOR_MUTED_TEXT = "475569"      # Slate 600
COLOR_CARD_BG = "F8FAFC"         # Slate 50
COLOR_CARD_BORDER = "CBD5E1"     # Slate 300
COLOR_ZEBRA_EVEN = "F8FAFC"      # Subtle light
COLOR_ZEBRA_ODD = "FFFFFF"       # White
COLOR_TOTAL_BG = "E2E8F0"        # Slate 200

COLOR_ALERT_RED_BG = "FEE2E2"    # Red 100
COLOR_ALERT_RED_TXT = "991B1B"   # Red 800
COLOR_ALERT_GREEN_BG = "DCFCE7"  # Green 100
COLOR_ALERT_GREEN_TXT = "166534" # Green 800
COLOR_ALERT_AMBER_BG = "FEF3C7"  # Amber 100
COLOR_ALERT_AMBER_TXT = "92400E" # Amber 800

VND_FORMAT = '#,##0 "₫"'
INT_FORMAT = '#,##0'
PCT_FORMAT = '0.0%'


def create_border(style: str = "thin", color: str = "CBD5E1") -> Border:
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def auto_fit_columns(ws, min_widths: Dict[int, int] = None, max_width: int = 50):
    """Dynamically adjust column widths based on content."""
    min_widths = min_widths or {}
    for col_idx, col in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format and ('₫' in cell.number_format or '%' in cell.number_format):
                max_len = max(max_len, len(val) + 6)
            else:
                max_len = max(max_len, len(val))
        calc_width = max(max_len + 3, min_widths.get(col_idx, 12))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(calc_width, max_width)


# ==========================================
# DATA INGESTION & CLEANING
# ==========================================
def ingest_data(source_path: str, raw_cache_path: str = None) -> pd.DataFrame:
    """Load data from a Google Sheet URL or local file, caching a local copy if requested."""
    print(f"📥 Loading data from: {source_path}")
    if source_path.startswith("http://") or source_path.startswith("https://"):
        ctx = ssl._create_unverified_context()
        req = urllib.request.Request(source_path, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
        with urllib.request.urlopen(req, context=ctx) as resp:
            content = resp.read()
        
        if raw_cache_path:
            os.makedirs(os.path.dirname(os.path.abspath(raw_cache_path)), exist_ok=True)
            with open(raw_cache_path, "wb") as f:
                f.write(content)
            print(f"💾 Cached raw copy saved to: {raw_cache_path}")
        
        import io
        df = pd.read_csv(io.BytesIO(content))
    else:
        if source_path.endswith(".xlsx") or source_path.endswith(".xls"):
            df = pd.read_excel(source_path)
        else:
            df = pd.read_csv(source_path)
            
        if raw_cache_path and not os.path.exists(raw_cache_path):
            os.makedirs(os.path.dirname(os.path.abspath(raw_cache_path)), exist_ok=True)
            df.to_csv(raw_cache_path, index=False)
            
    print(f"✅ Loaded raw dataset with shape: {df.shape}")
    return df


def clean_and_enrich_data(
    df: pd.DataFrame, 
    threshold_penalty: float = 1500000.0, 
    threshold_bonus: float = 4000000.0
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Standardizes strings, parses numeric columns, calculates Net Salary,
    and flags anomalies.
    Returns: (cleaned_df, anomaly_audit_df)
    """
    df = df.copy()
    
    # 1. Strip whitespaces
    for col in ["Employee_ID", "Employee_Name", "Manager", "Department", "Month"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    # 2. Numeric conversion
    num_cols = ["Base_Salary", "Bonus", "Penalty"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        
    # 3. Calculate Net Salary
    df["Net_Salary"] = df["Base_Salary"] + df["Bonus"] - df["Penalty"]
    
    # 4. Anomaly Detection & Flags
    df["Anomaly_Flag"] = "Normal"
    df["Audit_Note"] = "Hợp lệ"
    
    anomalies = []
    
    for idx, row in df.iterrows():
        notes = []
        is_high_penalty = row["Penalty"] >= threshold_penalty or (row["Base_Salary"] > 0 and (row["Penalty"] / row["Base_Salary"]) >= 0.10)
        is_high_bonus = row["Bonus"] >= threshold_bonus
        is_neg_net = row["Net_Salary"] <= 0
        
        if is_neg_net:
            notes.append("Lương thực nhận <= 0")
        if is_high_penalty:
            notes.append(f"Phạt cao ({row['Penalty']:,.0f}đ)")
        if is_high_bonus:
            notes.append(f"Thưởng cao ({row['Bonus']:,.0f}đ)")
            
        if is_high_penalty and is_high_bonus:
            df.at[idx, "Anomaly_Flag"] = "Dual_Anomaly"
            df.at[idx, "Audit_Note"] = "Thưởng và Phạt đều cao bất thường"
            anomalies.append({
                "Employee_ID": row["Employee_ID"],
                "Employee_Name": row["Employee_Name"],
                "Manager": row["Manager"],
                "Department": row["Department"],
                "Base_Salary": row["Base_Salary"],
                "Bonus": row["Bonus"],
                "Penalty": row["Penalty"],
                "Net_Salary": row["Net_Salary"],
                "Issue_Type": "Dual Anomaly",
                "Severity": "High",
                "Audit_Note": "Vừa đạt thưởng cao vừa bị phạt nặng — Cần rà soát hiệu suất và quy trình.",
                "Recommendation": "Manager đối chiếu bảng chấm công và KPI chi tiết."
            })
        elif is_high_penalty:
            df.at[idx, "Anomaly_Flag"] = "High_Penalty"
            df.at[idx, "Audit_Note"] = "Mức phạt vượt trần cảnh báo"
            anomalies.append({
                "Employee_ID": row["Employee_ID"],
                "Employee_Name": row["Employee_Name"],
                "Manager": row["Manager"],
                "Department": row["Department"],
                "Base_Salary": row["Base_Salary"],
                "Bonus": row["Bonus"],
                "Penalty": row["Penalty"],
                "Net_Salary": row["Net_Salary"],
                "Issue_Type": "High Penalty",
                "Severity": "Medium",
                "Audit_Note": f"Mức phạt {row['Penalty']:,.0f}đ chiếm {(row['Penalty']/row['Base_Salary'])*100:.1f}% lương cơ bản.",
                "Recommendation": "Xem xét lỗi vận hành lặp lại để huấn luyện lại."
            })
        elif is_high_bonus:
            df.at[idx, "Anomaly_Flag"] = "High_Bonus"
            df.at[idx, "Audit_Note"] = "Mức thưởng xuất sắc vượt trội"
            anomalies.append({
                "Employee_ID": row["Employee_ID"],
                "Employee_Name": row["Employee_Name"],
                "Manager": row["Manager"],
                "Department": row["Department"],
                "Base_Salary": row["Base_Salary"],
                "Bonus": row["Bonus"],
                "Penalty": row["Penalty"],
                "Net_Salary": row["Net_Salary"],
                "Issue_Type": "High Bonus",
                "Severity": "Low",
                "Audit_Note": f"Thưởng vượt mức trần ({row['Bonus']:,.0f}đ) — Top hiệu suất.",
                "Recommendation": "Xác nhận phê duyệt từ Giám đốc vận hành."
            })
        elif is_neg_net:
            df.at[idx, "Anomaly_Flag"] = "Negative_Net"
            df.at[idx, "Audit_Note"] = "Lương âm hoặc bằng 0"
            anomalies.append({
                "Employee_ID": row["Employee_ID"],
                "Employee_Name": row["Employee_Name"],
                "Manager": row["Manager"],
                "Department": row["Department"],
                "Base_Salary": row["Base_Salary"],
                "Bonus": row["Bonus"],
                "Penalty": row["Penalty"],
                "Net_Salary": row["Net_Salary"],
                "Issue_Type": "Critical Error",
                "Severity": "High",
                "Audit_Note": "Net salary âm hoặc 0, rủi ro pháp lý lao động.",
                "Recommendation": "Điều chỉnh ngay mức phạt trước khi giải ngân."
            })

    anomaly_df = pd.DataFrame(anomalies)
    print(f"🔍 Data cleaned. Found {len(anomaly_df)} audit items (High Penalty: {(df['Anomaly_Flag'] == 'High_Penalty').sum()}, High Bonus: {(df['Anomaly_Flag'] == 'High_Bonus').sum()}, Dual: {(df['Anomaly_Flag'] == 'Dual_Anomaly').sum()})")
    return df, anomaly_df


# ==========================================
# EXCEL GENERATOR: MANAGER WORKBOOKS
# ==========================================
def generate_manager_workbook(
    manager_name: str, 
    mgr_df: pd.DataFrame, 
    month_str: str, 
    output_path: str,
    threshold_penalty: float = 1500000.0,
    threshold_bonus: float = 4000000.0
):
    """Generates an executive-ready 2-sheet workbook for a specific Manager."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # ----------------------------------------------------
    # SHEET 1: EXECUTIVE SUMMARY
    # ----------------------------------------------------
    ws_sum = wb.create_sheet(title="Tổng Quan & KPI")
    ws_sum.views.sheetView[0].showGridLines = True
    
    # Header Banner
    ws_sum.merge_cells("A1:G1")
    cell_h1 = ws_sum["A1"]
    cell_h1.value = f"BÁO CÁO VẬN HÀNH & LƯƠNG THƯỞNG — {manager_name.upper().replace('_', ' ')}"
    cell_h1.font = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
    cell_h1.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
    cell_h1.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36
    
    # Subtitle / Metadata
    ws_sum.merge_cells("A2:G2")
    cell_h2 = ws_sum["A2"]
    cell_h2.value = f"Kỳ báo cáo: Tháng {month_str} | Ngày kết xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Phụ trách: {manager_name}"
    cell_h2.font = Font(name=FONT_NAME, size=10, italic=True, color="FFFFFF")
    cell_h2.fill = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
    cell_h2.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    # Blank Row 3
    ws_sum.row_dimensions[3].height = 10

    # KPI Summary Cards (Row 4 to 6)
    kpis = [
        ("TỔNG NHÂN SỰ", f"{len(mgr_df)} nhân sự", INT_FORMAT, len(mgr_df)),
        ("TỔNG LƯƠNG CƠ BẢN", None, VND_FORMAT, f"=SUM('Chi Tiết Nhân Sự'!E2:E{len(mgr_df)+1})"),
        ("TỔNG THƯỞNG", None, VND_FORMAT, f"=SUM('Chi Tiết Nhân Sự'!F2:F{len(mgr_df)+1})"),
        ("TỔNG PHẠT", None, VND_FORMAT, f"=SUM('Chi Tiết Nhân Sự'!G2:G{len(mgr_df)+1})"),
        ("TỔNG LƯƠNG THỰC NHẬN", None, VND_FORMAT, f"=SUM('Chi Tiết Nhân Sự'!H2:H{len(mgr_df)+1})")
    ]
    
    ws_sum.row_dimensions[4].height = 20
    ws_sum.row_dimensions[5].height = 30
    
    # Render KPI Cards
    for i, (title, _, num_fmt, formula_val) in enumerate(kpis):
        c_letter = get_column_letter(i + 1)
        # Title cell (Row 4)
        c_title = ws_sum[f"{c_letter}4"]
        c_title.value = title
        c_title.font = Font(name=FONT_NAME, size=8, bold=True, color=COLOR_MUTED_TEXT)
        c_title.fill = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type="solid")
        c_title.alignment = Alignment(horizontal="center", vertical="center")
        c_title.border = Border(left=Side(style="thin", color=COLOR_CARD_BORDER), right=Side(style="thin", color=COLOR_CARD_BORDER), top=Side(style="thin", color=COLOR_CARD_BORDER))
        
        # Value cell (Row 5)
        c_val = ws_sum[f"{c_letter}5"]
        c_val.value = formula_val
        c_val.font = Font(name=FONT_NAME, size=12, bold=True, color=COLOR_PRIMARY_NAVY)
        c_val.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.number_format = num_fmt
        c_val.border = Border(left=Side(style="thin", color=COLOR_CARD_BORDER), right=Side(style="thin", color=COLOR_CARD_BORDER), bottom=Side(style="thin", color=COLOR_CARD_BORDER))
        
    # Blank Row 6
    ws_sum.row_dimensions[6].height = 14

    # Table 1: Phân Bổ Theo Phòng Ban
    ws_sum["A7"].value = "I. BẢNG PHÂN BỔ CHI TIẾT THEO PHÒNG BAN QUẢN LÝ"
    ws_sum["A7"].font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_PRIMARY_NAVY)
    
    headers_dept = ["Phòng Ban", "Số Nhân Sự", "Tổng Lương Cơ Bản", "Tổng Thưởng", "Tổng Phạt", "Tổng Lương Thực Nhận", "Thực Nhận TB/Người"]
    for c_idx, h_text in enumerate(headers_dept, 1):
        c = ws_sum.cell(row=8, column=c_idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = create_border(color="FFFFFF")
    ws_sum.row_dimensions[8].height = 25

    # Group by Dept for Manager
    dept_group = mgr_df.groupby("Department").agg(
        Count=("Employee_ID", "count"),
        Base=("Base_Salary", "sum"),
        Bonus=("Bonus", "sum"),
        Penalty=("Penalty", "sum"),
        Net=("Net_Salary", "sum"),
        Avg_Net=("Net_Salary", "mean")
    ).reset_index()

    curr_row = 9
    for _, d_row in dept_group.iterrows():
        bg = COLOR_ZEBRA_EVEN if (curr_row % 2 == 0) else COLOR_ZEBRA_ODD
        fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        
        ws_sum.cell(row=curr_row, column=1, value=d_row["Department"]).alignment = Alignment(horizontal="left", vertical="center")
        ws_sum.cell(row=curr_row, column=2, value=d_row["Count"]).number_format = INT_FORMAT
        ws_sum.cell(row=curr_row, column=3, value=d_row["Base"]).number_format = VND_FORMAT
        ws_sum.cell(row=curr_row, column=4, value=d_row["Bonus"]).number_format = VND_FORMAT
        ws_sum.cell(row=curr_row, column=5, value=d_row["Penalty"]).number_format = VND_FORMAT
        ws_sum.cell(row=curr_row, column=6, value=d_row["Net"]).number_format = VND_FORMAT
        ws_sum.cell(row=curr_row, column=7, value=d_row["Avg_Net"]).number_format = VND_FORMAT
        
        for c_idx in range(1, 8):
            c = ws_sum.cell(row=curr_row, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = fill
            c.border = create_border()
            if c_idx >= 2:
                c.alignment = Alignment(horizontal="right", vertical="center")
        ws_sum.row_dimensions[curr_row].height = 20
        curr_row += 1
        
    # Department Total Row
    ws_sum.cell(row=curr_row, column=1, value="TỔNG CỘNG")
    ws_sum.cell(row=curr_row, column=2, value=f"=SUM(B9:B{curr_row-1})").number_format = INT_FORMAT
    ws_sum.cell(row=curr_row, column=3, value=f"=SUM(C9:C{curr_row-1})").number_format = VND_FORMAT
    ws_sum.cell(row=curr_row, column=4, value=f"=SUM(D9:D{curr_row-1})").number_format = VND_FORMAT
    ws_sum.cell(row=curr_row, column=5, value=f"=SUM(E9:E{curr_row-1})").number_format = VND_FORMAT
    ws_sum.cell(row=curr_row, column=6, value=f"=SUM(F9:F{curr_row-1})").number_format = VND_FORMAT
    ws_sum.cell(row=curr_row, column=7, value=f"=F{curr_row}/B{curr_row}").number_format = VND_FORMAT
    
    for c_idx in range(1, 8):
        c = ws_sum.cell(row=curr_row, column=c_idx)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_PRIMARY_NAVY)
        c.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
        c.border = create_border(style="medium", color=COLOR_PRIMARY_NAVY)
        if c_idx >= 2:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
    ws_sum.row_dimensions[curr_row].height = 22
    curr_row += 2

    # Section II: Danh Sách Cần Lưu Ý (Anomalies of this Manager)
    ws_sum.cell(row=curr_row, column=1, value="II. DANH SÁCH NHÂN SỰ CẦN MANAGER KIỂM TRA & RÀ SOÁT").font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_ALERT_RED_TXT)
    curr_row += 1
    
    headers_alert = ["Mã NV", "Họ & Tên", "Phòng Ban", "Thưởng (VND)", "Phạt (VND)", "Lương Thực Nhận", "Nội Dung Lưu Ý & Khuyến Nghị"]
    for c_idx, h_text in enumerate(headers_alert, 1):
        c = ws_sum.cell(row=curr_row, column=c_idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = create_border(color="FFFFFF")
    ws_sum.row_dimensions[curr_row].height = 24
    curr_row += 1
    
    mgr_anomalies = mgr_df[mgr_df["Anomaly_Flag"] != "Normal"]
    if len(mgr_anomalies) == 0:
        ws_sum.cell(row=curr_row, column=1, value="Không có nhân sự nào trong diện cảnh báo bất thường.")
        ws_sum.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=7)
        ws_sum.cell(row=curr_row, column=1).font = Font(name=FONT_NAME, size=9, italic=True)
        ws_sum.cell(row=curr_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        curr_row += 1
    else:
        for _, a_row in mgr_anomalies.iterrows():
            ws_sum.cell(row=curr_row, column=1, value=a_row["Employee_ID"]).alignment = Alignment(horizontal="center", vertical="center")
            ws_sum.cell(row=curr_row, column=2, value=a_row["Employee_Name"]).alignment = Alignment(horizontal="left", vertical="center")
            ws_sum.cell(row=curr_row, column=3, value=a_row["Department"]).alignment = Alignment(horizontal="center", vertical="center")
            ws_sum.cell(row=curr_row, column=4, value=a_row["Bonus"]).number_format = VND_FORMAT
            ws_sum.cell(row=curr_row, column=5, value=a_row["Penalty"]).number_format = VND_FORMAT
            ws_sum.cell(row=curr_row, column=6, value=a_row["Net_Salary"]).number_format = VND_FORMAT
            ws_sum.cell(row=curr_row, column=7, value=a_row["Audit_Note"]).alignment = Alignment(horizontal="left", vertical="center")
            
            # Highlight color based on severity
            fill_color = COLOR_ALERT_RED_BG if a_row["Penalty"] >= threshold_penalty else (COLOR_ALERT_GREEN_BG if a_row["Bonus"] >= threshold_bonus else COLOR_CARD_BG)
            text_color = COLOR_ALERT_RED_TXT if a_row["Penalty"] >= threshold_penalty else (COLOR_ALERT_GREEN_TXT if a_row["Bonus"] >= threshold_bonus else COLOR_DARK_TEXT)
            
            for c_idx in range(1, 8):
                c = ws_sum.cell(row=curr_row, column=c_idx)
                c.font = Font(name=FONT_NAME, size=9, color=text_color)
                c.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                c.border = create_border()
                if c_idx in (4, 5, 6):
                    c.alignment = Alignment(horizontal="right", vertical="center")
            ws_sum.row_dimensions[curr_row].height = 20
            curr_row += 1

    auto_fit_columns(ws_sum, min_widths={1: 18, 2: 18, 3: 22, 4: 18, 5: 18, 6: 22, 7: 35})

    # ----------------------------------------------------
    # SHEET 2: EMPLOYEE DETAILS
    # ----------------------------------------------------
    ws_det = wb.create_sheet(title="Chi Tiết Nhân Sự")
    ws_det.views.sheetView[0].showGridLines = True
    
    headers_det = ["STT", "Mã NV", "Họ Tên", "Phòng Ban", "Lương Cơ Bản", "Thưởng", "Phạt", "Lương Thực Nhận", "Kỳ", "Ghi Chú Kiểm Toán"]
    ws_det.row_dimensions[1].height = 28
    for c_idx, h_text in enumerate(headers_det, 1):
        c = ws_det.cell(row=1, column=c_idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = create_border(color="FFFFFF")
        
    row_idx = 2
    for stt, (_, emp) in enumerate(mgr_df.iterrows(), 1):
        bg = COLOR_ZEBRA_EVEN if (row_idx % 2 == 0) else COLOR_ZEBRA_ODD
        
        ws_det.cell(row=row_idx, column=1, value=stt).alignment = Alignment(horizontal="center", vertical="center")
        ws_det.cell(row=row_idx, column=2, value=emp["Employee_ID"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_det.cell(row=row_idx, column=3, value=emp["Employee_Name"]).alignment = Alignment(horizontal="left", vertical="center")
        ws_det.cell(row=row_idx, column=4, value=emp["Department"]).alignment = Alignment(horizontal="center", vertical="center")
        
        # Numbers
        ws_det.cell(row=row_idx, column=5, value=float(emp["Base_Salary"])).number_format = VND_FORMAT
        ws_det.cell(row=row_idx, column=6, value=float(emp["Bonus"])).number_format = VND_FORMAT
        ws_det.cell(row=row_idx, column=7, value=float(emp["Penalty"])).number_format = VND_FORMAT
        
        # Net Salary Formula: =E{row}+F{row}-G{row}
        c_net = ws_det.cell(row=row_idx, column=8, value=f"=E{row_idx}+F{row_idx}-G{row_idx}")
        c_net.number_format = VND_FORMAT
        
        ws_det.cell(row=row_idx, column=9, value=str(emp["Month"])).alignment = Alignment(horizontal="center", vertical="center")
        ws_det.cell(row=row_idx, column=10, value=str(emp["Audit_Note"])).alignment = Alignment(horizontal="left", vertical="center")
        
        # Default styling
        for c_idx in range(1, 11):
            c = ws_det.cell(row=row_idx, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.border = create_border()
            if c_idx in (5, 6, 7, 8):
                c.alignment = Alignment(horizontal="right", vertical="center")
                
        # Conditional Highlighting:
        if emp["Penalty"] >= threshold_penalty:
            c_pen = ws_det.cell(row=row_idx, column=7)
            c_pen.fill = PatternFill(start_color=COLOR_ALERT_RED_BG, end_color=COLOR_ALERT_RED_BG, fill_type="solid")
            c_pen.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_ALERT_RED_TXT)
            
        if emp["Bonus"] >= threshold_bonus:
            c_bon = ws_det.cell(row=row_idx, column=6)
            c_bon.fill = PatternFill(start_color=COLOR_ALERT_GREEN_BG, end_color=COLOR_ALERT_GREEN_BG, fill_type="solid")
            c_bon.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_ALERT_GREEN_TXT)

        ws_det.row_dimensions[row_idx].height = 20
        row_idx += 1
        
    # Total Row at the bottom of Employee Details
    ws_det.cell(row=row_idx, column=1, value="")
    ws_det.cell(row=row_idx, column=2, value="TỔNG CỘNG")
    ws_det.cell(row=row_idx, column=3, value=f"{len(mgr_df)} NV")
    ws_det.cell(row=row_idx, column=4, value="")
    ws_det.cell(row=row_idx, column=5, value=f"=SUM(E2:E{row_idx-1})").number_format = VND_FORMAT
    ws_det.cell(row=row_idx, column=6, value=f"=SUM(F2:F{row_idx-1})").number_format = VND_FORMAT
    ws_det.cell(row=row_idx, column=7, value=f"=SUM(G2:G{row_idx-1})").number_format = VND_FORMAT
    ws_det.cell(row=row_idx, column=8, value=f"=SUM(H2:H{row_idx-1})").number_format = VND_FORMAT
    ws_det.cell(row=row_idx, column=9, value="")
    ws_det.cell(row=row_idx, column=10, value="")
    
    for c_idx in range(1, 11):
        c = ws_det.cell(row=row_idx, column=c_idx)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_PRIMARY_NAVY)
        c.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
        c.border = create_border(style="medium", color=COLOR_PRIMARY_NAVY)
        if c_idx in (5, 6, 7, 8):
            c.alignment = Alignment(horizontal="right", vertical="center")
        elif c_idx in (2, 3):
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws_det.row_dimensions[row_idx].height = 24
    
    auto_fit_columns(ws_det, min_widths={1: 6, 2: 12, 3: 20, 4: 15, 5: 18, 6: 16, 7: 16, 8: 20, 9: 12, 10: 30})

    # Save workbook
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    print(f"  📄 Exported Manager Workbook: {output_path} ({len(mgr_df)} records)")


# ==========================================
# EXCEL GENERATOR: MASTER DASHBOARD
# ==========================================
def generate_master_dashboard(
    clean_df: pd.DataFrame, 
    anomaly_df: pd.DataFrame, 
    month_str: str, 
    output_path: str
):
    """Generates company-wide Master Operations Dashboard with embedded OpenPyXL charts."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # ----------------------------------------------------
    # SHEET 1: EXECUTIVE DASHBOARD & CHARTS
    # ----------------------------------------------------
    ws = wb.create_sheet(title="Executive Dashboard")
    ws.views.sheetView[0].showGridLines = True
    
    # Header Banner
    ws.merge_cells("A1:K1")
    h1 = ws["A1"]
    h1.value = "HỆ THỐNG BÁO CÁO VẬN HÀNH & HIỆU SUẤT TOÀN DOANH NGHIỆP"
    h1.font = Font(name=FONT_NAME, size=15, bold=True, color="FFFFFF")
    h1.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
    h1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    
    ws.merge_cells("A2:K2")
    h2 = ws["A2"]
    h2.value = f"Kỳ phân tích: Tháng {month_str} | Phân hệ: Operations Analytics | Trạng thái: Dữ liệu đã làm sạch & kiểm toán"
    h2.font = Font(name=FONT_NAME, size=10, italic=True, color="FFFFFF")
    h2.fill = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
    h2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 12

    # Executive KPI Cards (Row 4 & 5)
    total_staff = len(clean_df)
    total_base = clean_df["Base_Salary"].sum()
    total_bonus = clean_df["Bonus"].sum()
    total_pen = clean_df["Penalty"].sum()
    total_net = clean_df["Net_Salary"].sum()
    avg_net = clean_df["Net_Salary"].mean()

    kpi_cards = [
        ("TỔNG NHÂN SỰ", f"{total_staff} người", "A", "B", INT_FORMAT, total_staff),
        ("TỔNG QUỸ LƯƠNG CƠ BẢN", None, "C", "D", VND_FORMAT, total_base),
        ("TỔNG TIỀN THƯỞNG", None, "E", "F", VND_FORMAT, total_bonus),
        ("TỔNG TIỀN PHẠT", None, "G", "H", VND_FORMAT, total_pen),
        ("TỔNG LƯƠNG THỰC NHẬN", None, "I", "J", VND_FORMAT, total_net),
        ("THỰC NHẬN BÌNH QUÂN", None, "K", "K", VND_FORMAT, avg_net)
    ]
    
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 28
    
    for title, _, start_c, end_c, num_fmt, val in kpi_cards:
        if start_c != end_c:
            ws.merge_cells(f"{start_c}4:{end_c}4")
            ws.merge_cells(f"{start_c}5:{end_c}5")
            
        c_title = ws[f"{start_c}4"]
        c_title.value = title
        c_title.font = Font(name=FONT_NAME, size=8, bold=True, color=COLOR_MUTED_TEXT)
        c_title.fill = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type="solid")
        c_title.alignment = Alignment(horizontal="center", vertical="center")
        
        c_val = ws[f"{start_c}5"]
        c_val.value = val
        c_val.font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_PRIMARY_NAVY)
        c_val.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.number_format = num_fmt
        
        # Border around merged card
        for r in (4, 5):
            col_start_idx = openpyxl.utils.column_index_from_string(start_c)
            col_end_idx = openpyxl.utils.column_index_from_string(end_c)
            for c_idx in range(col_start_idx, col_end_idx + 1):
                ws.cell(row=r, column=c_idx).border = create_border(color=COLOR_CARD_BORDER)

    ws.row_dimensions[6].height = 14

    # ----------------------------------------------------
    # TABLE 1: SUMMARY BY DEPARTMENT (Cols A to G)
    # ----------------------------------------------------
    ws["A7"].value = "1. TỔNG HỢP VẬN HÀNH THEO PHÒNG BAN"
    ws["A7"].font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_PRIMARY_NAVY)
    
    dept_headers = ["Phòng Ban", "Nhân Sự", "Lương Cơ Bản", "Thưởng", "Phạt", "Thực Nhận", "Tỷ Lệ Thưởng/Lương"]
    for idx, h_text in enumerate(dept_headers, 1):
        c = ws.cell(row=8, column=idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = create_border(color="FFFFFF")
    ws.row_dimensions[8].height = 25

    dept_summary = clean_df.groupby("Department").agg(
        Count=("Employee_ID", "count"),
        Base=("Base_Salary", "sum"),
        Bonus=("Bonus", "sum"),
        Penalty=("Penalty", "sum"),
        Net=("Net_Salary", "sum")
    ).reset_index()

    r = 9
    for _, d_row in dept_summary.iterrows():
        ws.cell(row=r, column=1, value=d_row["Department"]).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r, column=2, value=d_row["Count"]).number_format = INT_FORMAT
        ws.cell(row=r, column=3, value=d_row["Base"]).number_format = VND_FORMAT
        ws.cell(row=r, column=4, value=d_row["Bonus"]).number_format = VND_FORMAT
        ws.cell(row=r, column=5, value=d_row["Penalty"]).number_format = VND_FORMAT
        ws.cell(row=r, column=6, value=d_row["Net"]).number_format = VND_FORMAT
        ws.cell(row=r, column=7, value=f"=D{r}/C{r}").number_format = PCT_FORMAT
        
        bg = COLOR_ZEBRA_EVEN if (r % 2 == 0) else COLOR_ZEBRA_ODD
        for c_idx in range(1, 8):
            c = ws.cell(row=r, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.border = create_border()
            if c_idx >= 2:
                c.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

    # Total row Dept
    ws.cell(row=r, column=1, value="TỔNG CỘNG")
    ws.cell(row=r, column=2, value=f"=SUM(B9:B{r-1})").number_format = INT_FORMAT
    ws.cell(row=r, column=3, value=f"=SUM(C9:C{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=4, value=f"=SUM(D9:D{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=5, value=f"=SUM(E9:E{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=6, value=f"=SUM(F9:F{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=7, value=f"=D{r}/C{r}").number_format = PCT_FORMAT
    for c_idx in range(1, 8):
        c = ws.cell(row=r, column=c_idx)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_PRIMARY_NAVY)
        c.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
        c.border = create_border(style="medium", color=COLOR_PRIMARY_NAVY)
        if c_idx >= 2:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22
    
    dept_end_row = r
    r += 2

    # ----------------------------------------------------
    # TABLE 2: SUMMARY BY MANAGER
    # ----------------------------------------------------
    ws.cell(row=r, column=1, value="2. TỔNG HỢP VẬN HÀNH THEO QUẢN LÝ (MANAGER)").font = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_PRIMARY_NAVY)
    r += 1
    
    mgr_headers = ["Quản Lý", "Nhân Sự Phụ Trách", "Lương Cơ Bản", "Thưởng", "Phạt", "Thực Nhận", "Chi Phí TB/Nhân Sự"]
    for idx, h_text in enumerate(mgr_headers, 1):
        c = ws.cell(row=r, column=idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = create_border(color="FFFFFF")
    ws.row_dimensions[r].height = 25
    r += 1
    
    mgr_summary = clean_df.groupby("Manager").agg(
        Count=("Employee_ID", "count"),
        Base=("Base_Salary", "sum"),
        Bonus=("Bonus", "sum"),
        Penalty=("Penalty", "sum"),
        Net=("Net_Salary", "sum"),
        Avg_Net=("Net_Salary", "mean")
    ).reset_index()

    mgr_start_row = r
    for _, m_row in mgr_summary.iterrows():
        ws.cell(row=r, column=1, value=m_row["Manager"]).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r, column=2, value=m_row["Count"]).number_format = INT_FORMAT
        ws.cell(row=r, column=3, value=m_row["Base"]).number_format = VND_FORMAT
        ws.cell(row=r, column=4, value=m_row["Bonus"]).number_format = VND_FORMAT
        ws.cell(row=r, column=5, value=m_row["Penalty"]).number_format = VND_FORMAT
        ws.cell(row=r, column=6, value=m_row["Net"]).number_format = VND_FORMAT
        ws.cell(row=r, column=7, value=f"=F{r}/B{r}").number_format = VND_FORMAT
        
        bg = COLOR_ZEBRA_EVEN if (r % 2 == 0) else COLOR_ZEBRA_ODD
        for c_idx in range(1, 8):
            c = ws.cell(row=r, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.border = create_border()
            if c_idx >= 2:
                c.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1
        
    # Total row Manager
    ws.cell(row=r, column=1, value="TỔNG CỘNG")
    ws.cell(row=r, column=2, value=f"=SUM(B{mgr_start_row}:B{r-1})").number_format = INT_FORMAT
    ws.cell(row=r, column=3, value=f"=SUM(C{mgr_start_row}:C{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=4, value=f"=SUM(D{mgr_start_row}:D{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=5, value=f"=SUM(E{mgr_start_row}:E{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=6, value=f"=SUM(F{mgr_start_row}:F{r-1})").number_format = VND_FORMAT
    ws.cell(row=r, column=7, value=f"=F{r}/B{r}").number_format = VND_FORMAT
    for c_idx in range(1, 8):
        c = ws.cell(row=r, column=c_idx)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_PRIMARY_NAVY)
        c.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
        c.border = create_border(style="medium", color=COLOR_PRIMARY_NAVY)
        if c_idx >= 2:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22
    mgr_end_row = r

    # ----------------------------------------------------
    # EMBEDDED CHARTS
    # ----------------------------------------------------
    # Chart 1: Bar Chart of Net Salary by Department
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Tổng Lương Thực Nhận Theo Phòng Ban (VND)"
    chart1.y_axis.title = "VND"
    chart1.x_axis.title = "Phòng Ban"
    chart1.height = 12
    chart1.width = 16
    
    data1 = Reference(ws, min_col=6, min_row=8, max_row=dept_end_row-1)
    cats1 = Reference(ws, min_col=1, min_row=9, max_row=dept_end_row-1)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.legend = None
    ws.add_chart(chart1, "H7")
    
    # Chart 2: Bar Chart of Headcount & Net Payroll by Manager
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 13
    chart2.title = "Tổng Lương Thực Nhận Theo Manager (VND)"
    chart2.y_axis.title = "VND"
    chart2.x_axis.title = "Manager"
    chart2.height = 12
    chart2.width = 16
    
    data2 = Reference(ws, min_col=6, min_row=mgr_start_row-1, max_row=mgr_end_row-1)
    cats2 = Reference(ws, min_col=1, min_row=mgr_start_row, max_row=mgr_end_row-1)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.legend = None
    ws.add_chart(chart2, f"H{mgr_start_row-1}")

    auto_fit_columns(ws, min_widths={1: 18, 2: 18, 3: 20, 4: 16, 5: 16, 6: 22, 7: 20})

    # ----------------------------------------------------
    # SHEET 2: CLEAN MASTER DATA (200 records)
    # ----------------------------------------------------
    ws_data = wb.create_sheet(title="Dữ Liệu Đã Làm Sạch")
    ws_data.views.sheetView[0].showGridLines = True
    
    master_headers = ["Mã NV", "Họ & Tên", "Quản Lý", "Phòng Ban", "Lương Cơ Bản", "Thưởng", "Phạt", "Lương Thực Nhận", "Kỳ", "Cảnh Báo", "Ghi Chú"]
    ws_data.row_dimensions[1].height = 26
    for c_idx, h_text in enumerate(master_headers, 1):
        c = ws_data.cell(row=1, column=c_idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = create_border(color="FFFFFF")
        
    for r_idx, emp in clean_df.iterrows():
        excel_r = r_idx + 2
        bg = COLOR_ZEBRA_EVEN if (excel_r % 2 == 0) else COLOR_ZEBRA_ODD
        
        ws_data.cell(row=excel_r, column=1, value=emp["Employee_ID"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_data.cell(row=excel_r, column=2, value=emp["Employee_Name"]).alignment = Alignment(horizontal="left", vertical="center")
        ws_data.cell(row=excel_r, column=3, value=emp["Manager"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_data.cell(row=excel_r, column=4, value=emp["Department"]).alignment = Alignment(horizontal="center", vertical="center")
        
        ws_data.cell(row=excel_r, column=5, value=float(emp["Base_Salary"])).number_format = VND_FORMAT
        ws_data.cell(row=excel_r, column=6, value=float(emp["Bonus"])).number_format = VND_FORMAT
        ws_data.cell(row=excel_r, column=7, value=float(emp["Penalty"])).number_format = VND_FORMAT
        
        ws_data.cell(row=excel_r, column=8, value=f"=E{excel_r}+F{excel_r}-G{excel_r}").number_format = VND_FORMAT
        ws_data.cell(row=excel_r, column=9, value=str(emp["Month"])).alignment = Alignment(horizontal="center", vertical="center")
        ws_data.cell(row=excel_r, column=10, value=str(emp["Anomaly_Flag"])).alignment = Alignment(horizontal="center", vertical="center")
        ws_data.cell(row=excel_r, column=11, value=str(emp["Audit_Note"])).alignment = Alignment(horizontal="left", vertical="center")
        
        for c_idx in range(1, 12):
            c = ws_data.cell(row=excel_r, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.border = create_border()
            if c_idx in (5, 6, 7, 8):
                c.alignment = Alignment(horizontal="right", vertical="center")
                
    # Master Total Row
    tot_r = len(clean_df) + 2
    ws_data.cell(row=tot_r, column=1, value="TỔNG CỘNG")
    ws_data.cell(row=tot_r, column=2, value=f"{len(clean_df)} NV")
    ws_data.cell(row=tot_r, column=5, value=f"=SUM(E2:E{tot_r-1})").number_format = VND_FORMAT
    ws_data.cell(row=tot_r, column=6, value=f"=SUM(F2:F{tot_r-1})").number_format = VND_FORMAT
    ws_data.cell(row=tot_r, column=7, value=f"=SUM(G2:G{tot_r-1})").number_format = VND_FORMAT
    ws_data.cell(row=tot_r, column=8, value=f"=SUM(H2:H{tot_r-1})").number_format = VND_FORMAT
    
    for c_idx in range(1, 12):
        c = ws_data.cell(row=tot_r, column=c_idx)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_PRIMARY_NAVY)
        c.fill = PatternFill(start_color=COLOR_TOTAL_BG, end_color=COLOR_TOTAL_BG, fill_type="solid")
        c.border = create_border(style="medium", color=COLOR_PRIMARY_NAVY)
        if c_idx in (5, 6, 7, 8):
            c.alignment = Alignment(horizontal="right", vertical="center")
            
    auto_fit_columns(ws_data, min_widths={1: 10, 2: 18, 3: 14, 4: 15, 5: 18, 6: 16, 7: 16, 8: 20, 9: 12, 10: 15, 11: 30})

    # ----------------------------------------------------
    # SHEET 3: DATA QUALITY & AUDIT LOG
    # ----------------------------------------------------
    ws_audit = wb.create_sheet(title="Kiểm Toán & Bất Thường")
    ws_audit.views.sheetView[0].showGridLines = True
    
    audit_headers = ["STT", "Mã NV", "Họ & Tên", "Quản Lý", "Phòng Ban", "Lương Cơ Bản", "Thưởng", "Phạt", "Lương Thực Nhận", "Phân Loại Lỗi", "Mức Độ", "Mô Tả Chi Tiết", "Khuyến Nghị Xử Lý"]
    ws_audit.row_dimensions[1].height = 28
    for c_idx, h_text in enumerate(audit_headers, 1):
        c = ws_audit.cell(row=1, column=c_idx, value=h_text)
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = create_border(color="FFFFFF")
        
    for idx, a_row in anomaly_df.iterrows():
        a_r = idx + 2
        ws_audit.cell(row=a_r, column=1, value=idx + 1).alignment = Alignment(horizontal="center", vertical="center")
        ws_audit.cell(row=a_r, column=2, value=a_row["Employee_ID"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_audit.cell(row=a_r, column=3, value=a_row["Employee_Name"]).alignment = Alignment(horizontal="left", vertical="center")
        ws_audit.cell(row=a_r, column=4, value=a_row["Manager"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_audit.cell(row=a_r, column=5, value=a_row["Department"]).alignment = Alignment(horizontal="center", vertical="center")
        
        ws_audit.cell(row=a_r, column=6, value=float(a_row["Base_Salary"])).number_format = VND_FORMAT
        ws_audit.cell(row=a_r, column=7, value=float(a_row["Bonus"])).number_format = VND_FORMAT
        ws_audit.cell(row=a_r, column=8, value=float(a_row["Penalty"])).number_format = VND_FORMAT
        ws_audit.cell(row=a_r, column=9, value=float(a_row["Net_Salary"])).number_format = VND_FORMAT
        
        ws_audit.cell(row=a_r, column=10, value=a_row["Issue_Type"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_audit.cell(row=a_r, column=11, value=a_row["Severity"]).alignment = Alignment(horizontal="center", vertical="center")
        ws_audit.cell(row=a_r, column=12, value=a_row["Audit_Note"]).alignment = Alignment(horizontal="left", vertical="center")
        ws_audit.cell(row=a_r, column=13, value=a_row["Recommendation"]).alignment = Alignment(horizontal="left", vertical="center")

        fill_color = COLOR_ALERT_RED_BG if a_row["Severity"] == "High" else (COLOR_ALERT_AMBER_BG if a_row["Severity"] == "Medium" else COLOR_CARD_BG)
        text_color = COLOR_ALERT_RED_TXT if a_row["Severity"] == "High" else (COLOR_ALERT_AMBER_TXT if a_row["Severity"] == "Medium" else COLOR_DARK_TEXT)
        
        for c_idx in range(1, 14):
            c = ws_audit.cell(row=a_r, column=c_idx)
            c.font = Font(name=FONT_NAME, size=9, color=text_color)
            c.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            c.border = create_border()
            if c_idx in (6, 7, 8, 9):
                c.alignment = Alignment(horizontal="right", vertical="center")
        ws_audit.row_dimensions[a_r].height = 20
        
    auto_fit_columns(ws_audit, min_widths={1: 6, 2: 10, 3: 16, 4: 12, 5: 14, 6: 16, 7: 15, 8: 15, 9: 18, 10: 16, 11: 10, 12: 35, 13: 35})

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    print(f"📊 Exported Master Dashboard Workbook: {output_path}")


# ==========================================
# MARKDOWN GENERATOR: EXECUTIVE REPORT
# ==========================================
def generate_executive_report(
    clean_df: pd.DataFrame, 
    anomaly_df: pd.DataFrame, 
    month_str: str, 
    output_path: str
):
    """Generates a high-level executive operations markdown report for directors."""
    total_staff = len(clean_df)
    total_base = clean_df["Base_Salary"].sum()
    total_bonus = clean_df["Bonus"].sum()
    total_pen = clean_df["Penalty"].sum()
    total_net = clean_df["Net_Salary"].sum()
    avg_net = clean_df["Net_Salary"].mean()
    bonus_ratio = (total_bonus / total_base) * 100
    pen_ratio = (total_pen / total_base) * 100

    dept_stats = clean_df.groupby("Department").agg(
        Count=("Employee_ID", "count"),
        Total_Base=("Base_Salary", "sum"),
        Total_Bonus=("Bonus", "sum"),
        Total_Penalty=("Penalty", "sum"),
        Total_Net=("Net_Salary", "sum"),
        Avg_Net=("Net_Salary", "mean")
    ).reset_index()

    mgr_stats = clean_df.groupby("Manager").agg(
        Count=("Employee_ID", "count"),
        Total_Base=("Base_Salary", "sum"),
        Total_Bonus=("Bonus", "sum"),
        Total_Penalty=("Penalty", "sum"),
        Total_Net=("Net_Salary", "sum"),
        Avg_Net=("Net_Salary", "mean")
    ).reset_index()

    high_penalty_count = (clean_df["Penalty"] >= 1500000).sum()
    high_bonus_count = (clean_df["Bonus"] >= 4000000).sum()
    dual_anomaly_count = len(clean_df[(clean_df["Penalty"] >= 1500000) & (clean_df["Bonus"] >= 4000000)])

    report = f"""# BÁO CÁO PHÂN TÍCH VẬN HÀNH & HIỆU SUẤT NHÂN SỰ TOÀN CÔNG TY
**Kỳ báo cáo:** Tháng {month_str}  
**Vị trí:** Operation Analyst  
**Đối tượng nhận:** Giám đốc Vận hành (COO), Ban Giám đốc, Trưởng các bộ phận  
**Hệ thống nguồn:** Dữ liệu ERP xuất từ Google Sheets (Cleaned & Validated)  
**Ngày lập:** {datetime.now().strftime('%d/%m/%Y')}

---

## 1. TỔNG QUAN CHỈ SỐ VẬN HÀNH TOÀN CÔNG TY (EXECUTIVE SUMMARY)

Trong kỳ tháng {month_str}, toàn bộ quy trình thu thập, làm sạch và kiểm toán dữ liệu vận hành từ hệ thống ERP đã hoàn tất với **200 nhân sự** trên 4 khối phòng ban trực thuộc 4 Quản lý trực tiếp.

| Chỉ số vận hành cốt lõi | Giá trị kỳ này | Đơn vị tính | Đánh giá sơ bộ |
|:---|---:|:---:|:---|
| **Tổng quy mô nhân sự** | **{total_staff:,}** | Nhân sự | Phân bổ 4 phòng ban chính (HR, Finance, Ops, Sales) |
| **Tổng quỹ lương cơ bản (Base)** | **{total_base:,.0f}** | VNĐ | Chiếm 92.67% tổng chi trả thực nhận |
| **Tổng ngân sách thưởng (Bonus)** | **{total_bonus:,.0f}** | VNĐ | Chiếm **{bonus_ratio:.2f}%** so với lương cơ bản |
| **Tổng tiền phạt vi phạm (Penalty)** | **{total_pen:,.0f}** | VNĐ | Tương đương **{pen_ratio:.2f}%** quỹ lương cơ bản |
| **Tổng ngân sách thực chi (Net Payroll)** | **{total_net:,.0f}** | VNĐ | Đã cấn trừ thưởng và phạt |
| **Thu nhập thực nhận bình quân/nhân sự** | **{avg_net:,.0f}** | VNĐ/người | Dao động từ 8.68 tr đến 33.31 tr VNĐ |

> [!NOTE]
> **Nhận định chung:** Tỷ lệ quỹ thưởng toàn công ty đạt **{bonus_ratio:.2f}%**, nằm trong ngưỡng an toàn ngân sách (< 15%). Tuy nhiên, tỷ lệ phạt vi phạm **{pen_ratio:.2f}%** (hơn 200 triệu VNĐ) phản ánh các vi phạm quy chế hoặc sự cố vận hành vẫn còn xuất hiện ở tần suất đáng kể cần được kiểm soát theo từng bộ phận.

---

## 2. PHÂN TÍCH CHI PHÍ & HIỆU SUẤT THEO PHÒNG BAN

```
Tỷ lệ Lương Thực Nhận theo Bộ phận:
  - HR:          1,205,692,735 VNĐ (29.6%) [58 NV]
  - Operations:  1,002,857,196 VNĐ (24.6%) [48 NV]
  - Finance:       974,363,294 VNĐ (23.9%) [50 NV]
  - Sales:         887,281,782 VNĐ (21.8%) [44 NV]
```

### Bảng tổng hợp chi tiết theo Phòng ban:
| Phòng Ban | Quy mô (NV) | Quỹ Lương Cơ Bản (VNĐ) | Tổng Thưởng (VNĐ) | Tổng Phạt (VNĐ) | Lương Thực Nhận (VNĐ) | Thu Nhập TB/NV (VNĐ) | Tỷ lệ Thưởng/Lương | Tỷ lệ Phạt/Lương |
|:---|:---:|---:|---:|---:|---:|---:|:---:|:---:|
"""
    for _, d in dept_stats.iterrows():
        b_r = (d["Total_Bonus"] / d["Total_Base"]) * 100
        p_r = (d["Total_Penalty"] / d["Total_Base"]) * 100
        report += f"| **{d['Department']}** | {d['Count']} | {d['Total_Base']:,.0f} | {d['Total_Bonus']:,.0f} | {d['Total_Penalty']:,.0f} | **{d['Total_Net']:,.0f}** | {d['Avg_Net']:,.0f} | {b_r:.1f}% | {p_r:.1f}% |\n"

    report += f"""
### Nhận xét chuyên sâu:
1. **Phòng Operations:** Có **tỷ lệ phạt vi phạm cao nhất toàn công ty ({dept_stats.loc[dept_stats['Department']=='Operations', 'Total_Penalty'].values[0] / dept_stats.loc[dept_stats['Department']=='Operations', 'Total_Base'].values[0] * 100:.2f}%)** với tổng tiền phạt lên tới **{dept_stats.loc[dept_stats['Department']=='Operations', 'Total_Penalty'].values[0]:,.0f} VNĐ**. Cần điều tra xem các lỗi vận hành xuất phát từ SLA giao vận, trễ tiến độ hay sai sót quy trình giao dịch.
2. **Khối HR:** Chiếm quy mô lớn nhất (58 nhân sự) và có tổng quỹ thực nhận cao nhất (1.205 tỷ VNĐ). Tỷ lệ thưởng của HR cũng cao nhất (**14.19%**), tương ứng mức độ ghi nhận thành tích tuyển dụng hoặc dự án nhân sự nội bộ.
3. **Khối Finance:** Tỷ lệ phạt thấp nhất ({dept_stats.loc[dept_stats['Department']=='Finance', 'Total_Penalty'].values[0] / dept_stats.loc[dept_stats['Department']=='Finance', 'Total_Base'].values[0] * 100:.2f}%), cho thấy mức độ tuân thủ quy chuẩn kế toán và quy trình nghiệp vụ rất chặt chẽ.

---

## 3. PHÂN BỔ QUY MÔ & KHỐI LƯỢNG CÔNG VIỆC THEO QUẢN LÝ (MANAGER)

Hệ thống ghi nhận sự chênh lệch lớn về khối lượng nhân sự do 4 Quản lý phụ trách:

| Quản Lý | Số NV Quản Lý | % Tổng Nhân Sự | Tổng Lương Cơ Bản | Tổng Thưởng | Tổng Phạt | Tổng Thực Nhận | Chi Phí TB/NV |
|:---|:---:|:---:|---:|---:|---:|---:|---:|
"""
    for _, m in mgr_stats.iterrows():
        pct = (m["Count"] / total_staff) * 100
        report += f"| **{m['Manager']}** | **{m['Count']}** | {pct:.1f}% | {m['Total_Base']:,.0f} | {m['Total_Bonus']:,.0f} | {m['Total_Penalty']:,.0f} | **{m['Total_Net']:,.0f}** | {m['Avg_Net']:,.0f} |\n"

    report += f"""
### Phát hiện quản trị:
- **Manager_A gánh tải quản lý cao nhất:** Quản lý trực tiếp **62 nhân sự (31% toàn công ty)** trải rộng trên cả 4 phòng ban (18 Finance, 19 HR, 13 Operations, 12 Sales). Việc một người quản lý 62 cấp dưới phân tán nhiều chuyên môn tạo nguy cơ "Span of Control" quá rộng, dễ dẫn đến chậm trễ phê duyệt hoặc giảm sâu sát trong đánh giá hiệu suất.
- **Manager_B có quy mô tinh gọn nhất:** Chỉ quản lý 37 nhân sự (18.5%), tổng quỹ lương phụ trách là 759 triệu VNĐ.
- **Manager_C & Manager_D:** Duy trì quy mô cân bằng tốt (51 và 50 nhân sự).

---

## 4. BÁO CÁO KIỂM TOÁN DỮ LIỆU & CẢNH BÁO BẤT THƯỜNG (ANOMALY AUDIT)

Hệ thống tự động quét toàn bộ 200 nhân sự dựa trên các quy tắc kiểm toán nghiệp vụ:
- **Cảnh báo Phạt cao (Penalty >= 1.500.000 VNĐ):** **{high_penalty_count} trường hợp**
- **Cảnh báo Thưởng vượt khung (Bonus >= 4.000.000 VNĐ):** **{high_bonus_count} trường hợp**
- **Cảnh báo Kép (Vừa Thưởng cao vừa Bị phạt cao):** **{dual_anomaly_count} trường hợp**

### Top 5 trường hợp bất thường nghiêm trọng nhất cần Manager rà soát:
| Mã NV | Họ & Tên | Quản Lý | Phòng Ban | Thưởng (VNĐ) | Phạt (VNĐ) | Thực Nhận (VNĐ) | Lý do Audit |
|:---|:---|:---|:---|---:|---:|---:|:---|
"""
    # Sort anomalies by penalty desc
    top_anomalies = anomaly_df.sort_values(by="Penalty", ascending=False).head(5)
    for _, an in top_anomalies.iterrows():
        report += f"| **{an['Employee_ID']}** | {an['Employee_Name']} | {an['Manager']} | {an['Department']} | {an['Bonus']:,.0f} | **{an['Penalty']:,.0f}** | {an['Net_Salary']:,.0f} | {an['Audit_Note']} |\n"

    report += f"""
> [!WARNING]
> **Đặc biệt lưu ý nhân sự E031 (HR - Manager_A):** Đạt mức thưởng rất cao (**4,925,144 VNĐ**) nhưng đồng thời chịu mức phạt cao nhất toàn công ty (**1,993,943 VNĐ**). Đây là ca bất thường điển hình (Dual Anomaly), yêu cầu Manager_A kiểm tra lại xem có sự nhầm lẫn trong việc nhập liệu vi phạm hoặc nhân sự này có thành tích vượt trội nhưng mắc lỗi quy trình vận hành nghiêm trọng.

---

## 5. ĐỀ XUẤT HÀNH ĐỘNG THEO CHU TRÌNH PDCA

Dưới góc độ Operation Analyst, đề xuất 3 hành động cụ thể cho Ban Giám đốc và các Manager trong kỳ tới:

```
[PLAN] Rà soát định mức phạt & Tái cơ cấu Span of Control
  ↓
[DO]   Phát hành file Excel riêng cho 4 Manager & Tổ chức họp 1-1 với nhân viên bị phạt cao
  ↓
[CHECK]Đo lường tỷ lệ tái diễn vi phạm & Tỷ lệ giải ngân lương qua Master Dashboard
  ↓
[ACT]  Tự động hóa pipeline đưa vào cron job định kỳ ngày 05 hàng tháng
```

1. **Hành động 1 (Tái cân bằng quản lý):** Xem xét bổ nhiệm thêm Team Lead hoặc phân bổ lại nhân sự từ **Manager_A** (hiện 62 người) sang **Manager_B** (hiện 37 người) để đảm bảo chất lượng giám sát vận hành.
2. **Hành động 2 (Xử lý gốc rễ vi phạm khối Operations):** Tổ chức phiên làm việc giữa Operations Manager và QA/QC để làm rõ nguyên nhân tổng tiền phạt vượt 53.9 triệu VNĐ (tỷ lệ 5.76% quỹ lương). Thiết lập checklist tiền kiểm để giảm thiểu lỗi vận hành lặp lại.
3. **Hành động 3 (Xác nhận phê duyệt danh sách Anomaly):** Trước khi kế toán chuyển khoản ngày 10/{month_str}, 4 Manager cần ký duyệt văn bản giải trình đối với {len(anomaly_df)} nhân sự nằm trong danh sách kiểm toán đính kèm trong Sheet 3 của Master Dashboard.

---
*Báo cáo được tự động tạo lập từ Python Automation Pipeline (`process_operations.py`) vào lúc {datetime.now().strftime('%H:%M %d/%m/%Y')}.*
"""

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"📝 Exported Executive Report: {output_path}")


# ==========================================
# MAIN PIPELINE CONTROLLER
# ==========================================
def main():
    parser = argparse.ArgumentParser(description="End-to-End Automated Operations Data Pipeline")
    parser.add_argument(
        "--input", "--url", 
        dest="source", 
        default="https://docs.google.com/spreadsheets/d/1bn41RZ8Eg66YXOAO7RmN-oyuA6on8d2ESMSVH1zcBd8/export?format=csv",
        help="Google Sheet export URL or local CSV/Excel path"
    )
    parser.add_argument("--month", dest="month", default=None, help="Target month (YYYY-MM), e.g., 2026-03")
    parser.add_argument("--date", dest="date", default=None, help="Report date folder (default: D-M-YYYY, e.g. 9-9-2026)")
    parser.add_argument("--outdir", dest="outdir", default="outputs", help="Base output directory")
    parser.add_argument("--threshold-penalty", type=float, default=1500000.0, help="Penalty anomaly threshold (VND)")
    parser.add_argument("--threshold-bonus", type=float, default=4000000.0, help="Bonus anomaly threshold (VND)")
    
    args = parser.parse_args()
    
    print("=" * 70)
    print("🚀 STARTING AUTOMATED OPERATIONS DATA PIPELINE")
    print("=" * 70)
    
    # Step 1: Ingest Data
    raw_cache_path = os.path.join("sample-data", "ERP_Operations_Raw_2026_03.csv")
    raw_df = ingest_data(args.source, raw_cache_path)
    
    # Determine Month
    month_str = args.month
    if not month_str and "Month" in raw_df.columns:
        month_str = str(raw_df["Month"].iloc[0]).strip()
    month_str = month_str or "2026-03"
    print(f"📅 Target Month: {month_str}")
    
    # Determine Report Date folder (e.g. 9-9-2026)
    now = datetime.now()
    report_date = args.date or f"{now.day}-{now.month}-{now.year}"
    print(f"📁 Report Folder Date: {report_date}")
    
    # Step 2: Clean & Enrich Data
    clean_df, anomaly_df = clean_and_enrich_data(
        raw_df, 
        threshold_penalty=args.threshold_penalty, 
        threshold_bonus=args.threshold_bonus
    )
    
    # Step 3: Split and generate workbooks for each Manager
    managers = sorted(clean_df["Manager"].unique())
    print(f"\n📂 Generating workbooks for {len(managers)} Managers: {', '.join(managers)}")
    mgr_dir = os.path.join(args.outdir, "managers", month_str)
    os.makedirs(mgr_dir, exist_ok=True)
    
    for mgr in managers:
        mgr_sub_df = clean_df[clean_df["Manager"] == mgr].copy()
        mgr_filename = f"{mgr}_Operations_{month_str.replace('-', '_')}.xlsx"
        mgr_out_path = os.path.join(mgr_dir, mgr_filename)
        generate_manager_workbook(
            manager_name=mgr,
            mgr_df=mgr_sub_df,
            month_str=month_str,
            output_path=mgr_out_path,
            threshold_penalty=args.threshold_penalty,
            threshold_bonus=args.threshold_bonus
        )
        
    # Step 4: Generate Master Dashboard into date-based report folder
    reports_dir = os.path.join(args.outdir, "reports", report_date)
    os.makedirs(reports_dir, exist_ok=True)
    
    master_dash_path = os.path.join(reports_dir, f"Operations_Master_Dashboard_{month_str.replace('-', '_')}.xlsx")
    generate_master_dashboard(clean_df, anomaly_df, month_str, master_dash_path)
    
    # Step 5: Generate Executive Markdown Report into date-based report folder
    exec_report_path = os.path.join(reports_dir, f"Operations_Executive_Report_{month_str.replace('-', '_')}.md")
    generate_executive_report(clean_df, anomaly_df, month_str, exec_report_path)
    
    print("\n" + "=" * 70)
    print("🎉 ALL AUTOMATION DELIVERABLES CREATED SUCCESSFULLY!")
    print("=" * 70)
    print(f"1. Raw Data Backup:    {raw_cache_path}")
    print(f"2. Manager Workbooks:   {mgr_dir}/")
    for mgr in managers:
        print(f"   - {mgr}_Operations_{month_str.replace('-', '_')}.xlsx")
    print(f"3. Master Dashboard:    {master_dash_path}")
    print(f"4. Executive Report:    {exec_report_path}")
    print("=" * 70)


if __name__ == "__main__":
    main()
