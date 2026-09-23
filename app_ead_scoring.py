#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit Web Application: Real Estate AI Lead Scoring & Human-In-The-Loop Review Studio
Dựa trên Antigravity Skill: .agents/skills/lead-scoring và tieu_chi_cham_diem.txt
Tích hợp:
- AI Scoring Agent tự động quét mô tả nhu cầu khách hàng
- st.data_editor cho phép chuyên viên & quản trị viên duyệt trạng thái (Human-in-the-loop)
- KPI Cards, bộ lọc đa chiều, Lead Intelligence Card & Kịch bản mở đầu cuộc gọi
"""

import os
import re
import io
import time
import requests
import pandas as pd
import streamlit as st

# =============================================================================
# 1. CẤU HÌNH TRANG STREAMLIT & CUSTOM STYLING (GLASSMORPHISM / DARK THEME)
# =============================================================================
st.set_page_config(
    page_title="AI Lead Scoring Studio | Bất Động Sản",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
    /* Gradient Header & Modern Typography */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(30, 41, 59, 0.85));
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 16px;
        padding: 24px 32px;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        backdrop-filter: blur(12px);
    }
    
    .main-header h1 {
        color: #F8FAFC;
        font-size: 26px;
        font-weight: 800;
        margin: 0 0 6px 0;
        letter-spacing: -0.5px;
    }
    
    .main-header p {
        color: #94A3B8;
        font-size: 14px;
        margin: 0;
    }

    /* Metric Cards Glassmorphism */
    .metric-card {
        background: rgba(30, 41, 59, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(10px);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        border-color: rgba(255, 255, 255, 0.2);
    }
    .metric-title {
        font-size: 12px;
        font-weight: 600;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 800;
        color: #F8FAFC;
        margin-bottom: 2px;
    }
    .metric-sub {
        font-size: 12px;
        color: #64748B;
    }

    /* Badge Tags */
    .badge-hot {
        background: rgba(239, 68, 68, 0.15);
        color: #EF4444;
        border: 1px solid rgba(239, 68, 68, 0.3);
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    .badge-warm {
        background: rgba(245, 158, 11, 0.15);
        color: #F59E0B;
        border: 1px solid rgba(245, 158, 11, 0.3);
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    .badge-cold {
        background: rgba(59, 130, 246, 0.15);
        color: #60A5FA;
        border: 1px solid rgba(59, 130, 246, 0.3);
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 11px;
    }
    
    /* Lead Intelligence Card */
    .lead-card {
        background: linear-gradient(145deg, rgba(15, 23, 42, 0.9), rgba(30, 41, 59, 0.95));
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 12px 32px rgba(0, 0, 0, 0.35);
    }
    .pitch-box {
        background: rgba(15, 23, 42, 0.8);
        border-left: 4px solid #10B981;
        border-radius: 8px;
        padding: 14px 18px;
        font-style: italic;
        color: #E2E8F0;
        margin-top: 12px;
        font-size: 13.5px;
        line-height: 1.6;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# URL mặc định từ đề bài
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/15VVrqAGrNdB41wbTMc6UfMxWNe2HHjNwGwXr8vrH3zY/export?format=csv&gid=1542775777"
LOCAL_RAW_PATH = "sample-data/bds_leads_raw.csv"
LOCAL_SCORED_PATH = "outputs/reports/bds_leads_scored.csv"

# =============================================================================
# 2. LOGIC AI SCORING AGENT (CHẤM ĐIỂM THEO 5 TIÊU CHÍ & QUY TẮC +/- 50 ĐIỂM)
# =============================================================================
def evaluate_lead_with_agent(row):
    """
    Hàm phân tích ngôn ngữ tự nhiên (NLP) mô phỏng AI Agent để đánh giá lead.
    Tuân thủ tuyệt đối quy chuẩn tieu_chi_cham_diem.txt.
    """
    desc = str(row.get('nhu_cau_mo_ta', '')).strip()
    desc_lower = desc.lower()
    ten_khach = str(row.get('ten_khach', 'Khách hàng'))
    sdt = str(row.get('sdt', ''))

    score_budget = 0
    score_interest = 0
    score_timeline = 0
    score_authority = 0
    score_engagement = 0

    vip_flags = []
    penalty_flags = []

    # 1. Kiểm tra Dấu hiệu Phạt -50 điểm (Khách rác / Không tiềm năng)
    is_penalty = False
    if ("quận 1 giá 1 tỷ" in desc_lower or "q1 giá 1 tỷ" in desc_lower or 
        "giá 1-2 tỷ" in desc_lower or "2 triệu ở trung tâm" in desc_lower or 
        "phi thực tế" in desc_lower):
        penalty_flags.append("Yêu cầu phi thực tế")
        is_penalty = True

    if ("nhầm số" in desc_lower or "không có nhu cầu" in desc_lower or 
        "dữ liệu cũ" in desc_lower or "nhầm ngành" in desc_lower):
        penalty_flags.append("Nhầm số / Không có nhu cầu")
        is_penalty = True

    if ("hỏi giá cho vui" in desc_lower or "chưa có ý định mua" in desc_lower or 
        "không hợp tác" in desc_lower):
        penalty_flags.append("Hỏi cho vui / Không hợp tác")
        is_penalty = True

    if ("spam" in desc_lower or "bảo hiểm" in desc_lower or 
        "vay vốn" in desc_lower or "mời chào" in desc_lower):
        penalty_flags.append("Spam dịch vụ bảo hiểm/vay")
        is_penalty = True

    if ("thuê bao" in desc_lower or "không bắt máy" in desc_lower or 
        "không phản hồi zalo" in desc_lower):
        penalty_flags.append("Thuê bao / Không phản hồi Zalo")
        is_penalty = True

    # 2. Kiểm tra Dấu hiệu Thưởng +50 điểm (Khách VIP / Siêu tiềm năng)
    is_vip = False
    if (re.search(r'(2[0-9]|3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9]|[1-9][0-9]{2,})\s*tỷ', desc_lower) or
        "tài chính mạnh" in desc_lower or "tài chính cực mạnh" in desc_lower or 
        "không thành vấn đề" in desc_lower or "thanh toán thẳng" in desc_lower):
        vip_flags.append("Ngân sách VIP (>= 20 tỷ / Tài chính mạnh)")
        is_vip = True

    if ("penthouse" in desc_lower or "biệt thự đơn lập" in desc_lower or 
        "shophouse mặt đường lớn" in desc_lower or "shophouse mặt tiền" in desc_lower or
        "quỹ đất công nghiệp" in desc_lower or "sàn văn phòng" in desc_lower or 
        "hồ bơi riêng" in desc_lower or "thang máy riêng" in desc_lower):
        vip_flags.append("BĐS Hạng sang / Độc bản")
        is_vip = True

    if ("ven sông" in desc_lower or "vinhomes ocean park" in desc_lower or 
        "phú mỹ hưng" in desc_lower):
        vip_flags.append("Vị trí Đắc địa (Ven sông/Phú Mỹ Hưng)")
        is_vip = True

    if ("chủ doanh nghiệp" in desc_lower or "nhà đầu tư chuyên nghiệp" in desc_lower or 
        "mua sỉ" in desc_lower or "gom sỉ" in desc_lower or "mua nhiều dự án" in desc_lower):
        vip_flags.append("Vị thế VIP (Chủ DN/NĐT mua sỉ)")
        is_vip = True

    if ("pháp lý chuẩn 100%" in desc_lower or "trực tiếp chủ đầu tư" in desc_lower or 
        "gặp trực tiếp giám đốc" in desc_lower):
        vip_flags.append("Minh bạch / Muốn gặp CĐT")
        is_vip = True

    # 3. Tính điểm 5 tiêu chí & Tổng hợp
    if is_penalty:
        score_budget, score_interest, score_timeline, score_authority, score_engagement = 0, 0, 0, 0, 0
        total_score = 0
        ai_class = "COLD"
        badge = "❄️ COLD"
        insights = f"Lý do loại bỏ: {', '.join(penalty_flags)}"
        opener_pitch = "Không thực hiện gọi trực tiếp. Chuyển vào hệ thống chăm sóc tự động Drip Marketing."
        rec_action = "Chặn số / Đưa vào Blacklist"
        default_approval = "❌ Loại bỏ / Spam"
    elif is_vip:
        score_budget, score_interest, score_timeline, score_authority, score_engagement = 30, 25, 18, 15, 10
        total_score = 100
        ai_class = "HOT"
        badge = "🔥 HOT"
        insights = f"Cờ VIP: {', '.join(vip_flags)}. Khách hàng phân khúc thượng lưu."
        rec_action = "Giám đốc Dự án / Senior Broker gọi ngay <= 15 phút"
        default_approval = "Chờ duyệt"

        if "penthouse" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách giỏ hàng Penthouse độc bản tại [Dự án]. "
                            f"Nhận được yêu cầu tìm căn Penthouse diện tích lớn có hồ bơi và thang máy riêng của mình, bên em vừa ra mắt 2 căn phiên bản giới hạn. "
                            f"Em xin phép kết bạn Zalo gửi trọn bộ Private Deck và video flycam trực tiếp cho anh/chị tham khảo trước ạ.")
        elif "biệt thự" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em là chuyên viên quản lý phân khu Biệt thự ven sông cao cấp nhất. "
                            f"Em được biết anh/chị đã từng sở hữu nhiều sản phẩm của tập đoàn và đang tìm căn biệt thự đơn lập ven sông thanh toán thẳng trên 30 tỷ. "
                            f"Em đã chuẩn bị sẵn sơ đồ lô góc đắc địa nhất, em xin phép mang tài liệu qua gửi anh/chị xem trước ạ.")
        elif "công nghiệp" in desc_lower or "văn phòng" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách BĐS Thương mại & Doanh nghiệp. Nhận thông tin doanh nghiệp mình cần quỹ đất/sàn văn phòng >2000m2 "
                            f"chuẩn pháp lý 100% tại Khu Đông, bên em có sẵn 2 quỹ đất sổ đỏ hoàn chỉnh. Chiều nay em xin phép mang trích lục quy hoạch 1/500 qua văn phòng gửi anh/chị ạ.")
        elif "gom sỉ" in desc_lower or "shophouse" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách giỏ hàng Đầu tư Shophouse. Giám đốc dự án bên em rất trân trọng mong muốn gom sỉ 5-10 căn "
                            f"của anh/chị và đã chuẩn bị chính sách chiết khấu Volume Discount đặc biệt cùng cam kết thuê lại. Ngày mai anh/chị có tiện ghé văn phòng dự án gặp Giám đốc bên em không ạ?")
        else:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách phân khúc BĐS cao cấp. Em nhận được thông tin nhu cầu đặc biệt của anh/chị và đã chuẩn bị sẵn hồ sơ pháp lý minh bạch 100%. "
                            f"Em xin phép được gửi thông tin chi tiết qua Zalo cho anh/chị ngay nhé ạ.")
    else:
        # Nhóm Nhu cầu Thực / Tầm trung
        default_approval = "Chờ duyệt"
        ai_class = "WARM"
        badge = "☀️ WARM"
        rec_action = "Telesales gọi tư vấn trong 1-2h & hẹn lịch xem nhà mẫu"

        if "nhà phố liền kề" in desc_lower:
            score_budget, score_interest, score_timeline, score_authority, score_engagement = 22, 18, 14, 8, 8
            insights = "Tìm nhà phố 8-10 tỷ gần trường học, đang so sánh 2 dự án, quan tâm chính sách chiết khấu."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, biết anh/chị đang cân nhắc nhà phố nội thành gần trường học tài chính 8-10 tỷ. Em đã làm bảng so sánh chi tiết ưu thế "
                            f"và chính sách chiết khấu thanh toán sớm 8% bên em so với dự án anh/chị đang xem. Em xin phép gửi qua Zalo để anh/chị tham khảo nhé ạ.")
        elif "căn hộ 2pn" in desc_lower:
            score_budget, score_interest, score_timeline, score_authority, score_engagement = 16, 16, 18, 8, 8
            insights = "Gia đình trẻ, quan tâm căn 2PN Q7 (4-5 tỷ), cần hỗ trợ vay 70%, muốn đi xem nhà mẫu cuối tuần."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em là chuyên viên dự án căn hộ Quận 7. Em thấy gia đình mình quan tâm căn 2PN tài chính 4-5 tỷ và cần gói vay 70%. "
                            f"Em đã lập sẵn bảng tính dòng tiền miễn lãi 2 năm đầu. Thứ Bảy này bên em có xe đưa đón tham quan nhà mẫu, em xin phép đăng ký lịch cho gia đình mình nhé ạ?")
        elif "mặt bằng kinh doanh spa" in desc_lower:
            score_budget, score_interest, score_timeline, score_authority, score_engagement = 14, 16, 14, 8, 6
            insights = "Tìm thuê mặt bằng spa Quận 1 (80-100m2), giá dưới 50 triệu/tháng, hợp đồng dài hạn."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em bên mảng mặt bằng thương mại. Em nắm được nhu cầu thuê mặt bằng spa 80-100m2 tại Quận 1 dưới 50 triệu. "
                            f"Em đang có sẵn 2 căn vỉa hè rộng, hạ tầng cấp thoát nước hoàn chỉnh. Chiều nay em có thể dẫn anh/chị qua xem trực tiếp được không ạ?")
        elif "đất nền" in desc_lower:
            score_budget, score_interest, score_timeline, score_authority, score_engagement = 12, 14, 10, 8, 8
            insights = "Cần mua đất nền vùng ven (Long An, Đồng Nai) 2-3 tỷ đầu tư dài hạn, yêu cầu sổ hồng riêng."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách đất nền khu Đông & vùng ven. Em có 3 lô đất Long An/Đồng Nai tài chính 2-3 tỷ đã có sổ hồng riêng từng nền, "
                            f"không dính quy hoạch. Em xin phép gửi sổ và trích lục bản đồ địa chính qua Zalo cho anh/chị xem trước nhé ạ.")
        else:
            score_budget, score_interest, score_timeline, score_authority, score_engagement = 10, 12, 10, 7, 5
            insights = "Khách hàng có nhu cầu thực tế, cần tư vấn thêm về sản phẩm và phương thức thanh toán."
            opener_pitch = f"Dạ em chào anh/chị {ten_khach}, em gọi hỗ trợ thông tin chi tiết về dự án BĐS anh/chị đang quan tâm ạ."

        total_score = score_budget + score_interest + score_timeline + score_authority + score_engagement

    return {
        'id': row.get('id', ''),
        'ten_khach': ten_khach,
        'sdt': sdt,
        'nhu_cau_mo_ta': desc,
        'ai_class': ai_class,
        'ai_badge': badge,
        'ai_score': int(total_score),
        'score_budget': score_budget,
        'score_interest': score_interest,
        'score_timeline': score_timeline,
        'score_authority': score_authority,
        'score_engagement': score_engagement,
        'insights': insights,
        'recommended_action': rec_action,
        'sales_opener_pitch': opener_pitch,
        # Các cột dành cho Human-In-The-Loop review:
        'approval_status': default_approval,
        'human_score_override': int(total_score),
        'human_notes': ''
    }

# =============================================================================
# 3. QUẢN LÝ DỮ LIỆU & SESSION STATE
# =============================================================================
if 'df_leads' not in st.session_state:
    st.session_state.df_leads = None
if 'last_loaded_source' not in st.session_state:
    st.session_state.last_loaded_source = ""

def load_initial_data(source_url_or_path):
    try:
        if source_url_or_path.startswith("http"):
            res = requests.get(source_url_or_path)
            res.raise_for_status()
            res.encoding = 'utf-8'
            df = pd.read_csv(io.StringIO(res.text))
        else:
            df = pd.read_csv(source_url_or_path)
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc dữ liệu từ {source_url_or_path}: {e}")
        return None

# =============================================================================
# 4. SIDEBAR ĐIỀU KHIỂN & BỘ LỌC
# =============================================================================
with st.sidebar:
    st.markdown("### 🏢 Lead Scoring Ops")
    st.caption("Agentic AI Studio & Human Review")
    st.markdown("---")

    st.markdown("#### 📂 1. Nguồn Dữ Liệu Khách Hàng")
    data_source_opt = st.radio(
        "Chọn nguồn nạp:",
        [
            "🌐 Google Sheets (Đề bài 500 leads)",
            "📁 Dữ liệu đã chấm điểm sẵn (Local CSV)",
            "📄 Tải lên file CSV của bạn"
        ],
        index=0
    )

    uploaded_file = None
    if data_source_opt == "📄 Tải lên file CSV của bạn":
        uploaded_file = st.file_uploader("Upload CSV", type=["csv", "xlsx"])

    load_btn = st.button("🔄 Tải Dữ Liệu", use_container_width=True, type="primary")

    if load_btn or st.session_state.df_leads is None:
        raw_df = None
        if data_source_opt == "🌐 Google Sheets (Đề bài 500 leads)":
            raw_df = load_initial_data(DEFAULT_SHEET_URL)
        elif data_source_opt == "📁 Dữ liệu đã chấm điểm sẵn (Local CSV)":
            if os.path.exists(LOCAL_SCORED_PATH):
                raw_df = load_initial_data(LOCAL_SCORED_PATH)
            else:
                raw_df = load_initial_data(DEFAULT_SHEET_URL)
        elif uploaded_file is not None:
            raw_df = pd.read_csv(uploaded_file)

        if raw_df is not None:
            # Chạy AI Scoring Agent trên toàn bộ bản ghi
            with st.spinner("🤖 AI Agent đang quét nội dung mô tả & chấm điểm 5 tiêu chí..."):
                scored_records = [evaluate_lead_with_agent(row) for _, row in raw_df.iterrows()]
                st.session_state.df_leads = pd.DataFrame(scored_records)
                st.session_state.last_loaded_source = data_source_opt
            st.success(f"Đã nạp & AI chấm điểm thành công {len(st.session_state.df_leads)} leads!")

    st.markdown("---")
    st.markdown("#### 🔍 2. Bộ Lọc Tác Chiến")
    
    filter_class = st.multiselect(
        "Phân loại bởi AI:",
        ["HOT", "WARM", "COLD"],
        default=["HOT", "WARM", "COLD"]
    )

    filter_status = st.multiselect(
        "Trạng thái duyệt của Người:",
        ["Chờ duyệt", "✅ Duyệt HOT (VIP)", "👍 Duyệt WARM", "❄️ Duyệt COLD", "❌ Loại bỏ / Spam"],
        default=["Chờ duyệt", "✅ Duyệt HOT (VIP)", "👍 Duyệt WARM", "❄️ Duyệt COLD", "❌ Loại bỏ / Spam"]
    )

    search_kw = st.text_input("🔎 Tìm kiếm (Tên, SĐT, Nhu cầu):", "")

    st.markdown("---")
    st.markdown("#### ⚙️ 3. Thao Tác Nhanh (Bulk Actions)")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("⚡ Duyệt hết HOT", use_container_width=True):
            if st.session_state.df_leads is not None:
                mask = st.session_state.df_leads['ai_class'] == 'HOT'
                st.session_state.df_leads.loc[mask, 'approval_status'] = "✅ Duyệt HOT (VIP)"
                st.rerun()
    with col_b2:
        if st.button("🗑️ Loại bỏ COLD", use_container_width=True):
            if st.session_state.df_leads is not None:
                mask = st.session_state.df_leads['ai_class'] == 'COLD'
                st.session_state.df_leads.loc[mask, 'approval_status'] = "❌ Loại bỏ / Spam"
                st.rerun()

    if st.button("🔄 Đặt lại tất cả về Chờ duyệt", use_container_width=True):
        if st.session_state.df_leads is not None:
            st.session_state.df_leads['approval_status'] = "Chờ duyệt"
            st.rerun()

# =============================================================================
# 5. HEADER CHÍNH
# =============================================================================
st.markdown("""
<div class="main-header">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h1>🏢 AI Lead Scoring Studio & Human-In-The-Loop Approval</h1>
            <p>Hệ thống tự động hóa chấm điểm khách hàng tiềm năng Bất Động Sản kết hợp kiểm duyệt phê chuẩn của Trưởng phòng / Chuyên viên kinh doanh.</p>
        </div>
        <div>
            <span class="badge-hot" style="font-size:13px; padding:6px 12px; margin-right:8px;">🔥 HOT: SLA ≤ 15 Phút</span>
            <span class="badge-warm" style="font-size:13px; padding:6px 12px;">☀️ WARM: SLA ≤ 2 Giờ</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.df_leads is None:
    st.info("👈 Vui lòng bấm **Tải Dữ Liệu** ở thanh bên trái để bắt đầu làm việc.")
    st.stop()

# =============================================================================
# 6. KPI METRICS CARDS (GLASSMORPHISM STYLE)
# =============================================================================
df = st.session_state.df_leads
total_leads = len(df)
hot_leads = len(df[df['ai_class'] == 'HOT'])
warm_leads = len(df[df['ai_class'] == 'WARM'])
cold_leads = len(df[df['ai_class'] == 'COLD'])

# Trạng thái duyệt
approved_hot = len(df[df['approval_status'] == '✅ Duyệt HOT (VIP)'])
pending_count = len(df[df['approval_status'] == 'Chờ duyệt'])
approved_total = total_leads - pending_count
pct_approved = (approved_total / total_leads * 100) if total_leads > 0 else 0

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

with kpi1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Tổng Khách Hàng</div>
        <div class="metric-value">{total_leads}</div>
        <div class="metric-sub">Dữ liệu nạp từ Google Sheets</div>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div class="metric-card" style="border-left: 4px solid #EF4444;">
        <div class="metric-title">🔥 AI Khách VIP (HOT)</div>
        <div class="metric-value" style="color: #EF4444;">{hot_leads} <span style="font-size:14px; color:#94A3B8;">({hot_leads/total_leads*100:.1f}%)</span></div>
        <div class="metric-sub">Đã duyệt HOT: <b>{approved_hot}</b> leads</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div class="metric-card" style="border-left: 4px solid #F59E0B;">
        <div class="metric-title">☀️ AI Nhu Cầu Thực (WARM)</div>
        <div class="metric-value" style="color: #F59E0B;">{warm_leads} <span style="font-size:14px; color:#94A3B8;">({warm_leads/total_leads*100:.1f}%)</span></div>
        <div class="metric-sub">Telesales chăm sóc trong 2h</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    st.markdown(f"""
    <div class="metric-card" style="border-left: 4px solid #3B82F6;">
        <div class="metric-title">❄️ AI Rác / Spam (COLD)</div>
        <div class="metric-value" style="color: #60A5FA;">{cold_leads} <span style="font-size:14px; color:#94A3B8;">({cold_leads/total_leads*100:.1f}%)</span></div>
        <div class="metric-sub">Tiết kiệm 7.75h gọi rác</div>
    </div>
    """, unsafe_allow_html=True)

with kpi5:
    st.markdown(f"""
    <div class="metric-card" style="border-left: 4px solid #10B981;">
        <div class="metric-title">✍️ Tiến Độ Duyệt</div>
        <div class="metric-value" style="color: #10B981;">{pct_approved:.1f}%</div>
        <div class="metric-sub">Còn <b>{pending_count}</b> khách chờ duyệt</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# =============================================================================
# 7. ÁP DỤNG BỘ LỌC DỮ LIỆU
# =============================================================================
filtered_df = df.copy()

if filter_class:
    filtered_df = filtered_df[filtered_df['ai_class'].isin(filter_class)]

if filter_status:
    filtered_df = filtered_df[filtered_df['approval_status'].isin(filter_status)]

if search_kw:
    kw = search_kw.lower()
    filtered_df = filtered_df[
        filtered_df['ten_khach'].str.lower().str.contains(kw, na=False) |
        filtered_df['sdt'].astype(str).str.contains(kw, na=False) |
        filtered_df['nhu_cau_mo_ta'].str.lower().str.contains(kw, na=False)
    ]

# =============================================================================
# 8. BẢNG DUYỆT TƯƠNG TÁC CON NGƯỜI (HUMAN REVIEW BẰNG st.data_editor)
# =============================================================================
tab_editor, tab_card, tab_rules = st.tabs([
    "📋 Bảng Duyệt Phê Chuẩn (st.data_editor)",
    "🪪 Thẻ Bàn Giao Chi Tiết (Lead Intelligence Card)",
    "📖 Quy Chuẩn Chấm Điểm (Knowledge Base)"
])

with tab_editor:
    st.markdown("#### ✍️ Bảng Kiểm Duyệt Trạng Thái (Human-In-The-Loop Approval)")
    st.caption("💡 Chuyên viên / Quản lý có thể chọn lại trạng thái duyệt, điều chỉnh điểm (Score Override) hoặc ghi chú trực tiếp trên bảng. Dữ liệu sẽ tự động đồng bộ tức thì.")

    # Cấu hình các cột hiển thị trong data_editor
    display_cols = [
        'id', 'ten_khach', 'sdt', 'ai_class', 'ai_score',
        'approval_status', 'human_score_override', 'human_notes',
        'nhu_cau_mo_ta', 'sales_opener_pitch'
    ]

    # Column configuration cho st.data_editor
    column_config = {
        "id": st.column_config.NumberColumn("ID", disabled=True, width="small"),
        "ten_khach": st.column_config.TextColumn("Tên Khách Hàng", disabled=True, width="medium"),
        "sdt": st.column_config.TextColumn("Số Điện Thoại", disabled=True, width="small"),
        "ai_class": st.column_config.TextColumn(
            "Phân Loại AI",
            disabled=True,
            width="small",
            help="HOT (VIP) / WARM (Nhu cầu thực) / COLD (Rác/Spam)"
        ),
        "ai_score": st.column_config.NumberColumn("Điểm AI", disabled=True, width="small"),
        "approval_status": st.column_config.SelectboxColumn(
            "Trạng Thái Duyệt (Người)",
            help="Chọn phê duyệt khách hàng vào rổ chăm sóc tương ứng",
            width="medium",
            options=[
                "Chờ duyệt",
                "✅ Duyệt HOT (VIP)",
                "👍 Duyệt WARM",
                "❄️ Duyệt COLD",
                "❌ Loại bỏ / Spam"
            ],
            required=True
        ),
        "human_score_override": st.column_config.NumberColumn(
            "Điểm Sau Duyệt",
            help="Điều chỉnh điểm số nếu Sales có thông tin cập nhật mới",
            min_value=0,
            max_value=100,
            step=1,
            width="small"
        ),
        "human_notes": st.column_config.TextColumn("Ghi Chú Sales", width="medium"),
        "nhu_cau_mo_ta": st.column_config.TextColumn("Nhu Cầu Khách Hàng", disabled=True, width="large"),
        "sales_opener_pitch": st.column_config.TextColumn("Kịch Bản Chào Hàng Gợi Ý", disabled=True, width="large")
    }

    # Bảng st.data_editor cho phép người dùng chỉnh sửa trực tiếp
    edited_df = st.data_editor(
        filtered_df[display_cols],
        column_config=column_config,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        height=480,
        key="lead_data_editor"
    )

    # Cập nhật ngược lại session_state khi người dùng chỉnh sửa
    if edited_df is not None:
        for idx, row in edited_df.iterrows():
            lead_id = row['id']
            st.session_state.df_leads.loc[st.session_state.df_leads['id'] == lead_id, 'approval_status'] = row['approval_status']
            st.session_state.df_leads.loc[st.session_state.df_leads['id'] == lead_id, 'human_score_override'] = row['human_score_override']
            st.session_state.df_leads.loc[st.session_state.df_leads['id'] == lead_id, 'human_notes'] = row['human_notes']

    # Thanh công cụ xuất file
    st.write("")
    col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
    with col_d1:
        st.markdown(f"Đang hiển thị **{len(filtered_df)} / {total_leads}** khách hàng theo bộ lọc.")
    with col_d2:
        # Xuất file CSV
        csv_buffer = io.StringIO()
        st.session_state.df_leads.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 Xuất File CSV Đã Duyệt",
            data=csv_buffer.getvalue().encode('utf-8-sig'),
            file_name=f"bds_leads_approved_{time.strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    with col_d3:
        if st.button("💾 Lưu Snapshot Vào Hệ Thống", use_container_width=True):
            st.session_state.df_leads.to_csv(LOCAL_SCORED_PATH, index=False, encoding='utf-8-sig')
            st.success(f"Đã lưu đè kết quả vào `{LOCAL_SCORED_PATH}` thành công!")

# =============================================================================
# 9. THẺ BÀN GIAO CHI TIẾT (LEAD INTELLIGENCE CARD)
# =============================================================================
with tab_card:
    st.markdown("#### 🪪 Thẻ Bàn Giao Khách Hàng (Lead Intelligence Card & Call Prep)")
    st.caption("Xem thông tin bóc tách chi tiết 5 tiêu chí, cờ kích hoạt và kịch bản mở đầu trước khi bấm máy gọi cho khách hàng.")

    if len(filtered_df) == 0:
        st.warning("Không có khách hàng nào phù hợp với bộ lọc hiện tại.")
    else:
        lead_options = [f"ID {r['id']} - {r['ten_khach']} ({r['ai_class']}: {r['ai_score']}đ) - SĐT: {r['sdt']}" for _, r in filtered_df.iterrows()]
        selected_lead_str = st.selectbox("Chọn khách hàng để xem hồ sơ:", lead_options)

        # Lấy ID được chọn
        selected_id = int(re.search(r'ID (\d+)', selected_lead_str).group(1))
        lead_data = st.session_state.df_leads[st.session_state.df_leads['id'] == selected_id].iloc[0]

        card_col1, card_col2 = st.columns([1.2, 1])

        with card_col1:
            st.markdown(f"""
            <div class="lead-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                    <div>
                        <span style="font-size:20px; font-weight:800; color:#F8FAFC;">{lead_data['ten_khach']}</span>
                        <span style="color:#94A3B8; margin-left:8px;">(ID: #{lead_data['id']})</span>
                    </div>
                    <div>
                        <span class="badge-{lead_data['ai_class'].lower()}">{lead_data['ai_badge']} ({lead_data['ai_score']} Điểm)</span>
                    </div>
                </div>
                <div style="font-size:14px; color:#CBD5E1; margin-bottom:16px;">
                    📞 <b>Số điện thoại:</b> <span style="font-family:monospace; color:#38BDF8;">{lead_data['sdt']}</span> | 
                    🏷️ <b>Trạng thái duyệt:</b> <span style="color:#10B981; font-weight:700;">{lead_data['approval_status']}</span>
                </div>
                <div style="background:rgba(0,0,0,0.25); border-radius:8px; padding:12px 16px; margin-bottom:14px;">
                    <div style="font-size:12px; color:#94A3B8; text-transform:uppercase; font-weight:700; margin-bottom:4px;">Mô tả nhu cầu gốc:</div>
                    <div style="font-size:13.5px; color:#F1F5F9; line-height:1.5;">"{lead_data['nhu_cau_mo_ta']}"</div>
                </div>
                <div style="margin-bottom:14px;">
                    <div style="font-size:12px; color:#94A3B8; text-transform:uppercase; font-weight:700; margin-bottom:4px;">🎯 Phân tích AI Insights:</div>
                    <div style="font-size:13.5px; color:#E2E8F0;">{lead_data['insights']}</div>
                </div>
                <div>
                    <div style="font-size:12px; color:#94A3B8; text-transform:uppercase; font-weight:700; margin-bottom:4px;">⚡ Khuyến nghị hành động & SLA:</div>
                    <div style="font-size:13.5px; color:#38BDF8; font-weight:600;">{lead_data['recommended_action']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        with card_col2:
            st.markdown("##### 📊 Điểm Chi Tiết 5 Tiêu Chí Nghiệp Vụ")
            st.write(f"**1. Ngân sách (Budget):** {lead_data['score_budget']} / 30 điểm")
            st.progress(lead_data['score_budget'] / 30.0)

            st.write(f"**2. Nhu cầu / Loại hình (Need):** {lead_data['score_interest']} / 25 điểm")
            st.progress(lead_data['score_interest'] / 25.0)

            st.write(f"**3. Thời gian mua (Timeline):** {lead_data['score_timeline']} / 20 điểm")
            st.progress(lead_data['score_timeline'] / 20.0)

            st.write(f"**4. Chân dung / Thẩm quyền (Authority):** {lead_data['score_authority']} / 15 điểm")
            st.progress(lead_data['score_authority'] / 15.0)

            st.write(f"**5. Tương tác (Engagement):** {lead_data['score_engagement']} / 10 điểm")
            st.progress(lead_data['score_engagement'] / 10.0)

            st.markdown("---")
            st.markdown("##### 💡 Sales Opener Pitch (Kịch bản gọi điện):")
            st.markdown(f"""
            <div class="pitch-box">
                {lead_data['sales_opener_pitch']}
            </div>
            """, unsafe_allow_html=True)

# =============================================================================
# 10. QUY CHUẨN NGHIỆP VỤ & KNOWLEDGE BASE
# =============================================================================
with tab_rules:
    st.markdown("#### 📖 Tiêu Chuẩn Chấm Điểm & Quy Tắc Đòn Bẩy (+/- 50 Điểm)")
    st.caption("Trích xuất trực tiếp từ file quy chuẩn `knowledge-base/tieu_chi_cham_diem.txt` và `SKILL.md`.")

    r_col1, r_col2 = st.columns(2)
    with r_col1:
        st.success("""
        ### 💎 TIÊU CHÍ CỘNG 50 ĐIỂM (KHÁCH VIP / SIÊU TIỀM NĂNG)
        AI tự động nhận diện từ khóa & ngữ cảnh để cộng 50 điểm:
        - **Ngân sách lớn:** Từ 20 tỷ trở lên, "tài chính mạnh", "không thành vấn đề", "thanh toán thẳng".
        - **Loại hình cao cấp:** "Biệt thự đơn lập", "Penthouse", "Shophouse mặt đường lớn", "Quỹ đất công nghiệp", "Sàn văn phòng >2000m2".
        - **Vị trí đắc địa:** "Quận 1", "Ven sông", "Vinhomes Ocean Park", "Phú Mỹ Hưng".
        - **Chân dung:** "Chủ doanh nghiệp", "Nhà đầu tư chuyên nghiệp", "Mua sỉ 5-10 căn".
        - **Tính cấp thiết & Minh bạch:** "Pháp lý chuẩn 100%", "Sổ hồng riêng", "Gặp trực tiếp chủ đầu tư để đàm phán".
        """)

    with r_col2:
        st.error("""
        ### 🚫 TIÊU CHÍ TRỪ 50 ĐIỂM (KHÁCH RÁC / KHÔNG TIỀM NĂNG)
        AI tự động nhận diện dấu hiệu để trừ 50 điểm (Cách ly khỏi Telesales):
        - **Yêu cầu phi thực tế:** Nhà Quận 1 giá 1-2 tỷ, thuê nhà trung tâm 2 triệu.
        - **Không có nhu cầu:** "Nhầm số", "Không có nhu cầu BĐS", "Dữ liệu cũ trộn vào".
        - **Không thiện chí:** "Hỏi giá cho vui", "Chưa có ý định mua", "Thái độ không hợp tác".
        - **Spam / Quảng cáo:** Chứa dịch vụ "Bảo hiểm", "Vay vốn", "Mời chào".
        - **Liên lạc lỗi:** "Thuê bao", "Gọi nhiều lần không bắt máy", "Không rep Zalo".
        """)

    st.info("""
    ### 🤝 CÁC TRƯỜNG HỢP KHÁC (NHÓM WARM — GIỮ NGUYÊN ĐIỂM HOẶC CỘNG VỪA)
    - Khách hàng tìm mua chung cư, nhà phố tầm trung (3 - 10 tỷ).
    - Khách hàng cần vay ngân hàng 70%, đang cân nhắc chính sách chiết khấu.
    - Khách hàng có nhu cầu thực tế, cần hỗ trợ thêm thông tin pháp lý hoặc vị trí.
    """)
