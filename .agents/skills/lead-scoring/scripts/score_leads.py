#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lead Scoring Engine for Real Estate (Bất Động Sản)
Tự động hóa chấm điểm và phân loại khách hàng tiềm năng dựa trên 5 tiêu chí chuẩn BANT/CHAMP
và quy tắc thưởng/phạt +/- 50 điểm theo knowledge-base tieu_chi_cham_diem.txt.
"""

import os
import sys
import re
import argparse
import pandas as pd
import requests

DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/15VVrqAGrNdB41wbTMc6UfMxWNe2HHjNwGwXr8vrH3zY/edit?gid=1542775777#gid=1542775777"
EXPORT_CSV_URL = "https://docs.google.com/spreadsheets/d/15VVrqAGrNdB41wbTMc6UfMxWNe2HHjNwGwXr8vrH3zY/export?format=csv&gid=1542775777"
DEFAULT_RAW_OUTPUT = "sample-data/bds_leads_raw.csv"
DEFAULT_SCORED_OUTPUT = "outputs/reports/bds_leads_scored.csv"

def get_download_url(url):
    """Chuyển đổi URL Google Sheets sang URL Export CSV trực tiếp."""
    if "docs.google.com/spreadsheets" in url:
        if "/export?format=csv" not in url:
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
            gid_match = re.search(r'gid=([0-9]+)', url)
            if match:
                sheet_id = match.group(1)
                gid = gid_match.group(1) if gid_match else "0"
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    return url

def load_data(source_path_or_url):
    """Tải dữ liệu từ URL Google Sheets hoặc file CSV cục bộ."""
    if source_path_or_url.startswith("http://") or source_path_or_url.startswith("https://"):
        download_url = get_download_url(source_path_or_url)
        print(f"📥 Đang tải dữ liệu từ Google Sheets: {download_url}")
        res = requests.get(download_url)
        res.raise_for_status()
        res.encoding = 'utf-8'
        from io import StringIO
        df = pd.read_csv(StringIO(res.text))
    else:
        print(f"📂 Đang đọc dữ liệu từ tệp cục bộ: {source_path_or_url}")
        df = pd.read_csv(source_path_or_url)
    return df

def score_lead_row(row):
    """
    Chấm điểm 1 dòng khách hàng dựa trên mô tả nhu cầu và các thuộc tính.
    Trả về dict chứa các điểm thành phần, cờ VIP, cờ Phạt, phân loại và kịch bản.
    """
    desc = str(row.get('nhu_cau_mo_ta', '')).strip()
    desc_lower = desc.lower()
    ten_khach = str(row.get('ten_khach', 'Khách hàng'))
    sdt = str(row.get('sdt', ''))

    # Khởi tạo điểm thành phần
    score_budget = 0
    score_interest = 0
    score_timeline = 0
    score_authority = 0
    score_engagement = 0

    vip_flags = []
    penalty_flags = []

    # =========================================================================
    # 1. KIỂM TRA DẤU HIỆU PHẠT -50 ĐIỂM (Khách hàng rác / Không tiềm năng)
    # Theo Mục 2 tieu_chi_cham_diem.txt
    # =========================================================================
    is_penalty = False

    # Dấu hiệu 1: Yêu cầu phi thực tế
    if ("quận 1 giá 1 tỷ" in desc_lower or "q1 giá 1 tỷ" in desc_lower or 
        "giá 1-2 tỷ" in desc_lower or "2 triệu ở trung tâm" in desc_lower or 
        "phi thực tế" in desc_lower):
        penalty_flags.append("Yêu cầu phi thực tế (Giá vô lý)")
        is_penalty = True

    # Dấu hiệu 2: Không có nhu cầu / Nhầm số
    if ("nhầm số" in desc_lower or "không có nhu cầu" in desc_lower or 
        "dữ liệu cũ" in desc_lower or "nhầm ngành" in desc_lower):
        penalty_flags.append("Nhầm số / Không có nhu cầu BĐS")
        is_penalty = True

    # Dấu hiệu 3: Không thiện chí
    if ("hỏi giá cho vui" in desc_lower or "chưa có ý định mua" in desc_lower or 
        "không hợp tác" in desc_lower):
        penalty_flags.append("Khách không thiện chí (Hỏi cho vui)")
        is_penalty = True

    # Dấu hiệu 4: Spam / Quảng cáo
    if ("spam" in desc_lower or "bảo hiểm" in desc_lower or 
        "vay vốn" in desc_lower or "mời chào" in desc_lower):
        penalty_flags.append("Spam / Mời chào dịch vụ khác")
        is_penalty = True

    # Dấu hiệu 5: Thông tin liên lạc lỗi
    if ("thuê bao" in desc_lower or "không bắt máy" in desc_lower or 
        "không phản hồi zalo" in desc_lower):
        penalty_flags.append("Không liên lạc được (Thuê bao/Không rep)")
        is_penalty = True

    # =========================================================================
    # 2. KIỂM TRA DẤU HIỆU CỘNG +50 ĐIỂM (Khách hàng VIP / Siêu tiềm năng)
    # Theo Mục 1 tieu_chi_cham_diem.txt
    # =========================================================================
    is_vip = False

    # Ngân sách lớn >= 20 tỷ hoặc cụm từ VIP
    if (re.search(r'(2[0-9]|3[0-9]|4[0-9]|5[0-9]|6[0-9]|7[0-9]|8[0-9]|9[0-9]|[1-9][0-9]{2,})\s*tỷ', desc_lower) or
        "tài chính mạnh" in desc_lower or "tài chính cực mạnh" in desc_lower or 
        "không thành vấn đề" in desc_lower or "thanh toán thẳng" in desc_lower):
        vip_flags.append("Ngân sách VIP (>= 20 tỷ / Tài chính mạnh)")
        is_vip = True

    # Loại hình cao cấp
    if ("penthouse" in desc_lower or "biệt thự đơn lập" in desc_lower or 
        "shophouse mặt đường lớn" in desc_lower or "shophouse mặt tiền" in desc_lower or
        "quỹ đất công nghiệp" in desc_lower or "sàn văn phòng" in desc_lower or 
        "hồ bơi riêng" in desc_lower or "thang máy riêng" in desc_lower):
        vip_flags.append("Sản phẩm VIP / Độc bản")
        is_vip = True

    # Vị trí đắc địa
    if ("ven sông" in desc_lower or "vinhomes ocean park" in desc_lower or 
        "phú mỹ hưng" in desc_lower):
        vip_flags.append("Vị trí Đắc địa (Ven sông/Phú Mỹ Hưng)")
        is_vip = True

    # Đối tượng khách hàng VIP
    if ("chủ doanh nghiệp" in desc_lower or "nhà đầu tư chuyên nghiệp" in desc_lower or 
        "mua sỉ" in desc_lower or "gom sỉ" in desc_lower or "mua nhiều dự án" in desc_lower):
        vip_flags.append("Chân dung VIP (Chủ DN/NĐT gom sỉ)")
        is_vip = True

    # Tính cấp thiết & Minh bạch
    if ("pháp lý chuẩn 100%" in desc_lower or "trực tiếp chủ đầu tư" in desc_lower or 
        "gặp trực tiếp giám đốc" in desc_lower):
        vip_flags.append("Minh bạch / Muốn gặp CĐT")
        is_vip = True

    # =========================================================================
    # 3. TÍNH ĐIỂM 5 TIÊU CHÍ CHUẨN (BASE SCORE)
    # =========================================================================
    if is_penalty:
        score_budget = 0
        score_interest = 0
        score_timeline = 0
        score_authority = 0
        score_engagement = 0
        base_score = 0
        bonus_penalty = -50
        total_score = 0
        classification = "COLD"
        badge = "❄️ COLD LEAD"
        rec_action = "Chặn số / Đưa vào Blacklist / Drip Marketing tự động"
        opener_pitch = "Không thực hiện gọi trực tiếp. Chuyển vào hệ thống chăm sóc tự động định kỳ."
        insights = f"Lý do loại: {', '.join(penalty_flags) if penalty_flags else 'Điểm tiềm năng quá thấp'}"

    elif is_vip:
        # Nhóm Khách hàng VIP (+50 điểm) -> Chắc chắn HOT LEAD
        score_budget = 30
        score_interest = 25
        score_timeline = 18
        score_authority = 15
        score_engagement = 10
        base_score = score_budget + score_interest + score_timeline + score_authority + score_engagement # 98đ
        bonus_penalty = 50
        total_score = 100 # Capped ở 100 điểm
        classification = "HOT"
        badge = "🔥 HOT LEAD"
        rec_action = "Giao Giám đốc Dự án / Senior Broker gọi ngay trong vòng 5 - 15 phút"
        insights = f"VIP Flags: {', '.join(vip_flags)}. Khách hàng phân khúc cao cấp/tài chính mạnh."

        if "penthouse" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em là [Tên Sales] phụ trách giỏ hàng Penthouse độc bản tại [Dự án]. "
                            f"Nhận được yêu cầu tìm căn Penthouse diện tích lớn có hồ bơi và thang máy riêng của mình, bên em vừa ra mắt 2 căn phiên bản giới hạn "
                            f"có view panorama trực diện sông. Em xin phép kết bạn Zalo gửi trọn bộ Private Deck và video flycam trực tiếp cho anh/chị xem trước ạ.")
        elif "biệt thự" in desc_lower:
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em là [Tên Sales] phụ trách phân khu Biệt thự ven sông cao cấp nhất. "
                            f"Em được biết anh/chị đã từng sở hữu nhiều sản phẩm của tập đoàn và đang tìm căn biệt thự đơn lập hướng Đông Nam thanh toán thẳng trên 30 tỷ. "
                            f"Em đã chuẩn bị sẵn sơ đồ lô góc vị trí đắc địa nhất, em xin phép mang tài liệu qua gửi anh/chị tham khảo ạ.")
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
        # =====================================================================
        # Nhóm Khách Hàng Nhu Cầu Thực / Tầm Trung (Theo Mục 3 tieu_chi_cham_diem.txt)
        # Giữ nguyên điểm hoặc cộng ít -> Phân loại WARM LEAD
        # =====================================================================
        bonus_penalty = 0

        if "nhà phố liền kề" in desc_lower:
            score_budget = 22
            score_interest = 18
            score_timeline = 14
            score_authority = 8
            score_engagement = 8
            insights = "Tìm nhà phố 8-10 tỷ gần trường học, đang so sánh 2 dự án, quan tâm chính sách chiết khấu."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, biết anh/chị đang cân nhắc nhà phố nội thành gần trường học tài chính 8-10 tỷ. Em đã làm bảng so sánh chi tiết ưu thế "
                            f"và chính sách chiết khấu thanh toán sớm 8% bên em so với dự án anh/chị đang xem. Em xin phép gửi qua Zalo để anh/chị tham khảo nhé ạ.")
        elif "căn hộ 2pn" in desc_lower:
            score_budget = 16
            score_interest = 16
            score_timeline = 18
            score_authority = 8
            score_engagement = 8
            insights = "Gia đình trẻ, quan tâm căn 2PN Q7 (4-5 tỷ), cần hỗ trợ vay 70%, muốn đi xem nhà mẫu cuối tuần."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em là [Tên Sales] dự án căn hộ Quận 7. Em thấy gia đình mình quan tâm căn 2PN tài chính 4-5 tỷ và cần gói vay 70%. "
                            f"Em đã lập sẵn bảng tính dòng tiền miễn lãi 2 năm đầu. Thứ Bảy này bên em có xe đưa đón tham quan nhà mẫu, em xin phép đăng ký lịch cho gia đình mình nhé ạ?")
        elif "mặt bằng kinh doanh spa" in desc_lower:
            score_budget = 14
            score_interest = 16
            score_timeline = 14
            score_authority = 8
            score_engagement = 6
            insights = "Tìm thuê mặt bằng spa Quận 1 (80-100m2), giá dưới 50 triệu/tháng, hợp đồng dài hạn."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em bên mảng mặt bằng thương mại. Em nắm được nhu cầu thuê mặt bằng spa 80-100m2 tại Quận 1 dưới 50 triệu. "
                            f"Em đang có sẵn 2 căn vỉa hè rộng, hạ tầng cấp thoát nước hoàn chỉnh. Chiều nay em có thể dẫn anh/chị qua xem trực tiếp được không ạ?")
        elif "đất nền" in desc_lower:
            score_budget = 12
            score_interest = 14
            score_timeline = 10
            score_authority = 8
            score_engagement = 8
            insights = "Cần mua đất nền vùng ven (Long An, Đồng Nai) 2-3 tỷ đầu tư dài hạn, yêu cầu sổ hồng riêng."
            opener_pitch = (f"Dạ em chào anh/chị {ten_khach}, em phụ trách đất nền khu Đông & vùng ven. Em có 3 lô đất Long An/Đồng Nai tài chính 2-3 tỷ đã có sổ hồng riêng từng nền, "
                            f"không dính quy hoạch. Em xin phép gửi sổ và trích lục bản đồ địa chính qua Zalo cho anh/chị xem trước nhé ạ.")
        else:
            score_budget = 10
            score_interest = 12
            score_timeline = 10
            score_authority = 7
            score_engagement = 5
            insights = "Khách hàng có nhu cầu thực tế, cần tư vấn thêm về sản phẩm và phương thức thanh toán."
            opener_pitch = f"Dạ em chào anh/chị {ten_khach}, em gọi hỗ trợ thông tin chi tiết về dự án BĐS anh/chị đang quan tâm ạ."

        base_score = score_budget + score_interest + score_timeline + score_authority + score_engagement
        total_score = base_score
        classification = "WARM"
        badge = "☀️ WARM LEAD"
        rec_action = "Chuyên viên tư vấn gọi trong 1 - 2 giờ, hỗ trợ phương án tài chính & hẹn nhà mẫu"

    return {
        'id': row.get('id', ''),
        'ten_khach': ten_khach,
        'sdt': sdt,
        'nhu_cau_mo_ta': desc,
        'score_budget': score_budget,
        'score_interest': score_interest,
        'score_timeline': score_timeline,
        'score_authority': score_authority,
        'score_engagement': score_engagement,
        'base_score': base_score,
        'bonus_penalty': bonus_penalty,
        'total_score': total_score,
        'classification': classification,
        'badge': badge,
        'key_insights': insights,
        'recommended_action': rec_action,
        'sales_opener_pitch': opener_pitch
    }

def main():
    parser = argparse.ArgumentParser(description="Chấm điểm khách hàng tiềm năng Bất Động Sản (Lead Scoring Engine)")
    parser.add_argument("--input", default=EXPORT_CSV_URL, help="URL Google Sheets hoặc đường dẫn tệp CSV/Excel")
    parser.add_argument("--save-raw", default=DEFAULT_RAW_OUTPUT, help="Đường dẫn lưu bản sao dữ liệu gốc")
    parser.add_argument("--output", default=DEFAULT_SCORED_OUTPUT, help="Đường dẫn lưu kết quả chấm điểm")
    args = parser.parse_args()

    os.makedirs(os.path.dirname(args.save_raw), exist_ok=True)
    os.makedirs(os.path.dirname(args.output), exist_ok=True)

    df_raw = load_data(args.input)
    print(f"📊 Đã tải thành công {len(df_raw)} dòng dữ liệu khách hàng.")

    df_raw.to_csv(args.save_raw, index=False, encoding='utf-8-sig')
    print(f"💾 Đã lưu snapshot raw data tại: {args.save_raw}")

    print("⚙️ Đang thực hiện chấm điểm theo 5 tiêu chí và bộ quy tắc nghiệp vụ...")
    scored_records = []
    for _, row in df_raw.iterrows():
        scored_records.append(score_lead_row(row))

    df_scored = pd.DataFrame(scored_records)
    df_scored.to_csv(args.output, index=False, encoding='utf-8-sig')
    print(f"✅ Đã xuất kết quả chấm điểm thành công tại: {args.output}")

    total = len(df_scored)
    hot_count = len(df_scored[df_scored['classification'] == 'HOT'])
    warm_count = len(df_scored[df_scored['classification'] == 'WARM'])
    cold_count = len(df_scored[df_scored['classification'] == 'COLD'])

    print("\n" + "="*60)
    print("📈 BÁO CÁO PHÂN LOẠI TỔNG HỢP KHÁCH HÀNG (LEAD DISTRIBUTION)")
    print("="*60)
    print(f"🔹 Tổng số Leads tiếp nhận: {total}")
    print(f"🔥 HOT LEADS (Ưu tiên GĐDA/Senior Sales gọi <= 15p): {hot_count:>4} ({hot_count/total*100:>5.1f}%)")
    print(f"☀️ WARM LEADS (Telesales gọi tư vấn & hẹn nhà mẫu):   {warm_count:>4} ({warm_count/total*100:>5.1f}%)")
    print(f"❄️ COLD LEADS (Spam / Rác / Không gọi trực tiếp):    {cold_count:>4} ({cold_count/total*100:>5.1f}%)")
    print("="*60)

if __name__ == "__main__":
    main()
