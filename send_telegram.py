#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==============================================================================
WORKFLOW: TỰ ĐỘNG HÓA LẤY TIN AI & GỬI TELEGRAM (MÔ HÌNH OIPO)
Role: Automation Engineer
Framework: OIPO (Objective - Input - Process - Output)
==============================================================================

[O] OBJECTIVE:
- Tự động hóa việc thu thập 1 tin tức trí tuệ nhân tạo (AI) mới nhất từ Google News RSS,
  dịch tiêu đề sang tiếng Việt và gửi thông báo tức thời vào Telegram của người dùng.

[I] INPUT:
- Nguồn tin tức: Google News RSS Feed (chủ đề Artificial Intelligence, OpenAI, Google AI).
- Cấu hình bảo mật: BOT_TOKEN, CHAT_ID được tải an toàn từ file môi trường `.env`.

[P] PROCESS:
- Bước 1: Fetch Google News RSS feed bằng thư viện requests và phân tích XML (ElementTree).
- Bước 2: Trích xuất tiêu đề (title) và đường dẫn bài viết (link) của tin mới nhất.
- Bước 3: Dịch tiêu đề sang tiếng Việt bằng API dịch thuật trực tiếp qua requests (không dùng googletrans).
- Bước 4: Định dạng bản tin theo mẫu chuẩn yêu cầu:
          🧠 Tin AI hôm nay - dd/mm/yyyy
          • [nội dung dịch]
          Nguồn:
          [link]
          #AI #TinCongNghe
- Bước 5: Gửi tin nhắn qua Telegram Bot API (endpoint sendMessage) bằng requests.

[O] OUTPUT:
- Tin nhắn gửi thành công đến người nhận trên ứng dụng Telegram.
- Log chi tiết từng bước vận hành trên terminal.
==============================================================================
"""

import os
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
import requests
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# 1. LOAD CONFIGURATION FROM .ENV
# -----------------------------------------------------------------------------
# Tải các biến môi trường từ tệp .env để đảm bảo an toàn bí mật API Token
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

RSS_URL = (
    "https://news.google.com/rss/search?"
    "q=Artificial+Intelligence+OR+OpenAI+OR+Google+AI&hl=en-US&gl=US&ceid=US:en"
)


# -----------------------------------------------------------------------------
# 2. CORE FUNCTIONS
# -----------------------------------------------------------------------------
def translate_to_vi(text: str) -> str:
    """
    Dịch văn bản tiếng Anh sang tiếng Việt bằng requests (không sử dụng thư viện googletrans).
    Áp dụng cơ chế Multi-Endpoint Fallback để đảm bảo tính ổn định và xử lý lỗi:
      1. Endpoint Chrome Dictionary Client của Google Translate.
      2. Endpoint MyMemory API miễn phí nếu endpoint 1 gặp sự cố.
      3. Trả về văn bản gốc nếu tất cả dịch vụ đều không phản hồi.
    """
    if not text or not text.strip():
        return text

    # Thử Phương án 1: Google Translate Client Endpoint qua requests
    try:
        url_gt = "https://translate.googleapis.com/translate_a/single"
        params_gt = {
            "client": "dict-chrome-ex",
            "sl": "en",
            "tl": "vi",
            "dt": "t",
            "q": text.strip(),
        }
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = requests.get(url_gt, params=params_gt, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            if data and isinstance(data, list) and len(data) > 0 and data[0]:
                translated_parts = [part[0] for part in data[0] if part and part[0]]
                translated_text = "".join(translated_parts).strip()
                if translated_text:
                    return translated_text
    except Exception as e_gt:
        print(f"⚠️ [Translate Warning] Google Translate Endpoint gặp lỗi: {e_gt}")

    # Thử Phương án 2 (Fallback): MyMemory Translation API qua requests
    try:
        url_mm = "https://api.mymemory.translated.net/get"
        params_mm = {"q": text.strip(), "langpair": "en|vi"}
        headers = {"User-Agent": "Mozilla/5.0"}
        resp_mm = requests.get(url_mm, params=params_mm, headers=headers, timeout=10)
        if resp_mm.status_code == 200:
            data_mm = resp_mm.json()
            translated_text = data_mm.get("responseData", {}).get("translatedText")
            if translated_text and translated_text.strip():
                return translated_text.strip()
    except Exception as e_mm:
        print(f"⚠️ [Translate Warning] MyMemory Endpoint gặp lỗi: {e_mm}")

    # Xử lý an toàn: Nếu cả 2 đều lỗi, giữ nguyên văn bản gốc
    print("ℹ️ Giữ nguyên văn bản gốc do không thể kết nối cổng dịch thuật.")
    return text


def get_ai_news() -> dict:
    """
    Lấy 1 tin tức AI mới nhất từ Google News RSS feed bằng requests.
    Bóc tách cấu trúc XML để lấy tiêu đề (title) và liên kết gốc (link).
    
    Returns:
        dict: Chứa {"title": str, "link": str} hoặc None nếu thất bại.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
    }

    try:
        response = requests.get(RSS_URL, headers=headers, timeout=12)
        response.raise_for_status()

        # Phân tích XML từ dữ liệu phản hồi
        root = ET.fromstring(response.content)
        item = root.find(".//item")

        if item is None:
            print("❌ Không tìm thấy bài viết nào trong RSS feed.")
            return None

        title_elem = item.find("title")
        link_elem = item.find("link")

        title = title_elem.text.strip() if title_elem is not None and title_elem.text else "Tin tức AI mới nhất"
        link = link_elem.text.strip() if link_elem is not None and link_elem.text else ""

        return {"title": title, "link": link}

    except requests.exceptions.RequestException as req_err:
        print(f"❌ [Network Error] Lỗi kết nối khi lấy RSS từ Google News: {req_err}")
        return None
    except ET.ParseError as xml_err:
        print(f"❌ [XML Error] Không thể giải mã dữ liệu RSS: {xml_err}")
        return None
    except Exception as general_err:
        print(f"❌ [Error] Lỗi không xác định khi lấy tin AI: {general_err}")
        return None


def send_telegram(message: str) -> bool:
    """
    Gửi tin nhắn định dạng văn bản đến Telegram bằng Telegram Bot API qua requests.
    
    Args:
        message (str): Nội dung tin nhắn cần gửi.
        
    Returns:
        bool: True nếu gửi thành công, False nếu thất bại.
    """
    if not BOT_TOKEN or not CHAT_ID:
        print("❌ [Config Error] Chưa cấu hình BOT_TOKEN hoặc CHAT_ID trong file .env!")
        return False

    api_url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": message,
        "disable_web_page_preview": False,
    }

    try:
        response = requests.post(api_url, json=payload, timeout=15)
        res_data = response.json()

        if response.status_code == 200 and res_data.get("ok"):
            print("🚀 Gửi tin nhắn đến Telegram thành công!")
            return True
        else:
            error_desc = res_data.get("description", "Không rõ nguyên nhân")
            error_code = res_data.get("error_code", response.status_code)
            print(f"❌ [Telegram Error {error_code}] {error_desc}")

            # Hướng dẫn chi tiết cho lỗi phổ biến: 'chat not found'
            if "chat not found" in error_desc.lower():
                print(
                    "\n💡 [HƯỚNG DẪN KÍCH HOẠT BOT]:\n"
                    "   Telegram yêu cầu người dùng phải mở bot và bấm nút /start trước\n"
                    "   thì bot mới có quyền gửi tin nhắn đến bạn.\n"
                    "   👉 Hãy mở ứng dụng Telegram, tìm bot và bấm START, sau đó chạy lại script!"
                )
            return False

    except requests.exceptions.RequestException as req_err:
        print(f"❌ [Network Error] Lỗi kết nối tới Telegram API: {req_err}")
        return False
    except Exception as e:
        print(f"❌ [Error] Lỗi khi gửi tin Telegram: {e}")
        return False


# -----------------------------------------------------------------------------
# 3. MAIN WORKFLOW EXECUTION (OIPO PIPELINE)
# -----------------------------------------------------------------------------
def main():
    print("=" * 65)
    print("🤖 AUTOMATION WORKFLOW: LẤY TIN AI & GỬI TELEGRAM (MÔ HÌNH OIPO)")
    print("=" * 65)

    # 1. Kiểm tra đầu vào môi trường
    if not BOT_TOKEN or not CHAT_ID:
        print("❌ Lỗi: Vui lòng kiểm tra file .env. Cần có BOT_TOKEN và CHAT_ID.")
        sys.exit(1)

    print("🔑 Đã nạp an toàn thông tin BOT_TOKEN và CHAT_ID từ .env")

    # 2. [Process 1] Lấy 1 tin AI mới nhất từ RSS
    print("⏳ [1/4] Đang lấy tin AI mới nhất từ Google News RSS...")
    news = get_ai_news()
    if not news:
        print("❌ Không thể tiếp tục quy trình do không lấy được tin.")
        sys.exit(1)

    raw_title = news["title"]
    news_link = news["link"]
    print(f"   📰 Tiêu đề gốc (EN): {raw_title}")
    print(f"   🔗 Liên kết: {news_link}")

    # 3. [Process 2] Dịch sang tiếng Việt (không dùng googletrans)
    print("⏳ [2/4] Đang dịch tiêu đề sang tiếng Việt (sử dụng requests)...")
    translated_title = translate_to_vi(raw_title)
    print(f"   🇻🇳 Tiêu đề dịch (VI): {translated_title}")

    # 4. [Process 3] Định dạng bản tin theo mẫu chuẩn
    # Format:
    # 🧠 Tin AI hôm nay - dd/mm/yyyy
    # • [nội dung]
    # Nguồn:
    # [link]
    # #AI #TinCongNghe
    today_str = datetime.now().strftime("%d/%m/%Y")
    formatted_message = (
        f"🧠 Tin AI hôm nay - {today_str}\n"
        f"• {translated_title}\n"
        f"Nguồn:\n"
        f"{news_link}\n"
        f"#AI #TinCongNghe"
    )

    print("⏳ [3/4] Bản tin đã được định dạng chuẩn:")
    print("-" * 50)
    print(formatted_message)
    print("-" * 50)

    # 5. [Process 4 & Output] Gửi tin nhắn vào Telegram
    print("⏳ [4/4] Đang gửi bản tin đến Telegram...")
    success = send_telegram(formatted_message)

    print("=" * 65)
    if success:
        print("🎉 QUY TRÌNH OIPO HOÀN TẤT: Tin tức đã xuất hiện trên Telegram!")
    else:
        print("⚠️ QUY TRÌNH KẾT THÚC VỚI CẢNH BÁO: Kiểm tra lại kết nối Telegram ở trên.")
    print("=" * 65)


if __name__ == "__main__":
    main()
