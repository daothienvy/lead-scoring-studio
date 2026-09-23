#!/usr/bin/env python3
"""
Script gửi file Dashboard Ngân Sách Beta Solutions qua Telegram Bot API.
Sử dụng thư viện chuẩn của Python (urllib.request), không cần cài đặt thêm thư viện ngoài.

Cách sử dụng:
  Cách 1 (Truyền tham số qua dòng lệnh):
    python3 scripts/send_to_telegram.py --token <BOT_TOKEN> --chat_id <CHAT_ID>

  Cách 2 (Sử dụng biến môi trường):
    export TELEGRAM_BOT_TOKEN="your_bot_token"
    export TELEGRAM_CHAT_ID="your_chat_id"
    python3 scripts/send_to_telegram.py
"""

import os
import sys
import json
import argparse
import urllib.request
import urllib.parse

def send_telegram_document(bot_token: str, chat_id: str, file_path: str, caption: str = ""):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    filename = os.path.basename(file_path)
    url = f"https://api.telegram.org/bot{bot_token}/sendDocument"

    # Boundary cho multipart/form-data
    boundary = "----WebKitFormBoundaryTelegramUploader7MA4YWxkTrZu0gW"
    
    with open(file_path, "rb") as f:
        file_bytes = f.read()

    body = bytearray()

    # Field: chat_id
    body.extend(f"--{boundary}\r\n".encode("utf-8"))
    body.extend(f'Content-Disposition: form-data; name="chat_id"\r\n\r\n'.encode("utf-8"))
    body.extend(f"{chat_id}\r\n".encode("utf-8"))

    # Field: parse_mode
    body.extend(f"--{boundary}\r\n".encode("utf-8"))
    body.extend(f'Content-Disposition: form-data; name="parse_mode"\r\n\r\n'.encode("utf-8"))
    body.extend(b"HTML\r\n")

    # Field: caption
    if caption:
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(f'Content-Disposition: form-data; name="caption"\r\n\r\n'.encode("utf-8"))
        body.extend(f"{caption}\r\n".encode("utf-8"))

    # Field: document (file)
    body.extend(f"--{boundary}\r\n".encode("utf-8"))
    body.extend(f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'.encode("utf-8"))
    body.extend(b"Content-Type: application/octet-stream\r\n\r\n")
    body.extend(file_bytes)
    body.extend(b"\r\n")

    # Closing boundary
    body.extend(f"--{boundary}--\r\n".encode("utf-8"))

    req = urllib.request.Request(
        url=url,
        data=bytes(body),
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "User-Agent": "BetaSolutions-Uploader/1.0"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as resp:
            resp_data = json.loads(resp.read().decode("utf-8"))
            if resp_data.get("ok"):
                print("✅ Gửi file qua Telegram thành công!")
                return resp_data
            else:
                print(f"❌ Telegram API trả về lỗi: {resp_data}")
                return resp_data
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        print(f"❌ HTTP Error {e.code}: {err_msg}")
        raise

def main():
    parser = argparse.ArgumentParser(description="Gửi Dashboard Ngân Sách qua Telegram")
    parser.add_argument("--token", default=os.getenv("TELEGRAM_BOT_TOKEN"), help="Telegram Bot Token")
    parser.add_argument("--chat_id", default=os.getenv("TELEGRAM_CHAT_ID"), help="Telegram Chat ID")
    parser.add_argument("--file", default="outputs/reports/BetaSolutions_NganSach_Dashboard.zip", help="Đường dẫn file cần gửi")
    parser.add_argument("--password", default="BetaSolutions@2026", help="Mật khẩu file ZIP")
    args = parser.parse_args()

    if not args.token or not args.chat_id:
        print("⚠️ Thiếu Bot Token hoặc Chat ID.")
        print("💡 Hướng dẫn nhanh:")
        print("  1. Tạo bot qua @BotFather trên Telegram để lấy Bot Token.")
        print("  2. Nhắn tin cho bot hoặc lấy Chat ID cá nhân/nhóm qua @userinfobot.")
        print("  3. Chạy lệnh:")
        print(f"     python3 scripts/send_to_telegram.py --token <BOT_TOKEN> --chat_id <CHAT_ID>")
        sys.exit(1)

    caption = (
        "<b>📊 BÁO CÁO NGÂN SÁCH BETA SOLUTIONS 2026</b>\n"
        "<i>Báo cáo Tài chính định kỳ · Trưởng phòng Tài chính</i>\n\n"
        "💰 <b>Tổng Ngân Sách:</b> 39.64 Tỷ VNĐ\n"
        "📊 <b>Tổng Chi Tiêu:</b> 39.98 Tỷ VNĐ (100.9% NS)\n"
        "🚨 <b>GD Vượt Ngân Sách:</b> 32 giao dịch (33.3%)\n\n"
        f"🔐 <b>Mật khẩu mở file:</b> <code>{args.password}</code>\n"
        "<i>Tải về, giải nén và mở file .html trên trình duyệt bất kỳ.</i>"
    )

    print(f"📤 Đang gửi file: {args.file} đến Chat ID: {args.chat_id}...")
    send_telegram_document(args.token, args.chat_id, args.file, caption)

if __name__ == "__main__":
    main()
