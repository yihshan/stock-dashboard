import requests
import json
import time

def send_line_message(channel_access_token, user_id, message):
    """
    使用 Line Messaging API 發送訊息。
    優先使用廣播模式 (Broadcast)，若無 User ID 則發送給所有好友。
    """
    if not channel_access_token:
        print("⚠️ 缺少 Line Channel Access Token，取消發送。")
        return False

    # 根據是否有 User ID 決定使用廣播還是推播
    if user_id and user_id.strip():
        url = "https://api.line.me/v2/bot/message/push"
        payload = {
            "to": user_id.strip(),
            "messages": [{"type": "text", "text": message}]
        }
        mode = "Push (特定用戶)"
    else:
        url = "https://api.line.me/v2/bot/message/broadcast"
        payload = {
            "messages": [{"type": "text", "text": message}]
        }
        mode = "Broadcast (廣播)"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {channel_access_token}"
    }

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            print(f"🚀 嘗試發送 Line {mode} 訊息 (第 {attempt} 次)...")
            response = requests.post(url, headers=headers, data=json.dumps(payload), timeout=15)
            
            if response.status_code == 200:
                print(f"✅ Line Messaging API 訊息發送成功 ({mode})")
                return True
            else:
                print(f"⚠️ Line API 回傳錯誤 ({response.status_code}): {response.text}")
                if response.status_code == 401:
                    print("💡 提示：請檢查 Channel Access Token 是否正確或已過期。")
                    break
        except Exception as e:
            print(f"⚠️ 第 {attempt} 次發送失敗: {str(e)}")
        
        if attempt < max_retries:
            time.sleep(5)
            
    print("❌ 已達最大重試次數，Line 通知發送失敗。")
    return False
