import requests
import time

def send_line_message(token, message, retries=3, delay=5):
    """
    發送 Line Notify 訊息，包含重試與逾時機制。
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    # 使用 data 而非 params，將訊息放入 POST Body，避免 URL 過長
    payload = {"message": message}
    
    for i in range(retries):
        try:
            # 增加 timeout 設定 (連線 10 秒, 讀取 30 秒)
            response = requests.post(
                "https://notify-api.line.me/api/notify", 
                headers=headers, 
                data=payload, 
                timeout=(10, 30)
            )
            response.raise_for_status()
            print("✅ Line 通知發送成功")
            return True
        except requests.exceptions.RequestException as e:
            print(f"⚠️ Line 通知發送嘗試 {i+1} 失敗: {e}")
            if i < retries - 1:
                print(f"   將在 {delay} 秒後重試...")
                time.sleep(delay)
            else:
                print("❌ 已達最大重試次數，Line 通知發送失敗。")
                print("💡 提示：若出現 'Failed to resolve', 請檢查網路連線或 DNS 設定 (建議改用 Google DNS 8.8.8.8)。")
                return False

if __name__ == '__main__':
    pass
