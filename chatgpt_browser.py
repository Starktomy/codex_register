import os
import re
import time
import json
import uuid
import random
import string
import imaplib
import email
from datetime import datetime, timedelta
from pathlib import Path
import requests as sync_req

# DrissionPage 必须安装：pip install DrissionPage --break-system-packages
from DrissionPage import ChromiumPage, ChromiumOptions

OUT_DIR = Path(__file__).parent.resolve()
PROXY_FILE = OUT_DIR / "proxy.txt"
OUTLOOK_ACCOUNTS_FILE = OUT_DIR / "accounts.txt"

def load_proxies():
    if not PROXY_FILE.exists():
        return []
    return [line.strip() for line in PROXY_FILE.read_text("utf-8").splitlines() if line.strip()]

def load_outlook_accounts():
    if not OUTLOOK_ACCOUNTS_FILE.exists():
        return []
    accs = []
    for line in OUTLOOK_ACCOUNTS_FILE.read_text("utf-8").splitlines():
        line = line.strip()
        if not line: continue
        parts = line.split("----")
        if len(parts) >= 4:
            accs.append({
                "email": parts[0],
                "password": parts[1],
                "client_id": parts[2],
                "refresh_token": parts[3]
            })
    return accs

def get_outlook_access_token(client_id, refresh_token):
    try:
        url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        data = {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token
        }
        r = sync_req.post(url, data=data, timeout=15)
        if r.status_code == 200:
            return r.json().get("access_token")
    except Exception as e:
        print(f"[!] 获取 Outlook Access Token 失败: {e}")
    return None

def fetch_outlook_otp(email_addr, access_token):
    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com")
        auth_string = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01"
        mail.authenticate("XOAUTH2", lambda x: auth_string.encode("utf-8"))
        
        for folder in ["Junk", '"Junk Email"', "INBOX"]:
            try:
                status, _ = mail.select(folder)
                if status != "OK": continue
                status, messages = mail.search(None, 'UNSEEN', 'FROM', '"OpenAI"')
                if status == "OK" and messages[0]:
                    msg_ids = messages[0].split()
                    for msg_id in reversed(msg_ids[-3:]):
                        status, msg_data = mail.fetch(msg_id, "(RFC822)")
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                body = ""
                                if msg.is_multipart():
                                    for part in msg.walk():
                                        if part.get_content_type() in ["text/plain", "text/html"]:
                                            payload = part.get_payload(decode=True)
                                            if payload: body += payload.decode(errors='ignore')
                                else:
                                    payload = msg.get_payload(decode=True)
                                    if payload: body = payload.decode(errors='ignore')
                                
                                mt = re.search(r"(?<!#)\b(\d{6})\b", body) or re.search(r"(?<!#)\b(\d{6})\b", str(msg["Subject"]))
                                if mt:
                                    mail.store(msg_id, '+FLAGS', '\\Seen')
                                    mail.logout()
                                    return mt.group(1)
            except Exception: pass
        mail.logout()
    except Exception as e:
        print(f"  [Debug] IMAP 获取验证码异常: {e}")
    return None

def setup_outlook(account):
    email_addr = account["email"]
    client_id = account["client_id"]
    refresh_token = account["refresh_token"]
    
    print(f"  [*] 尝试刷新 {email_addr} 的 Token...")
    access_token = get_outlook_access_token(client_id, refresh_token)
    if not access_token:
        print("  [Error] 无法获取 Access Token")
        return None, None

    def fetch_code():
        print("  [*] 正在等待验证码 (最多等待约8分钟)...")
        for _ in range(60):
            otp = fetch_outlook_otp(email_addr, access_token)
            if otp: return otp
            time.sleep(8)
        return None
        
    return email_addr, fetch_code

def _random_name() -> str:
    return ''.join(random.choice(string.ascii_lowercase) for _ in range(random.randint(5, 9))).capitalize()

def run_browser(account, proxy):
    email_addr, code_fetcher = setup_outlook(account)
    if not email_addr:
        return None
        
    print(f"[*] 启动 DrissionPage 浏览器 (代理: {proxy})...")
    co = ChromiumOptions()
    co.set_browser_path(r'/usr/bin/google-chrome')
    co.set_argument('--no-sandbox')
    co.set_argument('--disable-gpu')
    if proxy:
        co.set_argument(f'--proxy-server=http://{proxy}')
        
    page = ChromiumPage(addr_or_opts=co)
    
    try:
        print("[*] 第一步：跳转至 chatgpt.com 并绕过 CF 验证码")
        page.get("https://chatgpt.com/")
        time.sleep(5)
        
        signup_btn = page.ele('text:Sign up', timeout=15)
        if not signup_btn:
            print("[Error] 未能找到 Sign up 按钮，CF可能被拦截或代理不可用")
            return None
            
        signup_btn.click()
        time.sleep(5)
        
        print("[*] 第二步：输入注册邮箱")
        email_input = page.ele('@type=email', timeout=10)
        if not email_input:
            print("[Error] 未能找到邮箱输入框")
            return None
            
        email_input.input(email_addr)
        page.ele('@type=submit').click()
        time.sleep(5)
        
        # 此时已经触发 passwordless_signup 的验证码
        code = code_fetcher()
        if not code:
            print("[Error] 提取验证码失败或超时")
            return None
            
        print(f"[*] 成功提取验证码: {code}，正在填入...")
        page.ele('@name=code').input(code)
        page.ele('@type=submit').click()
        
        time.sleep(10)
        print(f"[*] 当前所在页面 URL: {page.url}")
        
        if "about-you" in page.url:
            print("[*] 第三步：该邮箱尚未完成注册，正在自动填写个人信息 (about-you)...")
            page.ele('@name=name').input(_random_name())
            page.ele('@name=age').input(str(random.randint(20, 45)))
            page.ele('@type=submit').click()
            time.sleep(10)
            print(f"[*] 提交注册信息后 URL: {page.url}")
            
            # 检查是否因为 IP 质量出现 Terms of Use 的注册限制报错 (registration_disallowed)
            if "We can't create your account" in page.html or "Terms of Use" in page.html:
                print(f"[Error] 账户注册被拒绝 (registration_disallowed)！")
                print(f"  -> CF验证已通过，但由于代理 IP 质量差被系统风控拦截。请更换干净的家庭宽带节点！")
                return None
            print("[*] 注册成功！")
        else:
            print("[*] 账号已注册，直接登录成功。")
            
        print("[*] 第四步：获取最终 Token")
        cookies = page.cookies(as_dict=True)
        session_token = cookies.get("__Secure-next-auth.session-token")
        
        if session_token:
            print("[🎉] 成功拿到 session-token!")
            return {"session_token": session_token}, email_addr, account["password"]
        else:
            print("[Error] 登录/注册成功但未找到 session-token，请检查网络或更换IP")
            return None
            
    except Exception as e:
        print(f"[Error] 浏览器流程异常: {e}")
    finally:
        page.quit()

def main():
    print("========================================")
    print("🚀 OpenAI 基于真实浏览器的终极注册机 (绕过 CF 验证码)")
    print("========================================")
    
    accounts = load_outlook_accounts()
    proxies = load_proxies()
    if not accounts:
        print("[!] 请确保目录下存在 accounts.txt")
        return
        
    print(f"[*] 成功加载 {len(accounts)} 个 Outlook 账号")
    if proxies: print(f"[*] 成功加载 {len(proxies)} 个代理")
    
    tokens_dir = OUT_DIR / "tokens"
    tokens_dir.mkdir(parents=True, exist_ok=True)
    
    count = 0
    for current_account in accounts:
        count += 1
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] >>> 开始第 {count} 次注册流程 ({current_account['email']}) <<<")
        current_proxy = random.choice(proxies) if proxies else None
        
        run_result = run_browser(current_account, current_proxy)
        if run_result:
            token_json, email_addr, password = run_result
            fname_email = email_addr.replace("@", "_")

            file_path = tokens_dir / f"token_{fname_email}_{int(time.time())}.json"
            file_path.write_text(json.dumps(token_json), encoding="utf-8")
            print(f"[🎉] 成功获取 Token！已保存至: {file_path}")

            acc_file = tokens_dir / "accounts_openai.txt"
            with open(acc_file, "a", encoding="utf-8") as f:
                f.write(f"{email_addr}----{password}----{token_json['session_token']}\n")
        else:
            print("[-] 本次注册流程断开。")
            
        print("[*] 冷却 15 秒...")
        time.sleep(15)

if __name__ == "__main__":
    main()
