import subprocess
import re
import urllib.request
import json
import base64
import time
import socket
import os
import sys

def get_github_token():
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    try:
        res = subprocess.run(
            ["git", "credential", "fill"],
            input="protocol=https\nhost=github.com\n",
            capture_output=True,
            text=True
        )
        for line in res.stdout.splitlines():
            if line.startswith("password="):
                return line.split("=", 1)[1].strip()
    except Exception:
        pass
    return ""

GITHUB_REPO = "abellrajuu/IPA-PRODUCTION-TEST-REPORT"
GITHUB_BRANCH = "gh-pages"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def is_port_open(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1.0)
        return s.connect_ex((host, port)) == 0

def ensure_server_running():
    if is_port_open("127.0.0.1", 5005):
        print("[OK] Flask backend server is already running on port 5005.")
        return
    print("[...] Starting Flask backend server on port 5005...")
    py_script = os.path.join(BASE_DIR, "backend_python.py")
    subprocess.Popen([sys.executable, py_script], cwd=BASE_DIR)
    for _ in range(15):
        time.sleep(1)
        if is_port_open("127.0.0.1", 5005):
            print("[OK] Flask backend server is up and listening on port 5005.")
            return
    print("[WARNING] Flask server did not respond on port 5005 within 15s.")

def publish_to_github(public_url):
    token = get_github_token()
    if not token:
        print("[!] No GitHub token available. Skipping cloud registry sync.")
        return False
    print(f"[...] Publishing {public_url} to GitHub Cloud Registry...")
    get_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/live_url.json?ref={GITHUB_BRANCH}"
    req = urllib.request.Request(get_url, headers={
        "Authorization": f"token {token}",
        "User-Agent": "Antigravity-Agent"
    })
    sha = None
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            sha = data['sha']
    except Exception as e:
        print("[!] Note fetching existing SHA:", e)

    payload = {
        "url": public_url,
        "version_code": 5,
        "version_name": "3.1.5",
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    content_b64 = base64.b64encode(json.dumps(payload, indent=2).encode('utf-8')).decode('utf-8')

    put_body = {
        "message": f"Update live mobile server URL: {public_url}",
        "content": content_b64,
        "branch": GITHUB_BRANCH
    }
    if sha:
        put_body["sha"] = sha

    put_req = urllib.request.Request(get_url, data=json.dumps(put_body).encode('utf-8'), method='PUT', headers={
        "Authorization": f"token {token}",
        "User-Agent": "Antigravity-Agent",
        "Content-Type": "application/json"
    })

    try:
        with urllib.request.urlopen(put_req) as resp:
            res = json.loads(resp.read().decode())
            print(f"[OK] Successfully updated GitHub! Commit: {res['commit']['sha'][:8]}")
            return True
    except Exception as e:
        print(f"[ERROR] Failed to update GitHub: {e}")
        return False

def update_local_text_file(public_url):
    link_path = os.path.join(BASE_DIR, "CURRENT_MOBILE_LINK.txt")
    banner = f"""===================================================================
      LIVE MOBILE DATA LINK (PHONE & TABLET ACCESS)
===================================================================

Active Public Server URL:
  👉 {public_url}

GitHub Cloud Registry (Auto-checked by Android APK):
  👉 https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/live_url.json

Status: ONLINE (Updated: {time.strftime('%Y-%m-%d %H:%M:%S')})
===================================================================
Zero Link Pasting:
The Android APK automatically discovers this address from GitHub.
Employees on personal 4G/5G mobile data simply open the app to connect.
===================================================================
"""
    try:
        with open(link_path, "w", encoding="utf-8") as f:
            f.write(banner)
    except Exception as e:
        print(f"[!] Warning writing link file: {e}")

def run_tunnel_loop():
    ensure_server_running()
    ssh_exe = r"C:\Windows\System32\OpenSSH\ssh.exe"
    if not os.path.exists(ssh_exe):
        ssh_exe = "ssh"

    cmd = [
        ssh_exe,
        "-R", "80:127.0.0.1:5005",
        "-o", "StrictHostKeyChecking=no",
        "-o", "ServerAliveInterval=30",
        "nokey@localhost.run"
    ]

    while True:
        print("\n[...] Establishing secure public tunnel over port 443/80...")
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
            tunnel_url = None
            for line in proc.stdout:
                line_str = line.strip()
                if line_str:
                    print("[Tunnel]", line_str)
                m = re.search(r'(https://[a-zA-Z0-9.-]+\.(?:lhr\.life|lhrtunnel\.link|lhr\.rocks))', line_str)
                if m and not tunnel_url:
                    base_url = m.group(1)
                    tunnel_url = base_url + "/portal"
                    print("\n" + "="*65)
                    print("  🎉 GLOBAL MOBILE DATA LINK IS LIVE!")
                    print("="*65)
                    print(f"  Public Portal URL : {tunnel_url}")
                    print(f"  APK Download URL  : {base_url}/download-app")
                    print("="*65)
                    update_local_text_file(tunnel_url)
                    publish_to_github(tunnel_url)
                    print("="*65 + "\n")
            proc.wait()
            print("[!] Tunnel disconnected. Reconnecting in 5 seconds...")
            time.sleep(5)
        except Exception as e:
            print(f"[!] Tunnel loop error: {e}. Retrying in 5 seconds...")
            time.sleep(5)

if __name__ == "__main__":
    run_tunnel_loop()
