# -*- coding: utf-8 -*-
import socket
import os

try:
    import qrcode
    HAS_QR = True
except ImportError:
    HAS_QR = False

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def main():
    ip = get_ip()
    port = int(os.environ.get("PORT", 5005))
    mdns_url = f"http://DESKTOP-ITTFPI2.local:{port}/portal"
    ip_url = f"http://{ip}:{port}/portal"

    print("===================================================================")
    print("       MOBILE & ANDROID WI-FI AUTO-CONNECTOR")
    print("===================================================================")
    print(f"Current Wi-Fi IP Address:   {ip}")
    print(f"Local Hostname URL:        {mdns_url}  (Works on ANY Wi-Fi!)")
    print(f"Direct Wi-Fi IP URL:       {ip_url}")
    print("-------------------------------------------------------------------")

    if HAS_QR:
        print("Scan this QR code with your mobile camera to open the Portal:\n")
        try:
            qr = qrcode.QRCode(border=1)
            qr.add_data(ip_url)
            qr.make(fit=True)
            qr.print_ascii(invert=True)
        except Exception:
            pass
    else:
        print(f"Open this URL on your phone browser: {ip_url}")

    print("")
    print("===================================================================")
    print("HOW TO CONNECT ACROSS DIFFERENT WI-FI NETWORKS:")
    print("1. Option A (Easiest): In phone browser, type: " + ip_url)
    print(f"2. Option B (Hostname): In phone browser, type: {mdns_url}")
    print("   (Hostname works if mDNS / Bonjour is active on your Wi-Fi router)")
    print("3. Option C (Public Mobile Link): Run START_MOBILE_DATA_LINK.bat")
    print("   to connect from your phone anywhere using 4G / 5G Mobile Data!")
    print("===================================================================")
    input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()
