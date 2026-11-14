import os, sys

p_local = r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\database\register.db"
p_net = r"\\192.168.100.248\prdndata\ABEL\SFT\PDN TO QC REGISTER\database\register.db"

print("Checking local DB:", p_local, "-> Exists:", os.path.exists(p_local))
try:
    print("Checking network DB:", p_net, "-> Exists:", os.path.exists(p_net))
except Exception as e:
    print("Network DB check exception:", e)
