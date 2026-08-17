#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pre-flight checker for CUL Daily Movement auto-update on non-leahliu hosts
(e.g. culadmin). Confirms the environment can produce a NON-EMPTY Maintenance
view BEFORE shipping a build, so a blank MAINT never goes live silently.

What it checks:
  0) Local C:\\CULINES copy present?  (leahliu machine -> no VPN needed)
  1) paramiko importable
  2) VPN / SFTP reachable (10.5.4.2:6622)  [fast socket pre-check + paramiko]
  3) MAINT xlsx exists on SFTP at the configured remote path

Usage:
  python check_maint_env.py
  MAINT_SFTP_REMOTE=/other/path.xlsx python check_maint_env.py
"""
import os
import sys
import time
import socket

HOST = os.environ.get("MAINT_SFTP_HOST", "10.5.4.2")
PORT = int(os.environ.get("MAINT_SFTP_PORT", "6622"))
USER = os.environ.get("MAINT_SFTP_USER", "leah")
PASS = os.environ.get("MAINT_SFTP_PASS", "Fine@B!")
REMOTE = os.environ.get(
    "MAINT_SFTP_REMOTE",
    "/finebi/Master Data - Leah/Vessel_Schedule_Maintain_Over_Time_Port_Log.xlsx",
)
LOCAL = r"C:\CULINES\Claw Report\Vessel_Schedule_Maintain_Over_Time_Port_Log.xlsx"


def section(t):
    print(f"\n=== {t} ===")


def main():
    # 0) local copy (leahliu machine)
    section("0) Local C:\\CULINES copy")
    if os.path.exists(LOCAL):
        print(f"  FOUND local copy -> {LOCAL}")
        print("  This host can use the local copy; no VPN needed. GO.")
        sys.exit(0)
    print("  Not present (expected on culadmin). Will rely on SFTP.")

    # 1) paramiko
    section("1) paramiko import")
    try:
        import paramiko  # noqa: F401
        print(f"  paramiko OK")
    except Exception as e:
        print(f"  MISSING: {e}")
        print("  FIX: pip install paramiko")
        sys.exit(1)

    # 2) VPN / SFTP reachable (fast socket pre-check)
    section(f"2) SFTP reachability {USER}@{HOST}:{PORT}")
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(6)
    try:
        s.connect((HOST, PORT))
        print("  socket CONNECTED (VPN is up)")
    except Exception as e:
        print(f"  UNREACHABLE: {e}")
        if "10060" in str(e) or "timed out" in str(e).lower():
            print("  FIX: connect VPN, then retry.")
        sys.exit(1)
    finally:
        try:
            s.close()
        except Exception:
            pass

    # 3) remote MAINT file
    section(f"3) MAINT file on SFTP: {REMOTE}")
    import paramiko as _p

    t = None
    try:
        t = _p.Transport((HOST, PORT))
        t.connect(username=USER, password=PASS)
        sftp = _p.SFTPClient.from_transport(t)
        st = sftp.stat(REMOTE)
        mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
        print(f"  FOUND  size={st.st_size:,}  mtime={mtime}")
        sftp.close()
    except Exception as e:
        print(f"  NOT FOUND / ERROR: {e}")
        print(f"  FIX: verify MAINT_SFTP_REMOTE (current='{REMOTE}')")
        sys.exit(1)
    finally:
        if t:
            try:
                t.close()
            except Exception:
                pass

    print("\nRESULT: GO ✅ auto-update will populate the Maintenance view")
    sys.exit(0)


if __name__ == "__main__":
    main()
