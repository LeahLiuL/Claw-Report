#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download the latest Vessel Bapfile.xlsx from the company SFTP server.

The scheduled rebuild pipeline (build_deploy.py) calls this first so the
GitHub Pages site always reflects the freshest spreadsheet on the FTP server
instead of a stale local copy.

IMPORTANT: the SFTP host 10.5.4.2 is an INTERNAL IP. VPN must be connected
before running, otherwise the connection times out (WinError 10060).

Connection defaults match the existing SFTP skills (sftp-vessel-data-update,
ptx-port-charges-update). Any value can be overridden by CLI flag or env var.
"""
import argparse
import os
import sys
import time

import paramiko

DEFAULTS = {
    "host": "10.5.4.2",
    "port": 6622,
    "user": "leah",
    "pass": "Fine@B!",
    "remote_path": "/finebi/Master Data - Leah/Vessel Bapfile.xlsx",
    "local_path": os.environ.get("BAPFILE_LOCAL") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "Vessel Bapfile.xlsx"),
}

ENV_MAP = {
    "host": "SFTP_HOST",
    "port": "SFTP_PORT",
    "user": "SFTP_USER",
    "pass": "SFTP_PASS",
    "remote_path": "SFTP_REMOTE_PATH",
    "local_path": "BAPFILE_LOCAL",
}


def resolve(name, args):
    env = ENV_MAP.get(name)
    if env and os.environ.get(env):
        return os.environ[env]
    val = getattr(args, name, None)
    if val:
        return val
    return DEFAULTS[name]


def main():
    ap = argparse.ArgumentParser(description="Fetch Vessel Bapfile.xlsx from SFTP")
    ap.add_argument("--host")
    ap.add_argument("--port", type=int)
    ap.add_argument("--user")
    ap.add_argument("--pass")
    ap.add_argument("--remote-path")
    ap.add_argument("--local-path")
    args = ap.parse_args()

    host = resolve("host", args)
    port = int(resolve("port", args))
    user = resolve("user", args)
    pw = resolve("pass", args)
    remote = resolve("remote_path", args)
    local = resolve("local_path", args)

    os.makedirs(os.path.dirname(local), exist_ok=True)

    print(f"[SFTP] connect {user}@{host}:{port} (VPN must be up)", flush=True)
    t = paramiko.Transport((host, port))
    sftp = None
    try:
        t.connect(username=user, password=pw)
        sftp = paramiko.SFTPClient.from_transport(t)
        st = sftp.stat(remote)
        rmt_time = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
        print(f"[SFTP] found {remote}  size={st.st_size:,}  mtime={rmt_time}", flush=True)
        tmp = local + ".part"
        sftp.get(remote, tmp)
        os.replace(tmp, local)
        print(f"[SFTP] saved -> {local}  ({os.path.getsize(local):,} bytes)", flush=True)
    except Exception as e:
        print(f"[SFTP] FAILED: {e}", file=sys.stderr, flush=True)
        if "10060" in str(e) or "timed out" in str(e).lower() or "Unable to connect" in str(e):
            print("[SFTP] Hint: 10.5.4.2 is an internal IP — connect VPN first.", file=sys.stderr, flush=True)
        sys.exit(1)
    finally:
        try:
            if sftp:
                sftp.close()
        except Exception:
            pass
        try:
            t.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
