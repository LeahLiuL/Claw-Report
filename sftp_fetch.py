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

Robust download: sftp.open('rb') + 256KB chunked read + keepalive 15s +
6 retries + resumable download (.dl.{PID}) + ZIP integrity verification +
safe os.replace with 5 retries.
"""
import argparse
import gc
import hashlib
import os
import socket
import sys
import time
import zipfile

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

CHUNK_SIZE = 256 * 1024  # 256 KB
SOCKET_TIMEOUT = 300  # 5 minutes per-socket
KEEPALIVE_INTERVAL = 15  # seconds
MAX_RETRIES = 6
TCP_PROBE_TIMEOUT = 15  # seconds, 3 retries


def resolve(name, args):
    env = ENV_MAP.get(name)
    if env and os.environ.get(env):
        return os.environ[env]
    val = getattr(args, name, None)
    if val:
        return val
    return DEFAULTS[name]


def _tcp_probe(host, port, timeout=TCP_PROBE_TIMEOUT, retries=3):
    """Check if the SFTP host is reachable on the given TCP port."""
    for i in range(retries):
        try:
            s = socket.create_connection((host, port), timeout=timeout)
            s.close()
            return True
        except Exception:
            if i < retries - 1:
                time.sleep(2)
    return False


def _verify_xlsx(path):
    """Verify that the downloaded file is a valid xlsx (ZIP) archive."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            return bad is None
    except Exception:
        return False


def _safe_replace(src, dst, retries=5):
    """os.replace with retry for WinError 32 (file in use)."""
    for i in range(retries):
        try:
            os.replace(src, dst)
            return True
        except OSError as e:
            if i < retries - 1:
                gc.collect()
                time.sleep(0.5)
            else:
                raise


def _download_chunked(sftp, remote_path, local_path, remote_size):
    """Download file using sftp.open('rb') + 256KB chunked read with resumable download."""
    pid = os.getpid()
    tmp = local_path + f".dl.{pid}"

    # Check for existing partial download (resumable)
    existing = os.path.getsize(tmp) if os.path.exists(tmp) else 0
    if existing > 0:
        print(f"[SFTP] resuming from {existing:,} bytes ({existing/remote_size*100:.1f}%)", flush=True)

    rfile = sftp.open(remote_path, "rb")
    try:
        rfile.settimeout(SOCKET_TIMEOUT)
        # Seek to offset for resumable download
        if existing > 0:
            rfile.seek(existing)
        with open(tmp, "ab" if existing > 0 else "wb") as f:
            offset = existing
            last_report = time.time()
            while True:
                data = rfile.read(CHUNK_SIZE)
                if not data:
                    break
                f.write(data)
                offset += len(data)
                now = time.time()
                if now - last_report >= 10:
                    pct = offset / remote_size * 100 if remote_size else 0
                    print(f"[SFTP] {offset:,}/{remote_size:,} bytes ({pct:.1f}%)", flush=True)
                    last_report = now
            print(f"[SFTP] download complete: {offset:,} bytes", flush=True)
    finally:
        try:
            rfile.close()
        except Exception:
            pass

    # Verify downloaded size
    actual = os.path.getsize(tmp)
    if actual != remote_size:
        print(f"[SFTP] WARNING: size mismatch downloaded={actual:,} vs remote={remote_size:,}", flush=True)
        os.remove(tmp)
        return False

    # Verify ZIP integrity
    if not _verify_xlsx(tmp):
        print("[SFTP] WARNING: ZIP integrity check failed, discarding download", flush=True)
        os.remove(tmp)
        return False

    # Safe replace
    _safe_replace(tmp, local_path)
    return True


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

    # VPN probe
    print(f"[SFTP] probing TCP {host}:{port} (VPN must be up)...", flush=True)
    if not _tcp_probe(host, port):
        print("[SFTP] FAILED: VPN not connected (10.5.4.2 is internal IP, connect VPN first)", file=sys.stderr, flush=True)
        sys.exit(1)
    print("[SFTP] VPN OK: host reachable", flush=True)

    # Get remote file info first (separate connection)
    remote_size = None
    remote_mtime = None

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"[SFTP] attempt {attempt}/{MAX_RETRIES}: connect {user}@{host}:{port}", flush=True)
        sock = None
        t = None
        sftp = None
        try:
            sock = socket.create_connection((host, port), timeout=SOCKET_TIMEOUT)
            sock.settimeout(SOCKET_TIMEOUT)
            t = paramiko.Transport(sock)
            t.banner_timeout = 120
            t.start_client(timeout=120)
            t.auth_password(username=user, password=pw)
            t.set_keepalive(KEEPALIVE_INTERVAL)

            sftp = paramiko.SFTPClient.from_transport(t)

            if remote_size is None:
                st = sftp.stat(remote)
                remote_size = st.st_size
                remote_mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
                print(f"[SFTP] found {remote}  size={remote_size:,}  mtime={remote_mtime}", flush=True)

                # Skip download if local file already matches
                if os.path.exists(local) and os.path.getsize(local) == remote_size:
                    local_mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(local)))
                    if _verify_xlsx(local):
                        print(f"[SFTP] local file already matches (size={remote_size:,}, mtime={remote_mtime}), skipping download", flush=True)
                        print(f"[SFTP] saved -> {local}  ({os.path.getsize(local):,} bytes) [already current]", flush=True)
                        return

            # Download with chunked read + resumable
            ok = _download_chunked(sftp, remote, local, remote_size)
            if ok:
                print(f"[SFTP] saved -> {local}  ({os.path.getsize(local):,} bytes)", flush=True)
                return
            else:
                print(f"[SFTP] attempt {attempt} failed integrity/size check, retrying...", flush=True)

        except Exception as e:
            err = str(e)
            print(f"[SFTP] attempt {attempt} error: {err}", flush=True)
            if "10060" in err or "timed out" in err.lower() or "Unable to connect" in err:
                print("[SFTP] Hint: 10.5.4.2 is an internal IP — connect VPN first.", file=sys.stderr, flush=True)
                sys.exit(1)
        finally:
            try:
                if sftp:
                    sftp.close()
            except Exception:
                pass
            try:
                if t:
                    t.close()
            except Exception:
                pass
            gc.collect()

        if attempt < MAX_RETRIES:
            wait = 5 + attempt * 5
            print(f"[SFTP] retrying in {wait}s ...", flush=True)
            time.sleep(wait)

    print("[SFTP] FAILED: all retries exhausted", file=sys.stderr, flush=True)
    sys.exit(1)


if __name__ == "__main__":
    main()
