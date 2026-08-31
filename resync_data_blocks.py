# -*- coding: utf-8 -*-
"""把远端(culadmin)HTML 的指定数据块注入到本地用最新 gen_html.py 生成的 HTML 中。

规则（2026-08-31 核对）：
  - TODAY_DATA   取远端：本地 rebuilt.xlsx 是残缺快照（44 船 / 无已下线标记），远端 54 船
  - MAINT_DATA   取远端：本地源 9283 条，远端 24774 条
  - AGENT_BY_PORT 保留本地：远端那份是空 {}（culadmin 机器读不到 P: 盘联系人表）
"""
import io, os, re, subprocess, sys

REMOTE = sys.argv[1] if len(sys.argv) > 1 else 'b81cf0b'
HTML = 'cul_daily_movement.html'

# (块名, 取值来源)  remote = 用远端, local = 保留本地
PLAN = [
    ('TODAY_DATA',     'remote'),
    ('MAINT_DATA',     'remote'),
    ('AGENT_BY_PORT',  'local'),
]

def sh(cmd):
    p = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    return p.returncode, (p.stdout or '') + (p.stderr or '')

def block(text, name):
    m = re.search(r"^\s*const\s+" + re.escape(name) + r"\s*=\s*", text, re.M)
    if not m:
        return None, None, None
    start = m.end()
    if text[start] not in '{[':
        return None, None, None
    depth = 0; i = start; in_str = False; esc = False
    while i < len(text):
        ch = text[i]
        if in_str:
            if esc: esc = False
            elif ch == '\\': esc = True
            elif ch == '"': in_str = False
        else:
            if ch == '"': in_str = True
            elif ch in '{[': depth += 1
            elif ch in '}]':
                depth -= 1
                if depth == 0:
                    return start, i + 1, text[start:i+1]
        i += 1
    return None, None, None

rc, remote_html = sh('git show %s:%s' % (REMOTE, HTML))
if rc != 0:
    print('FAILED reading remote html:\n' + remote_html)
    sys.exit(1)
local_html = io.open(HTML, encoding='utf-8').read()

for name, src in PLAN:
    rs, re_, rv = block(remote_html, name)
    ls, le_, lv = block(local_html, name)
    if rv is None or lv is None:
        print('  skip %s (missing: remote=%s local=%s)' % (name, rv is not None, lv is not None))
        continue
    if src == 'remote':
        local_html = local_html[:ls] + rv + local_html[le_:]
        print('  %-16s <- remote (%d bytes, was local %d)' % (name, len(rv), len(lv)))
    else:
        print('  %-16s <- keep local (%d bytes, remote was %d)' % (name, len(lv), len(rv)))

io.open(HTML, 'w', encoding='utf-8', newline='').write(local_html)
print('wrote %s (%d bytes)' % (HTML, len(local_html)))

# maint_snapshot.json 恢复远端版本（本地重生成的是残缺的 9283 条版本）
rc, out = sh('git show %s:maint_snapshot.json' % REMOTE)
if rc == 0:
    io.open('maint_snapshot.json', 'w', encoding='utf-8', newline='').write(out)
    print('restored maint_snapshot.json from %s (%d bytes)' % (REMOTE, len(out)))
else:
    print('WARN: cannot restore maint_snapshot.json: ' + out)
