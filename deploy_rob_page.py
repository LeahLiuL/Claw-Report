"""把【已部署的生产加密页】用修复版模板重新生成。

做法(与 build_preview.py 对称, 但输出加密部署页而非明文预览):
  1. 读本地 rob_oil_report.html(= 已 checkout 的最新生产版本, 数据来自 culadmin 上次推送)
  2. 用 cryptojs_decrypt 解出 payload(当前生产数据)
  3. 以修复版 templates/rob.html 为骨架, 把 payload 重新加密注入 __ENC__ 占位
  4. 写回 rob_oil_report.html(保留密码锁 + AES 解密逻辑)

只升级模板(红线 TDR vs Actual 逻辑), 数据内容保持为当前生产数据, 不改动任何数据文件。
用法: python deploy_rob_page.py
"""
import os
import re
import json
import sys

import rob_refresh as R

BASE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(BASE, "templates", "rob.html")
OUT = os.path.join(BASE, "rob_oil_report.html")


def build():
    # 1) 读当前已部署页(= 本地 checkout 的最新生产版本)
    html = open(OUT, encoding="utf-8").read()
    m = re.search(r'var ENC = "([^"]+)";', html)
    if not m:
        raise SystemExit("[ERR] rob_oil_report.html 里找不到 var ENC, 请先确认已 checkout 生产版本")
    enc = m.group(1)
    obj = json.loads(R.cryptojs_decrypt(enc, R.PASSWORD))
    print("解密生产数据: vessels=%d history=%d voyages=%d bunkering=%d"
          % (len(obj.get("vessels", [])), len(obj.get("history", [])),
             len(obj.get("voyages", [])), len(obj.get("bunkering", {}))))

    # 2) 用修复版模板重新加密(数据不变, 仅升级模板)
    new_enc = R.cryptojs_encrypt(json.dumps(obj, ensure_ascii=False), R.PASSWORD)
    # 自校验: Python 端能解开且结构一致
    back = json.loads(R.cryptojs_decrypt(new_enc, R.PASSWORD))
    assert back.get("updated") == obj.get("updated"), "updated 字段不一致"
    assert "history" in back and "vessels" in back

    tpl = open(TPL, encoding="utf-8").read()
    if "__ENC__" not in tpl:
        raise SystemExit("[ERR] 模板 templates/rob.html 无 __ENC__ 占位, 结构可能已变")
    tpl = tpl.replace("__ENC__", new_enc)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(tpl)

    # 3) 自检: 加密页保留密码锁, 且含修复版红线逻辑
    chk = open(OUT, encoding="utf-8").read()
    ok = ('var ENC = "' in chk) and ('function renderTDR' in chk) \
        and ("r.rtype === 'NOON' || !r.rtype" in chk) and ('doUnlock' in chk) \
        and ('id="lock"' in chk)
    print("deploy ->", OUT, "(%d bytes)" % len(chk))
    print("自检:", "OK" if ok else "FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(build())
