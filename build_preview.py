"""生成【免密预览页】供本地审阅样式(不改动生产加密页)。

做法:
  1. 读取已生成的 rob_oil_report.html, 取其 var ENC = "..."
  2. 用 rob_refresh.cryptojs_decrypt() 解出 payload(真实生产数据)
  3. 以 templates/rob.html 为骨架, 去掉密码锁 UI / AES 解密逻辑,
     把 payload 以明文 var DATA = {...} 嵌入, 打开即进页面

输出路径固定在工作区(仓库外), 因此明文数据不会进 git。
用法: python build_preview.py [输出路径]
"""
import os
import re
import sys
import json

BASE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(BASE, "templates", "rob.html")
OUT_HTML = os.path.join(BASE, "rob_oil_report.html")
DEFAULT_PREVIEW = r"C:\Users\leahliu\WorkBuddy\20260325163116\draft_preview.html"

LOCK_BLOCK = '''<div id="lock">
  <div class="lock-card">
    <h2>CUL ROB Bunker Report</h2>
    <div class="sub">This page is encrypted. Enter password to continue.</div>
    <input id="pwd" type="password" placeholder="Password" autofocus>
    <button onclick="doUnlock()">Unlock</button>
    <div id="lockErr"></div>
  </div>
</div>

'''

PWD_LISTENER = """document.getElementById('pwd').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doUnlock();
});
"""

# tryUnlock / doUnlock 整块(AES 解密 + 引用 #pwd / #lockErr)
UNLOCK_BLOCK = """function tryUnlock(pwd) {
  try {
    var dec = CryptoJS.AES.decrypt(ENC, pwd).toString(CryptoJS.enc.Utf8);
    if (!dec) return null;
    return JSON.parse(dec);
  } catch (e) { return null; }
}
function doUnlock() {
  var pwd = document.getElementById('pwd').value;
  var d = tryUnlock(pwd);
  if (d) {
    DATA = d; enterApp();
  } else {
    document.getElementById('lockErr').textContent = 'Wrong password, please try again';
  }
}
"""


def build(out_path=DEFAULT_PREVIEW):
    import rob_refresh as R

    html = open(OUT_HTML, encoding="utf-8").read()
    m = re.search(r'var ENC = "([^"]+)";', html)
    if not m:
        raise SystemExit("[ERR] rob_oil_report.html 里找不到 var ENC, 请先跑 rob_refresh.py")
    enc = m.group(1)
    payload = R.cryptojs_decrypt(enc, R.PASSWORD)
    obj = json.loads(payload)
    print("payload: vessels=%d history=%d voyages=%d bunkering=%d"
          % (len(obj.get("vessels", [])), len(obj.get("history", [])),
             len(obj.get("voyages", [])), len(obj.get("bunkering", {}))))

    tpl = open(TPL, encoding="utf-8").read()

    # 1) 去掉密码锁 UI
    if LOCK_BLOCK in tpl:
        tpl = tpl.replace(LOCK_BLOCK, "")
    else:
        tpl = re.sub(r'<div id="lock">.*?</div>\s*</div>\s*</div>', "", tpl, count=1, flags=re.S)

    # 2) app 默认可见(不再靠 doUnlock 切换)
    tpl = tpl.replace('<div id="app" style="display:none">', '<div id="app">')

    # 3) 明文注入 DATA, 去掉 ENC 占位
    old_assign = 'var ENC = "__ENC__";\nvar DATA = null, sortKey = "seq", sortAsc = true;'
    safe = payload.replace("</", "<\\/")          # 防止字符串里的 </script> 提前闭合
    new_assign = 'var DATA = %s;\nvar sortKey = "seq", sortAsc = true;' % safe
    if old_assign not in tpl:
        raise SystemExit("[ERR] 模板里找不到 DATA 赋值片段, 结构可能已变")
    tpl = tpl.replace(old_assign, new_assign)

    # 4) enterApp 里对已删除 #lock 的引用
    tpl = tpl.replace("  document.getElementById('lock').style.display = 'none';\n", "")

    # 5) 密码框回车监听 + AES 解锁函数整块
    tpl = tpl.replace(PWD_LISTENER, "/* preview: no password input */\n")
    tpl = tpl.replace(UNLOCK_BLOCK, "/* preview: no password gate */\n")

    # 6) 加载即进
    tpl = tpl.replace("</script>\n</body>", "enterApp();\n</script>\n</body>")
    if "enterApp();\n</script>" not in tpl:
        tpl = tpl.replace("</script>", "enterApp();\n</script>")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(tpl)

    # 自检
    chk = open(out_path, encoding="utf-8").read()
    bad = [t for t in ("doUnlock", "tryUnlock", "getElementById('lock')", "getElementById('pwd')", "__ENC__")
           if t in chk]
    ok = ('var DATA = {"updated"' in chk) and ('enterApp();' in chk) and not bad
    print("preview ->", out_path, "(%d bytes)" % len(chk))
    print("自检:", "OK" if ok else "FAILED -> 残留 %s" % bad)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(build(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PREVIEW))
