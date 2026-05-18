import urllib.request
import urllib.error
import json
import os

# Try to get Maton API key from config
config_path = os.path.expanduser('~/.config/imap-smtp-email/.env')
maton_key = ''

if os.path.exists(config_path):
    with open(config_path, 'r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('MATON_API_KEY='):
                maton_key = line.split('=', 1)[1].strip()
                break

if not maton_key:
    maton_key = os.environ.get('MATON_API_KEY', '')

if not maton_key:
    print("ERROR: MATON_API_KEY not found")
    print("Trying to list connections to check...")
    exit(1)

print(f"Using API key: {maton_key[:10]}...")

# Send email via Outlook API
report_url = "https://leahliul.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk20-2026.html"

email_body = f"""<html>
<head><meta charset='UTF-8'></head>
<body style='font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Arial, sans-serif; background: #f5f5f5; padding: 20px;'>
  <div style='max-width: 680px; margin: 0 auto; background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,0.08);'>
    <!-- Header -->
    <div style='background: linear-gradient(135deg, #0d1117, #1f2937); padding: 28px 32px;'>
      <div style='display: flex; align-items: center; gap: 12px;'>
        <div style='background: linear-gradient(135deg, #1f6feb, #58a6ff); border-radius: 8px; width: 40px; height: 40px; display: flex; align-items: center; justify-content: center; font-size: 20px; color: white; flex-shrink: 0;'>⛲</div>
        <div>
          <h1 style='margin: 0; font-size: 18px; font-weight: 700; color: #e6edf3;'>CUL — Alphaliner Intelligence Digest</h1>
          <p style='margin: 4px 0 0; font-size: 12px; color: #8b949e;'>Week 20, 2026 (May 12–18) · Powered by Alphaliner</p>
        </div>
      </div>
    </div>

    <!-- Content -->
    <div style='padding: 28px 32px;'>
      <h2 style='margin: 0 0 16px; font-size: 16px; color: #1f2937;'>📊 本周航运情报摘要</h2>

      <!-- Key Stats -->
      <div style='display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 24px;'>
        <div style='background: #f0f7ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 12px 18px; flex: 1; min-width: 140px;'>
          <div style='font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;'>WCI 综合指数</div>
          <div style='font-size: 22px; font-weight: 800; color: #1d4ed8;'>$2,553</div>
          <div style='font-size: 12px; color: #dc2626; font-weight: 600;'>▲ +12% 周环比</div>
        </div>
        <div style='background: #fef9ee; border: 1px solid #fde68a; border-radius: 8px; padding: 12px 18px; flex: 1; min-width: 140px;'>
          <div style='font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;'>新加坡 VLSFO</div>
          <div style='font-size: 22px; font-weight: 800; color: #92400e;'>$843/MT</div>
          <div style='font-size: 12px; color: #dc2626; font-weight: 600;'>▲ +4.2% 危机新高</div>
        </div>
        <div style='background: #fff1f2; border: 1px solid #fecdd3; border-radius: 8px; padding: 12px 18px; flex: 1; min-width: 140px;'>
          <div style='font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;'>霍尔木兹封锁</div>
          <div style='font-size: 22px; font-weight: 800; color: #991b1b;'>第12周</div>
          <div style='font-size: 12px; color: #dc2626; font-weight: 600;'>1,500艘船受困</div>
        </div>
      </div>

      <!-- Top Stories -->
      <h3 style='margin: 0 0 12px; font-size: 14px; color: #374151; border-left: 3px solid #3b82f6; padding-left: 10px;'>🔥 本周重点</h3>
      <table style='width: 100%; border-collapse: collapse; margin-bottom: 20px;'>
        <tr>
          <td style='padding: 10px 0; border-bottom: 1px solid #f3f4f6; vertical-align: top;'>
            <span style='background: #fef3c7; color: #92400e; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-right: 8px;'>⭐ 重点</span>
            <span style='font-size: 13px; color: #1f2937; font-weight: 600;'>Drewry WCI飙升12% — 各大船公司全面叠加峰季附加费</span>
            <p style='margin: 4px 0 0 40px; font-size: 12px; color: #6b7280;'>WCI至$2,553/40ft，CMA CGM推出$2,000/FEU PSS，MSC/Maersk/赫伯罗特同步EFS上调，部分航线PSS+EFS叠加达$7,200/FEU。</p>
          </td>
        </tr>
        <tr>
          <td style='padding: 10px 0; border-bottom: 1px solid #f3f4f6; vertical-align: top;'>
            <span style='background: #fef3c7; color: #92400e; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-right: 8px;'>⭐ 重点</span>
            <span style='font-size: 13px; color: #1f2937; font-weight: 600;'>霍尔木兹危机第12周 — 美军护航仅2艘商船通过</span>
            <p style='margin: 4px 0 0 40px; font-size: 12px; color: #6b7280;'>美军5月4日击沉6艘伊朗小艇并开辟护航通道，但停火仍脆弱，伊朗袭击UAE，1,500艘船仍受困。</p>
          </td>
        </tr>
        <tr>
          <td style='padding: 10px 0; border-bottom: 1px solid #f3f4f6; vertical-align: top;'>
            <span style='background: #fee2e2; color: #991b1b; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-right: 8px;'>⚡ CUL直接</span>
            <span style='font-size: 13px; color: #1f2937; font-weight: 600;'>青岛港拥堵达4天 — 华北港口峰季提前爆发</span>
            <p style='margin: 4px 0 0 40px; font-size: 12px; color: #6b7280;'>托运人在关税不确定性下抢运货物，青岛平均靠泊延误4天，影响班期准点率。</p>
          </td>
        </tr>
        <tr>
          <td style='padding: 10px 0; vertical-align: top;'>
            <span style='background: #dcfce7; color: #166534; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-right: 8px;'>市场</span>
            <span style='font-size: 13px; color: #1f2937; font-weight: 600;'>马士基/赫伯罗特维持好望角绕行 — 亚欧航次延长14天</span>
            <p style='margin: 4px 0 0 40px; font-size: 12px; color: #6b7280;'>三大联盟确认好望角绕行方案维持，苏伊士/曼德海峡暂不恢复，影响全球运力吸收。</p>
          </td>
        </tr>
      </table>

      <!-- View Report Button -->
      <div style='text-align: center; margin: 24px 0;'>
        <a href='{report_url}' style='background: linear-gradient(135deg, #1f6feb, #58a6ff); color: white; text-decoration: none; padding: 14px 36px; border-radius: 8px; font-size: 14px; font-weight: 700; display: inline-block;'>
          📊 查看完整 Week 20 报告 →
        </a>
      </div>
      <p style='text-align: center; font-size: 11px; color: #9ca3af; margin-top: 8px;'>{report_url}</p>

    </div>

    <!-- Footer -->
    <div style='background: #f9fafb; border-top: 1px solid #e5e7eb; padding: 16px 32px; text-align: center;'>
      <p style='margin: 0; font-size: 11px; color: #9ca3af;'>CUL — Alphaliner Intelligence Digest | Week 20, 2026 (May 12–18)</p>
      <p style='margin: 4px 0 0; font-size: 11px; color: #9ca3af;'>For internal use only · AI-assisted analysis · ⚠️ AI解读仅供参考</p>
    </div>
  </div>
</body>
</html>"""

payload = {
    "message": {
        "subject": "[Claw] Alphaliner Intelligence Digest - Week 20 2026",
        "body": {
            "contentType": "HTML",
            "content": email_body
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": "leahliu@culines.com"
                }
            }
        ]
    },
    "saveToSentItems": True
}

data = json.dumps(payload).encode('utf-8')
url = 'https://gateway.maton.ai/outlook/v1.0/me/sendMail'
req = urllib.request.Request(url, data=data, method='POST')
req.add_header('Authorization', 'Bearer ' + maton_key)
req.add_header('Content-Type', 'application/json')

try:
    resp = urllib.request.urlopen(req)
    print(f"Email sent successfully! Status: {resp.status}")
except urllib.error.HTTPError as e:
    body = e.read()
    print(f"HTTP Error {e.code}: {e.reason}")
    print(f"Response: {body.decode('utf-8', errors='ignore')}")
except Exception as e:
    print(f"Error: {e}")
