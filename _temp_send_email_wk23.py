import urllib.request
import urllib.parse
import os
import json

API_KEY = "am_us_fe0842c00141bdc1b872d0b9a1dda62163b350b6d3bb522e7cf3ba806237812b"

GITHUB_URL = "https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk23-2026.html"

subject = "[Claw] Alphaliner Intelligence Digest - Week 23 2026"
to_email = "leahliu@culines.com"

body_html = f"""
<html>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #0d1117; color: #e6edf3; padding: 20px; margin: 0;">
  <div style="max-width: 600px; margin: 0 auto; background: #161b22; border: 1px solid #21262d; border-radius: 12px; padding: 24px;">
    
    <div style="margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid #21262d;">
      <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
        <span style="background: linear-gradient(135deg, #1f6feb, #58a6ff); border-radius: 8px; padding: 8px 12px; font-size: 18px;">⛲</span>
        <div>
          <div style="font-size: 16px; font-weight: 700;">CUL — Alphaliner Intelligence Digest</div>
          <div style="font-size: 12px; color: #8b949e;">Week 23, 2026 (Jun 1–7) · Powered by Alphaliner + AI</div>
        </div>
      </div>
    </div>

    <h2 style="color: #58a6ff; font-size: 18px; margin-bottom: 12px;">Week 23 Key Highlights</h2>
    
    <div style="background: rgba(248,81,73,0.08); border: 1px solid rgba(248,81,73,0.25); border-radius: 8px; padding: 12px; margin-bottom: 12px;">
      <div style="font-weight: 700; color: #f85149; margin-bottom: 6px;">🚨 HORMUZ: Iran Declares Full IRGC Control</div>
      <div style="font-size: 13px; color: #8b949e;">May 30 IRGC announcement: All vessels require advance clearance. ADNOC confirms full normalization not before H1 2027. US-Iran ceasefire extended 60 days → Brent drops to ~$91/bbl (largest monthly decline since 2020).</div>
    </div>

    <div style="background: rgba(63,185,80,0.06); border: 1px solid rgba(63,185,80,0.2); border-radius: 8px; padding: 12px; margin-bottom: 12px;">
      <div style="font-weight: 700; color: #3fb950; margin-bottom: 6px;">📈 MARKET: WCI $2,800 (+3%) — Peak Season Erupts Early</div>
      <div style="font-size: 13px; color: #8b949e;">4th consecutive weekly rise. Shanghai-Rotterdam +15% to $2,773/FEU. Shanghai-Genoa +10% to $4,082/FEU. Persian Gulf +3.6% to $4,462/TEU. Early peak season 1-2 months ahead of historical norms. Multiple lanes "fully booked" for June.</div>
    </div>

    <div style="background: rgba(210,153,34,0.06); border: 1px solid rgba(210,153,34,0.25); border-radius: 8px; padding: 12px; margin-bottom: 12px;">
      <div style="font-weight: 700; color: #d29922; margin-bottom: 6px;">⚓ CHARTER: Bulk Carriers Being Converted to Boxships</div>
      <div style="font-size: 13px; color: #8b949e;">First bulk-to-container conversions since 2021 COVID surge. Global fleet hits 34M TEU. Zero spot availability above 3,000 TEU. 47 blank sailings announced W23-27 (only 4 on Asia-Europe — bullish signal).</div>
    </div>

    <div style="background: rgba(88,166,255,0.06); border: 1px solid rgba(88,166,255,0.2); border-radius: 8px; padding: 12px; margin-bottom: 16px;">
      <div style="font-weight: 700; color: #58a6ff; margin-bottom: 6px;">⛽ BUNKER: Singapore VLSFO $782/MT — Easing Trend</div>
      <div style="font-size: 13px; color: #8b949e;">Singapore VLSFO $782/MT (May 25) vs $860 (May 26 spike). Brent at $91/bbl suggests further VLSFO correction possible. Rotterdam remains globally lowest. Monitor Brent daily for ceasefire-driven relief.</div>
    </div>

    <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 20px; background: #21262d; border-radius: 8px; overflow: hidden;">
      <thead>
        <tr style="background: #30363d;">
          <th style="padding: 10px 14px; text-align: left; color: #8b949e; font-size: 11px; text-transform: uppercase;">Indicator</th>
          <th style="padding: 10px 14px; text-align: left; color: #8b949e; font-size: 11px; text-transform: uppercase;">Value</th>
          <th style="padding: 10px 14px; text-align: left; color: #8b949e; font-size: 11px; text-transform: uppercase;">Change</th>
        </tr>
      </thead>
      <tbody>
        <tr style="border-top: 1px solid #30363d;">
          <td style="padding: 9px 14px; color: #e6edf3;">Drewry WCI Composite</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">$2,800/FEU</td>
          <td style="padding: 9px 14px; color: #f85149; font-weight: 600;">▲ +3% WoW</td>
        </tr>
        <tr style="border-top: 1px solid #30363d; background: rgba(255,255,255,0.02);">
          <td style="padding: 9px 14px; color: #e6edf3;">Shanghai → Rotterdam</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">$2,773/FEU</td>
          <td style="padding: 9px 14px; color: #f85149; font-weight: 600;">▲ +15% WoW</td>
        </tr>
        <tr style="border-top: 1px solid #30363d;">
          <td style="padding: 9px 14px; color: #e6edf3;">Shanghai → Genoa</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">$4,082/FEU</td>
          <td style="padding: 9px 14px; color: #f85149; font-weight: 600;">▲ +10% WoW</td>
        </tr>
        <tr style="border-top: 1px solid #30363d; background: rgba(255,255,255,0.02);">
          <td style="padding: 9px 14px; color: #e6edf3;">Shanghai → Persian Gulf</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">$4,462/TEU</td>
          <td style="padding: 9px 14px; color: #f85149; font-weight: 600;">▲ +3.6% WoW</td>
        </tr>
        <tr style="border-top: 1px solid #30363d;">
          <td style="padding: 9px 14px; color: #e6edf3;">Singapore VLSFO</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">$782/MT</td>
          <td style="padding: 9px 14px; color: #3fb950; font-weight: 600;">▼ Easing</td>
        </tr>
        <tr style="border-top: 1px solid #30363d; background: rgba(255,255,255,0.02);">
          <td style="padding: 9px 14px; color: #e6edf3;">Brent Crude</td>
          <td style="padding: 9px 14px; font-weight: 700; color: #e6edf3;">~$91/bbl</td>
          <td style="padding: 9px 14px; color: #3fb950; font-weight: 600;">▼ -17% MTD</td>
        </tr>
      </tbody>
    </table>

    <div style="text-align: center; margin-bottom: 16px;">
      <a href="{GITHUB_URL}" style="background: #58a6ff; color: #0d1117; padding: 10px 24px; border-radius: 8px; font-weight: 700; font-size: 14px; text-decoration: none; display: inline-block;">📊 View Full Report — Week 23 →</a>
    </div>

    <div style="font-size: 11px; color: #8b949e; text-align: center; padding-top: 14px; border-top: 1px solid #21262d;">
      For internal use only · AI-assisted analysis based on public market data · AI解读仅供参考 · Not investment advice<br>
      Generated by Claw Automation · Week 23, 2026
    </div>
  </div>
</body>
</html>
"""

payload = {
    "message": {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body_html
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": to_email
                }
            }
        ]
    },
    "saveToSentItems": True
}

data = json.dumps(payload).encode('utf-8')
req = urllib.request.Request(
    "https://gateway.maton.ai/outlook/v1.0/me/sendMail",
    data=data,
    method="POST"
)
req.add_header("Authorization", f"Bearer {API_KEY}")
req.add_header("Content-Type", "application/json")

try:
    resp = urllib.request.urlopen(req)
    print(f"Success! Status: {resp.status}")
    print(f"Response: {resp.read().decode()}")
except urllib.error.HTTPError as e:
    print(f"HTTP Error {e.code}: {e.read().decode()}")
except Exception as ex:
    print(f"Error: {ex}")
