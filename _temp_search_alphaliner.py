import urllib.request, os, json, urllib.parse, pathlib

# Try to get API key
api_key = os.environ.get("MATON_API_KEY", "")
if not api_key:
    try:
        cfg_path = os.path.expanduser("~/.workbuddy/mcp.json")
        if os.path.exists(cfg_path):
            cfg = json.loads(pathlib.Path(cfg_path).read_text())
            for s in cfg.get("mcpServers", {}).values():
                env = s.get("env", {})
                if "MATON_API_KEY" in env:
                    api_key = env["MATON_API_KEY"]
                    break
    except Exception as e:
        print(f"Config read error: {e}")

if not api_key:
    print("ERROR: No MATON_API_KEY found in env or mcp.json")
    exit(1)

# Search for Alphaliner emails in Inbox
filter_q = urllib.parse.quote("contains(subject,'Alphaliner')")
url = f"https://gateway.maton.ai/outlook/v1.0/me/mailFolders/Inbox/messages?$filter={filter_q}&$orderby=receivedDateTime desc&$top=5&$select=subject,from,receivedDateTime,bodyPreview,id"

req = urllib.request.Request(url)
req.add_header("Authorization", f"Bearer {api_key}")

try:
    resp = urllib.request.urlopen(req)
    data = json.loads(resp.read().decode())
    messages = data.get("value", [])
    if not messages:
        print("NO_MESSAGES_FOUND")
    else:
        for m in messages:
            print(f"ID: {m['id']}")
            print(f"Subject: {m['subject']}")
            frm = m['from']['emailAddress']['address']
            print(f"From: {frm}")
            print(f"Date: {m['receivedDateTime']}")
            preview = m.get('bodyPreview', '')[:300]
            print(f"Preview: {preview}")
            print("---SEPARATOR---")
except urllib.error.HTTPError as e:
    body = e.read().decode()
    print(f"HTTP_ERROR {e.code}: {body}")
except Exception as e:
    print(f"ERROR: {e}")
