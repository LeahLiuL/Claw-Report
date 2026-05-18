import urllib.request, os, json, sys

api_key = os.environ.get("MATON_API_KEY", "")
if not api_key:
    print("ERROR: MATON_API_KEY not set")
    sys.exit(1)

url = "https://gateway.maton.ai/outlook/v1.0/me/messages"
url += "?$filter=contains(subject,'Alphaliner')"
url += "&$top=5"
url += "&$orderby=receivedDateTime desc"
url += "&$select=subject,from,receivedDateTime,body"

try:
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {api_key}")
    resp = urllib.request.urlopen(req, timeout=30)
    data = json.load(resp)
    
    messages = data.get("value", [])
    if not messages:
        print("NO_ALPHALINER_EMAIL_FOUND")
    else:
        for i, msg in enumerate(messages):
            print(f"EMAIL_{i}")
            print(f"SUBJECT: {msg['subject']}")
            sender = msg.get('from', {}).get('emailAddress', {}).get('address', 'N/A')
            print(f"FROM: {sender}")
            print(f"DATE: {msg.get('receivedDateTime', 'N/A')}")
            body = msg.get('body', {})
            print(f"BODY_TYPE: {body.get('contentType', 'N/A')}")
            content = body.get('content', '')
            print(f"BODY_LENGTH: {len(content)}")
            print(f"BODY_START: {content[:500]}")
            print(f"EMAIL_END")
except Exception as e:
    print(f"ERROR: {type(e).__name__}: {e}")
    sys.exit(1)
