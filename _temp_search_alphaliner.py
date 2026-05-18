import urllib.request
import os
import json

api_key = os.environ.get('MATON_API_KEY', '')
url = 'https://gateway.maton.ai/outlook/v1.0/me/messages?$search=%22Alphaliner%22&$top=10&$orderby=receivedDateTime%20desc'
req = urllib.request.Request(url)
req.add_header('Authorization', 'Bearer ' + api_key)

try:
    resp = urllib.request.urlopen(req)
    data = json.loads(resp.read())
    messages = data.get('value', [])
    print(f"Found {len(messages)} messages")
    for m in messages:
        sender = m.get('from', {}).get('emailAddress', {}).get('address', 'N/A')
        subject = m.get('subject', 'N/A')
        date = m.get('receivedDateTime', 'N/A')
        msg_id = m.get('id', 'N/A')
        print(f"ID: {msg_id[:30]}")
        print(f"  From: {sender}")
        print(f"  Subject: {subject}")
        print(f"  Date: {date}")
        print()
except Exception as e:
    print(f"Error: {e}")
