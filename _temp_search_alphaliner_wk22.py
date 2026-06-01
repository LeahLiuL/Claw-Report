"""Search Outlook for latest Alphaliner emails and extract content."""
import win32com.client as win32
import pythoncom
import json
import sys
import re
from datetime import datetime

pythoncom.CoInitialize()

outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder(6)  # olFolderInbox

items = inbox.Items
items.Sort('[ReceivedTime]', True)

found = []
for i in range(1, min(items.Count + 1, 500)):
    try:
        item = items.Item(i)
        subject = getattr(item, 'Subject', '')
        sender = ''
        if hasattr(item, 'Sender'):
            sender = getattr(item.Sender, 'EmailAddress', '') or getattr(item.Sender, 'Name', '')
        elif hasattr(item, 'SenderEmailAddress'):
            sender = item.SenderEmailAddress
        
        subject_lower = subject.lower()
        sender_lower = sender.lower()
        
        if 'alphaliner' in subject_lower or 'alphaliner' in sender_lower:
            received = ''
            if hasattr(item, 'ReceivedTime'):
                received = item.ReceivedTime.strftime('%Y-%m-%d %H:%M') if item.ReceivedTime else 'N/A'
            
            found.append({
                'subject': subject,
                'sender': sender,
                'received': received,
                'entryid': item.EntryID
            })
            if len(found) >= 5:
                break
    except Exception as e:
        continue

if not found:
    print('NO_ALPHALINER_EMAIL_FOUND')
    sys.exit(0)

# Print found emails summary
for idx, f in enumerate(found):
    print(f"[{idx}] SUBJECT: {f['subject'][:120]}")
    print(f"    SENDER: {f['sender']}")
    print(f"    RECEIVED: {f['received']}")
    print()

# Now extract the body of the latest email
latest = found[0]
print("=" * 60)
print("LATEST_EMAIL_BODY_START")
print("=" * 60)

# Use EntryID to get the item again
item = outlook.GetItemFromID(latest['entryid'])

# Try HTML body first
html_body = getattr(item, 'HTMLBody', '')
text_body = getattr(item, 'Body', '')

if html_body:
    print("CONTENT_TYPE: HTML")
    # Save HTML body to temp file for processing
    with open(r'C:\Users\culadmin\Claw-Report\_temp_alphaliner_body.html', 'w', encoding='utf-8') as f:
        f.write(html_body)
    print("HTML_SAVED_TO: _temp_alphaliner_body.html")
    # Also save text version
    with open(r'C:\Users\culadmin\Claw-Report\_temp_alphaliner_body.txt', 'w', encoding='utf-8') as f:
        f.write(text_body)
    print("TEXT_SAVED_TO: _temp_alphaliner_body.txt")
elif text_body:
    print("CONTENT_TYPE: TEXT")
    with open(r'C:\Users\culadmin\Claw-Report\_temp_alphaliner_body.txt', 'w', encoding='utf-8') as f:
        f.write(text_body)
    print("TEXT_SAVED_TO: _temp_alphaliner_body.txt")
else:
    print("NO_BODY_CONTENT")

print("=" * 60)
print("LATEST_EMAIL_BODY_END")
print("=" * 60)
