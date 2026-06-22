import win32com.client
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
messages.Sort('[ReceivedTime]', True)

# Search for Alphaliner/internal circulation emails
print("=== Search: Alphaliner / internal circulation / Mabel ===")
found = 0
for i, msg in enumerate(messages):
    if i > 500:
        break
    try:
        subject = (msg.Subject or "").lower()
        sender = (msg.SenderName or "").lower()
        
        if any(kw in subject for kw in ["internal", "circulation", "newsletter", "alphaliner", "mabel"]):
            found += 1
            received = msg.ReceivedTime.strftime("%Y-%m-%d %H:%M")
            print("[{}] {}".format(received, msg.Subject[:120]))
            print("  From: {}".format(msg.SenderName[:80]))
            print()
    except:
        pass

print("Found: {}".format(found))

# Search for recent emails with PDF attachments (potential Alphaliner newsletters)
print("\n=== Recent emails with PDF attachments (last 14 days) ===")
from datetime import datetime, timedelta
cutoff = datetime.now() - timedelta(days=14)
found2 = 0
for i, msg in enumerate(messages):
    if i > 300:
        break
    try:
        received = msg.ReceivedTime
        if received.replace(tzinfo=None) < cutoff:
            continue
        if msg.Attachments.Count > 0:
            att_names = []
            for j in range(1, msg.Attachments.Count + 1):
                fname = msg.Attachments[j].FileName or ""
                if ".pdf" in fname.lower():
                    att_names.append(fname)
            if att_names:
                found2 += 1
                print("[{}] {}".format(received.strftime("%Y-%m-%d %H:%M"), msg.Subject[:100]))
                print("  From: {}".format((msg.SenderName or "")[:60]))
                print("  PDFs: {}".format(att_names))
                print()
    except:
        pass

print("Found PDF emails: {}".format(found2))
