"""Quick Outlook search for Alphaliner emails - minimal version."""
import win32com.client as win32
import pythoncom
import sys

pythoncom.CoInitialize()

try:
    outlook = win32.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace('MAPI')
    inbox = ns.GetDefaultFolder(6)  # olFolderInbox
    
    # Use Filter/Restrict to avoid scanning all messages
    items = inbox.Items
    items.Sort('[ReceivedTime]', True)
    
    # Restrict by subject
    filtered = items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%alphaliner%'")
    
    count = 0
    try:
        item = filtered.GetFirst()
        while item and count < 5:
            print(f"FOUND: {item.Subject} | {item.SenderEmailAddress} | {item.ReceivedTime}")
            count += 1
            item = filtered.GetNext()
    except:
        pass
    
    if count == 0:
        print("NO_RESULTS_FILTER1")
        # Try case-insensitive search differently
        filtered2 = items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%Alphaliner%'")
        item2 = filtered2.GetFirst()
        while item2 and count < 3:
            print(f"FOUND2: {item2.Subject} | {item2.SenderEmailAddress} | {item2.ReceivedTime}")
            count += 1
            item2 = filtered2.GetNext()
    
    if count == 0:
        print("TOTAL_NO_ALPHALINER_FOUND")

except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
