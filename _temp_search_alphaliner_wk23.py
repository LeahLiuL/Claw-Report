import urllib.request
import urllib.parse
import os
import json

API_KEY = "am_us_fe0842c00141bdc1b872d0b9a1dda62163b350b6d3bb522e7cf3ba806237812b"

# Search for Alphaliner emails - use $search instead of $filter for broader match
url = "https://gateway.maton.ai/outlook/v1.0/me/messages?$search=%22alphaliner%22&$top=10&$orderby=receivedDateTime+desc&$select=id,subject,from,receivedDateTime,bodyPreview,body"

req = urllib.request.Request(url)
req.add_header("Authorization", f"Bearer {API_KEY}")
req.add_header("ConsistencyLevel", "eventual")

try:
    resp = urllib.request.urlopen(req)
    data = json.load(resp)
    print(json.dumps(data, indent=2, ensure_ascii=False))
except urllib.error.HTTPError as e:
    print(f"HTTP Error {e.code}: {e.read().decode()}")
except Exception as ex:
    print(f"Error: {ex}")
