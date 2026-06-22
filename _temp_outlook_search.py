import urllib.request, json, sys

API_KEY = "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICJteWZFenA3ODNLaV9KQ3g4Vm5jM1hfaXg2alpyYjZDZjVPTWtHWk1QSTNzIn0.eyJleHAiOjE4MDgzNzY2MTAsImlhdCI6MTc4MTY2MzMxMiwiYXV0aF90aW1lIjoxNzc2ODQwNjA5LCJqdGkiOiJjMTczNTZhMi1jMmRiLTQ0YWEtYjEyYy1kZTQ5YmM1YThkMmUiLCJpc3MiOiJodHRwczovL3d3dy5jb2RlYnVkZHkuY24vYXV0aC9yZWFsbXMvY29waWxvdCIsImF1ZCI6ImFjY291bnQiLCJzdWIiOiIxNjBiOTgxOS0wNjk2LTQ2NjUtYTEzZC1kZGNmMzg5NDRkODEiLCJ0eXAiOiJCZWFyZXIiLCJhenAiOiJjb25zb2xlIiwic2lkIjoiNzY1Yjk0ZDAtZjgwYy00ODJmLThiZDItYjlkYmZlYzVjZDY5IiwiYWNyIjoiMCIsImFsbG93ZWQtb3JpZ2lucyI6WyIqIl0sInJlYWxtX2FjY2VzcyI6eyJyb2xlcyI6WyJkZWZhdWx0LXJvbGVzIiwib2ZmbGluZV9hY2Nlc3MiLCJ1bWFfYXV0aG9yaXphdGlvbiJdfSwicmVzb3VyY2VfYWNjZXNzIjp7ImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgb2ZmbGluZV9hY2Nlc3MgZW1haWwiLCJlbWFpbF92ZXJpZmllZCI6ZmFsc2UsIm5pY2tuYW1lIjoiTGVhaCIsInByZWZlcnJlZF91c2VybmFtZSI6IjE4OTM1MDcxNTg4In0.Ih6blezYlvA8Ft8vrMLAw_MMwGZpO36tdjYAGdZukCIaJbCooiv3P7eti87SJOq2puWqOA9tSOuKCjlHSXbRgRVB1ekOrC_eCRankpNN7ygahgDo_Mmv2tXN7QKrfWbXtnLUkDRUZy75017eUf-6CS2MPBsoOtzmi4tlnfL4LAHdd65XXE6CdYQQhgFuH3QJV1d9CsuGFejwp8A_JrG_naLL-v-F7m8tc5MnUHmAKExE5LrYlDZ5MxphEvqQmczXZ864dJkaeBwm4jx6J_TY674ruIUlruUtEWoKQiOyHZeZEpvnNoyaMdvZB3fvsH0hDHeK_me6eew83JX4IKxbow"

def api_get(path):
    url = f"https://gateway.maton.ai/outlook/{path}"
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {API_KEY}")
    try:
        resp = urllib.request.urlopen(req, timeout=30)
        return json.load(resp)
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"HTTP ERROR {e.code}: {body}", file=sys.stderr)
        return None
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return None

# Step 1: Check connections
print("=== Checking connections ===")
req = urllib.request.Request("https://ctrl.maton.ai/connections?app=outlook&status=ACTIVE")
req.add_header("Authorization", f"Bearer {API_KEY}")
try:
    resp = urllib.request.urlopen(req, timeout=30)
    data = json.load(resp)
    print(json.dumps(data, indent=2, ensure_ascii=False))
except urllib.error.HTTPError as e:
    body = e.read().decode()
    print(f"Connection check failed {e.code}: {body}")
except Exception as e:
    print(f"Connection check error: {e}")
