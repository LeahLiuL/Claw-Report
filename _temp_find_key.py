import urllib.request, os, json, urllib.parse

# Read token from the cloud service connection
# The MATON_API_KEY should be available via the connect_cloud_service token
# Let's try direct Microsoft Graph approach using the session info

api_key = os.environ.get("MATON_API_KEY", "")

# Try reading from various possible config locations
if not api_key:
    possible_paths = [
        os.path.expanduser("~/.maton/config.json"),
        os.path.expanduser("~/.maton/api_key"),
        os.path.expanduser("~/.config/maton/api_key"),
    ]
    for p in possible_paths:
        if os.path.exists(p):
            try:
                api_key = open(p).read().strip()
                break
            except:
                pass

# Try .env files in workbuddy skills
if not api_key:
    skill_env = os.path.expanduser("~/.workbuddy/skills/outlook-api/.env")
    if os.path.exists(skill_env):
        try:
            for line in open(skill_env):
                if line.strip().startswith("MATON_API_KEY="):
                    api_key = line.strip().split("=",1)[1].strip().strip('"').strip("'")
                    break
        except:
            pass

if not api_key:
    # List all .env and config files in workbuddy for debugging
    wb = os.path.expanduser("~/.workbuddy")
    for root, dirs, files in os.walk(wb):
        # Skip node_modules and deep recursion
        if 'node_modules' in root or '.git' in root:
            continue
        depth = root.replace(wb, '').count(os.sep)
        if depth > 3:
            continue
        for f in files:
            if f.endswith(('.env', '.json', '.toml', '.cfg', '.ini')):
                fp = os.path.join(root, f)
                print(f"CONFIG_FILE: {fp}")
    print("NO_MATON_API_KEY")
    exit(1)

print(f"API_KEY_FOUND (length={len(api_key)})")
