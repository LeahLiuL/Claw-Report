import subprocess, os, sys

git = r"C:\Program Files\Git\cmd\git.exe"
repo = r"C:\Users\culadmin\Claw-Report"

def run(args, cwd=repo):
    result = subprocess.run([git] + args, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print(f"CMD: git {' '.join(args)}")
    if result.stdout.strip():
        print(f"STDOUT: {result.stdout.strip()}")
    if result.stderr.strip():
        print(f"STDERR: {result.stderr.strip()}")
    print(f"EXIT: {result.returncode}")
    return result

# Check git status
run(["status", "--short"])

# Add the new file
run(["add", "reports/alphaliner-intelligence-digest-wk19-2026.html"])

# Commit
r = run(["commit", "-m", "feat: Add Alphaliner Intelligence Digest Week 19, 2026 (May 4-10)"])
if r.returncode != 0 and "nothing to commit" in r.stdout + r.stderr:
    print("Already committed or nothing to commit")
elif r.returncode != 0:
    print(f"Commit failed with code {r.returncode}")
    sys.exit(1)

# Push to main
r = run(["push", "origin", "main"])
if r.returncode == 0:
    print("PUSH_SUCCESS")
else:
    print(f"PUSH_FAILED: {r.stderr}")
    sys.exit(1)
