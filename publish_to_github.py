import ctypes
from ctypes import wintypes
import urllib.request
import urllib.error
import json
import subprocess
import os

class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ('Flags', wintypes.DWORD),
        ('Type', wintypes.DWORD),
        ('TargetName', wintypes.LPWSTR),
        ('Comment', wintypes.LPWSTR),
        ('LastWritten', wintypes.FILETIME),
        ('CredentialBlobSize', wintypes.DWORD),
        ('CredentialBlob', ctypes.POINTER(ctypes.c_byte)),
        ('Persist', wintypes.DWORD),
        ('AttributeCount', wintypes.DWORD),
        ('Attributes', ctypes.c_void_p),
        ('TargetAlias', wintypes.LPWSTR),
        ('UserName', wintypes.LPWSTR),
    ]

def get_github_token():
    advapi32 = ctypes.windll.advapi32
    CredRead = advapi32.CredReadW
    CredRead.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
    CredRead.restype = wintypes.BOOL
    CredFree = advapi32.CredFree
    CredFree.argtypes = [ctypes.c_void_p]

    pcred = ctypes.POINTER(CREDENTIAL)()
    target = 'GitHub - https://api.github.com/saisathiya77'
    if CredRead(target, 1, 0, ctypes.byref(pcred)):
        blob = bytes(pcred.contents.CredentialBlob[:pcred.contents.CredentialBlobSize])
        token = blob.decode('utf-8')
        CredFree(pcred)
        return token
    raise RuntimeError("Could not find GitHub credentials in Windows Credential Manager.")

def create_or_get_repo(token, repo_name):
    url = "https://api.github.com/user/repos"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Project-Setup-Script"
    }
    payload = {
        "name": repo_name,
        "description": "Comprehensive AI Project Plan and Strategic Analysis Report: Autonomous Multi-Agent Swarms, Multimodal Perception, and Predictive Supply Chain Intelligence",
        "private": False,
        "has_issues": True,
        "has_projects": True,
        "has_wiki": False
    }
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req) as resp:
            repo_info = json.loads(resp.read().decode('utf-8'))
            print("Successfully created repository on GitHub!")
            return repo_info["html_url"], repo_info["clone_url"]
    except urllib.error.HTTPError as e:
        err_body = e.read().decode('utf-8')
        if e.code == 422: # Already exists
            print(f"Repository '{repo_name}' already exists. Retrieving details...")
            get_req = urllib.request.Request(f"https://api.github.com/repos/saisathiya77/{repo_name}", headers=headers)
            with urllib.request.urlopen(get_req) as resp:
                repo_info = json.loads(resp.read().decode('utf-8'))
                return repo_info["html_url"], repo_info["clone_url"]
        else:
            raise RuntimeError(f"GitHub API Error ({e.code}): {err_body}")

def push_to_github(token, repo_name):
    git_bin = r"C:\Users\Saisa\AppData\Local\GitHubDesktop\app-3.5.5\resources\app\git\cmd\git.exe"
    html_url, clone_url = create_or_get_repo(token, repo_name)
    print(f"Repository Web URL: {html_url}")
    
    auth_remote_url = f"https://{token}@github.com/saisathiya77/{repo_name}.git"
    
    # Configure git remote
    subprocess.run([git_bin, "remote", "remove", "origin"], stderr=subprocess.DEVNULL)
    subprocess.run([git_bin, "remote", "add", "origin", auth_remote_url], check=True)
    
    # Push branch main
    print("Pushing commits to GitHub main branch...")
    res = subprocess.run([git_bin, "push", "-u", "origin", "main", "--force"], capture_output=True, text=True)
    print("Git Push stdout:", res.stdout)
    if res.returncode != 0:
        print("Git Push stderr:", res.stderr)
        raise RuntimeError("Git push failed.")
    
    # Reset remote URL to clean public HTTPS URL (without embedded token for security)
    clean_url = f"https://github.com/saisathiya77/{repo_name}.git"
    subprocess.run([git_bin, "remote", "set-url", "origin", clean_url], check=True)
    
    print("\nSUCCESS!")
    print(f"Project is now live on GitHub at: {html_url}")
    return html_url

if __name__ == "__main__":
    token = get_github_token()
    repo_name = "enterprise-ai-project-plan-and-strategic-analysis"
    project_url = push_to_github(token, repo_name)
