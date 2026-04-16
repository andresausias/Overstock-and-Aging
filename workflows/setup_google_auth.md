# Workflow: Setup Google OAuth

## Objective
One-time setup to authorize the automation to access Google Drive and Slides APIs.

## Steps

### 1. Create a Google Cloud Project
1. Go to https://console.cloud.google.com/
2. Create a new project: **Overstock-Aging-Automation**
3. Enable these APIs:
   - Google Drive API
   - Google Slides API

### 2. Create OAuth Credentials
1. Go to **APIs & Services → Credentials**
2. Click **Create Credentials → OAuth client ID**
3. Application type: **Desktop app**
4. Name: `overstock-aging-local`
5. Download the JSON file
6. Rename it to `credentials.json`
7. Place it in the project root (same level as `CLAUDE.md`)

### 3. Run the OAuth Flow
```bash
cd /path/to/Overstock-and-Aging
python -c "from tools.auth_google import get_drive_service; get_drive_service()"
```
A browser window will open — log in and grant permissions.
`token.json` is saved automatically. You will not need to re-authenticate unless the token expires.

### 4. Verify
```bash
python -c "
from tools.auth_google import get_drive_service
svc = get_drive_service()
me = svc.about().get(fields='user').execute()
print('Authenticated as:', me['user']['emailAddress'])
"
```

## Notes
- `credentials.json` and `token.json` are gitignored — never commit them.
- If authentication fails later, delete `token.json` and re-run step 3.
- Scopes used: `drive` (full access) + `presentations` (read/write slides).
