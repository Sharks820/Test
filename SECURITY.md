# Security Guidelines

## Keep This Repository Private

1. Go to **GitHub repo Settings > General** (scroll to bottom) > **Danger Zone**
2. Click **Change visibility** and set to **Private**
3. Only invite collaborators you trust

## What NOT to Commit

Never commit any of the following — they are blocked by `.gitignore` but stay vigilant:

- Passwords or API keys
- `.env` files with secrets
- Private keys (`.pem`, `.key`, `.p12`)
- OAuth tokens or service account JSON files
- Personal data (SSNs, financial records, etc.)

## If You Accidentally Commit a Secret

1. **Rotate the secret immediately** (change the password/key)
2. Remove it from history:
   ```bash
   git filter-branch --force --index-filter \
     "git rm --cached --ignore-unmatch PATH_TO_FILE" \
     --prune-empty --tag-name-filter cat -- --all
   git push --force
   ```
3. Consider the old secret **permanently compromised**

## Branch Isolation

Each project branch is an orphan branch with no shared history. This means:
- Switching branches completely replaces your working directory
- One project's files never appear in another project's branch
- Deleting a project branch has zero effect on other projects
