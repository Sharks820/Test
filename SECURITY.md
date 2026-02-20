# Security Policy

## Sensitive Data Handling

**NEVER commit the following to this repository:**
- Passwords, API keys, tokens, or credentials
- `.env` files or environment configuration with secrets
- Private keys (`.pem`, `.key`, `.p12`)
- Personal data (SSNs, credit card numbers, etc.)

## If You Accidentally Commit Secrets

1. **Do NOT just delete the file** — it remains in git history
2. Rotate/revoke the exposed credential immediately
3. Use `git filter-branch` or [BFG Repo-Cleaner](https://rtyley.github.io/bfg-repo-cleaner/) to purge from history
4. Force push the cleaned history

## Repository Access

- This repo **must** remain **Private** on GitHub
- Review collaborator access regularly in Settings > Collaborators
- Remove access for anyone who no longer needs it

## Reporting

If you discover exposed credentials or a security issue in this repo, change the credentials immediately and clean the git history.
