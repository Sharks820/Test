# Sharks820 - Multi-Project Repository

This repository hosts **multiple independent projects**, each on its own branch. Projects do not share files or interfere with each other.

---

## Branch Structure

| Branch | Purpose |
|--------|---------|
| `main` | Base branch — contains only repo config files (.gitignore, .gitattributes, README) |
| `project/<name>` | Each project gets its own branch off `main` |

### Current Projects

| Project | Branch |
|---------|--------|
| Noah Kahn Ticket Tracker | `main` (legacy — existed before branching setup) |

### How to Start a New Project

1. **Create a new branch** from `main`:
   ```bash
   git checkout main
   git pull origin main
   git checkout -b project/my-new-project
   ```

2. **Add your project files** — they only exist on this branch.

3. **Push the branch**:
   ```bash
   git push -u origin project/my-new-project
   ```

4. Each project branch is **completely independent**. Changes on one branch do not affect others.

---

## Security Measures

This repo includes several layers of protection:

### 1. `.gitignore` — Blocks Sensitive Files
The following are **automatically excluded** from all commits:
- `.env` files (API keys, secrets, config)
- `*.key`, `*.pem`, `*.secret`, `*.credentials` files
- `private/` and `confidential/` directories
- OS junk files (`.DS_Store`, `Thumbs.db`)

### 2. `.gitattributes` — Protects Binary & Sensitive Files
- Binary files (Excel, images, PDFs) are handled correctly across branches
- Sensitive file types are marked as binary so they never appear in diffs/logs

### 3. Pre-Commit Hook — Catches Accidental Secret Leaks
A git hook scans every commit for:
- API key patterns (AWS, GitHub tokens, etc.)
- Password/secret assignments in code
- Blocked file extensions (`.env`, `.key`, `.pem`)

If detected, the commit is **blocked** with a warning.

### 4. Repository Visibility
- Ensure the GitHub repo is set to **Private** in Settings > General > Danger Zone
- Only collaborators you explicitly invite can see the code

---

## Rules for Keeping Branches Clean

1. **Never merge project branches into each other** — they are independent
2. **Only merge INTO `main`** if you want to archive/publish a project
3. **Keep secrets in `.env` files** — they are auto-ignored
4. **Use `private/` folder** for any data files you want git to ignore
5. **Each branch = one project** — do not mix projects on a single branch

---

## Quick Reference

```bash
# List all project branches
git branch -a

# Switch to a project
git checkout project/<name>

# Create a new project
git checkout main && git checkout -b project/<new-name>

# Check what's being tracked (security audit)
git ls-files

# Verify .gitignore is working
git status --ignored
```
