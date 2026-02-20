# Work Projects Repository

This repository uses a **branch-per-project** workflow. Each project lives on its own isolated branch so that projects do not interfere with one another.

## How It Works

- **`main`** — This branch contains only the repo setup, guidelines, and documentation.
- **Each project gets its own branch** — e.g., `projects/excel-tracker`, `projects/terminal-app`, etc.
- Branches are created as **orphan branches**, meaning they share no file history with other branches. Switching branches completely swaps the working directory contents.

## Creating a New Project

To start a new project on its own isolated branch:

```bash
git checkout --orphan projects/my-new-project
git rm -rf .
# Add your project files...
git add .
git commit -m "Initial commit for my-new-project"
git push -u origin projects/my-new-project
```

## Branch Naming Convention

Use the `projects/` prefix for all project branches:

| Branch | Description |
|--------|-------------|
| `main` | Repo docs and setup only |
| `projects/excel-tracker` | Example: Excel spreadsheet project |
| `projects/terminal-app` | Example: Terminal-based tool |
| `projects/<name>` | Your project here |

## Security

- **Never commit secrets** (passwords, API keys, tokens) — they are blocked by `.gitignore`.
- **Keep the repo private** on GitHub (Settings > General > Danger Zone > Change visibility).
- See `SECURITY.md` for full guidelines.
