# ⚠️  BEFORE MAKING THIS REPOSITORY PUBLIC

> Read this file carefully before changing the repository visibility to **Public**.

---

## Checklist — Things to remove or replace

### 🔴 Sensitive data (MUST remove)

- [ ] `config.json` — contains real GUIDs, tenant ID, SharePoint paths
      → Make sure `config.json` is in `.gitignore` and is NOT committed
      → Only `config.example.json` (with placeholders) should be in the repo

- [ ] `.token_cache.bin` — Microsoft auth token cache (if accidentally committed)
      → Confirm it is listed in `.gitignore`

- [ ] `input/.step1_data.json` — may contain real SAP numbers
- [ ] `input/.step2_data.json` — may contain real SAP keys and CFIN dates
      → Confirm the entire `input/` folder is in `.gitignore`

- [ ] `output/` — contains real generated RMR files
      → Confirm the entire `output/` folder is in `.gitignore`

---

### 🟠 Company references (replace or remove)

Search the entire project for these strings and replace them:

| Find | Replace with |
|------|-------------|
| `Securitas` | *(your choice: remove, anonymize, or replace with generic name)* |
| `securitas.com` | `[company].com` or remove |
| `securitasgroup-my.sharepoint.com` | `[tenant].sharepoint.com` or remove |
| `hector.arenas@securitas.com` | your personal email or remove |
| `arenahe` | remove or anonymize |
| `maria_patino_securitas_com` | remove or anonymize |
| `maria.patino@securitas.com` | remove or anonymize |
| `C:\Users\arenahe\OneDrive - Securitas\` | use a generic path like `C:\Users\[user]\` |

Files to check:
- `STEP1.py` (login print message mentions the email)
- `STEP1.bat`, `STEP2.bat`, `STEP3.bat` (hardcoded `cd` paths)
- `config.example.json` (check placeholder values)
- `README.md` (check if any company name slipped in)

---

### 🟡 BAT files — hardcoded paths

The `.bat` files contain:
```bat
cd /d "C:\Users\arenahe\OneDrive - Securitas\RMR_ACTIVATION"
```

Before going public, replace this with a relative path or a generic placeholder:
```bat
cd /d "%~dp0"
```
This makes the script run from whatever folder it's located in — no hardcoded path needed.

---

### 🟢 Nice-to-have improvements before going public

- [ ] Add a `LICENSE` file (MIT is a good choice for portfolio projects)
- [ ] Add screenshots or a short demo GIF to the README
- [ ] Add a "How it works" diagram to the README
- [ ] Write a short description for the GitHub repo (the subtitle under the name)
- [ ] Pin the repo to your GitHub profile

---

## How to verify nothing sensitive is committed

Run this in Git Bash before making public:

```bash
# Search for sensitive strings across all committed files
git log --all --full-history -- config.json
git grep "securitas" $(git rev-list --all)
git grep "arenahe" $(git rev-list --all)
```

If any sensitive data appears in the history, you'll need to purge it with `git filter-repo` before going public.

---

*Keep this file in the repo as a private reminder. Delete it or move it to a private note before publishing.*
