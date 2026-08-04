# Deploy the simple way — Streamlit Community Cloud

The fastest route to a shareable URL. No Azure, no CLI, no IT involvement, free.

**About 5 minutes.** You need: your GitHub login, an Anthropic API key, and a
password you choose for the team.

---

## Steps

### 1. Go to <https://share.streamlit.io> and sign in with GitHub

Use the **`musa2025-lab`** account — that's where the repo lives.

### 2. Click **Create app** → **Deploy a public app from GitHub**

If it asks for permission to read your repositories, allow it. The repo is
private, so Streamlit needs access to see it.

### 3. Fill in the form

| Field | Value |
|---|---|
| Repository | `musa2025-lab/names-pipeline-app` |
| Branch | `main` |
| Main file path | `app.py` |
| App URL | choose something like `delagua-names-pipeline` |

### 4. Before clicking Deploy — open **Advanced settings → Secrets**

Paste this, filling in your own values:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD = "choose-a-long-password"
```

Set the Python version to **3.11** or later in the same panel if offered.

> These secrets are stored by Streamlit, not in the repo. Never commit them.

### 5. Click **Deploy**

First build takes 2–3 minutes while it installs dependencies. You'll get a URL
like:

```
https://delagua-names-pipeline.streamlit.app
```

### 6. Test it, then share

Open the URL, enter the password, upload a slip. Once it works, send the team the
URL and the password (separately from each other, ideally).

---

## After it's live

- **Updates deploy themselves.** Push to `main` and the app rebuilds. No redeploy step.
- **Changing the password:** app settings → Secrets → edit → save. It restarts automatically.
- **Adding new locations:** replace `g_Locations_UG.csv` in the repo and push.
- **Sleeping:** free apps sleep after inactivity and take ~30 seconds to wake on the
  next visit. Not a fault — just tell the team to be patient on the first load.

---

## What you're trading for the simplicity

Worth understanding before you put real beneficiary data through it:

| | Streamlit Community Cloud | Azure App Service |
|---|---|---|
| Setup | ~5 min, no CLI | ~30 min, CLI or Portal |
| Cost | Free | ~£10–13/month |
| Runs on | Streamlit's infrastructure (Snowflake-owned) | **DelAgua's own Azure tenant** |
| Sign-in | Shared team password | Per-person Microsoft SSO |
| Audit trail | None — can't tell who processed what | Per-person |
| HTTPS | Yes | Yes |
| Sleeps when idle | Yes (~30s wake) | No (on a paid plan) |

**The two that actually matter for this app:**

1. **Beneficiary data would be processed on third-party infrastructure.** Slips
   contain names, phone numbers and national ID numbers. They aren't stored
   permanently — they're processed in memory and the Excel is downloaded — but
   they do pass through a platform outside DelAgua's control. Whether that's
   acceptable is a data-governance decision, not a technical one. **Worth a quick
   check with whoever owns data protection before real slips go through it.**

2. **A shared password means no per-person accountability.** You can't tell who
   processed which village, and removing one person's access means changing the
   password for everyone.

**A sensible middle path:** use Streamlit Cloud now to get the team testing with
**dummy or redacted slips** — zero friction, proves the workflow — and stand up
the Azure version (see [DEPLOY.md](DEPLOY.md)) for real data once you have time.
The code is identical; only the secrets differ.
