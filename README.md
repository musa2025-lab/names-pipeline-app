# Names Pipeline — scanned slips to village Excel files

Internal web app for the DelAgua beneficiary names pipeline. Team members upload
scanned CHW beneficiary registration slips, review the extracted rows on screen,
and download village Excel files ready for `compile_beneficiaries.py`.

Replaces the previous Power Automate flow **and** the AI Builder document model.

- **Nothing to install for users** — it's a web page.
- **No Claude account needed by users** — the API key lives only in the server's settings.
- **Locations validated against `g_Locations_UG.csv`**, so folder names can't drift
  from what the compile script looks up.
- **New/unknown areas are reported for review**, with typo detection.

---

## What it does

1. Upload one or many scanned slips (PDF).
2. Claude reads each slip and returns the header fields plus every beneficiary row.
3. **Location** is derived from the *file name* and matched against
   `g_Locations_UG.csv` — never from the handwriting (see *Design notes* below).
4. **Everything is editable on screen** — the operator checks against the paper and
   corrects anything misread before exporting.
5. Download a zip containing
   `Processed_Data/{Level 1}/{Level 2}/{Village}.xlsx`, plus
   `NEW_LOCATIONS_REVIEW.xlsx` if any location wasn't recognised.
6. Unzip into the country folder and run `RUN ME.bat` as usual.

---

## Run locally

Requires Python 3.11+.

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

Set your key and an access mode, then start it:

```powershell
# PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
$env:ALLOW_NO_AUTH = "true"        # local only - see Access control below
python -m streamlit run app.py
```

```bash
# bash
export ANTHROPIC_API_KEY="sk-ant-..."
export ALLOW_NO_AUTH=true          # local only
streamlit run app.py
```

Opens at <http://localhost:8501>.

> Name your PDFs after the village (`Ishanje.pdf`) — that's what lets the app
> resolve the location hierarchy automatically.

---

## Access control

The app processes beneficiary names, phone numbers and national ID numbers, so
**it refuses to start unless an access mode is configured.** Set exactly one:

| Variable | Use | Notes |
|---|---|---|
| `AUTH_MODE=platform` | **Preferred.** Azure App Service Authentication (Entra ID) in front of the app | Per-person sign-in with your existing Microsoft accounts. Set this *after* enabling Authentication in Azure. |
| `APP_PASSWORD=<password>` | **Fallback** when SSO isn't available | Weaker: no per-person audit trail, and revoking one person means rotating for everyone. |
| `ALLOW_NO_AUTH=true` | **Local development only** | Shows a persistent warning. Never set on a hosted URL. |

Both hosted modes **fail closed**: if `AUTH_MODE=platform` is set but no
signed-in user reaches the app (e.g. Authentication got switched off), the app
refuses to open rather than serving data.

---

## Deploy

Two options:

| Guide | Route | Time | Best for |
|---|---|---|---|
| **[DEPLOY_SIMPLE.md](DEPLOY_SIMPLE.md)** | Streamlit Community Cloud | ~5 min, free | Getting the team testing fast |
| **[DEPLOY.md](DEPLOY.md)** | Azure App Service | ~30 min, ~£10/mo | Real beneficiary data — stays in DelAgua's tenant, per-person SSO |

Both work from this same repo; only the secrets differ. Read the trade-off table
at the end of `DEPLOY_SIMPLE.md` before putting real slips through the free route.

Two things that are **not optional** either way:

1. **`ANTHROPIC_API_KEY` goes in the host's settings** — never in this repo.
2. **An access mode must be set** — `AUTH_MODE=platform` (Azure SSO) or
   `APP_PASSWORD` (shared password). The app refuses to start without one.

Configuration is read from `st.secrets` **or** environment variables, so the same
code runs unchanged on either host.

---

## Files

| File | Purpose |
|---|---|
| `app.py` | The web UI — upload, review/edit, download |
| `auth.py` | Access gate — platform SSO or shared password; fails closed |
| `settings.py` | Reads config from `st.secrets` or environment variables |
| `pipeline.py` | Extraction, Excel building, location validation, typo detection |
| `g_Locations_UG.csv` | Authoritative location list; drives the dropdowns |
| `requirements.txt` | Dependencies |
| `.streamlit/config.toml` | Upload limit, theme, XSRF settings |
| `DEPLOY.md` | Deployment and day-to-day usage guide |

---

## Design notes

Two decisions here are deliberate and shouldn't be "simplified" away:

**1. Location comes from the file name, not the scan.**
Testing the same slip twice, the handwritten header was read as
`KABIRIZI ISHANJE` once and `KABURIRA ISHANJE` the next; the VHT phone came out
`0796327963` then `0763296`. The beneficiary *table* was stable both times — it's
the scrawled header fields that vary. Since paths must match
`g_Locations_UG.csv` exactly (a `Rweikiriro` vs `Rweikiniro` mismatch silently
creates a phantom location and misfiles everyone under it), the app derives
location from the file name and shows the scanned values only as a cross-check.

**2. The app never assigns GIDs.**
`compile_beneficiaries.py` already owns GID allocation (Stage 3: walks the tree,
reuses existing parents, continues from the current maximum). Two systems minting
GIDs independently would collide. This app only *reports* unknown locations; the
compile script stays the single source of truth.

**Handwriting recognition is good, not perfect.** The on-screen review step is not
decorative — it is the quality gate. Don't remove it.

---

## Roadmap

- **Multi-country support.** `compile_beneficiaries.py` already handles varying
  hierarchies via `config.json` (`levels`) — Gambia 3 levels, Sierra Leone 4,
  Rwanda 5. This app currently hardcodes the 3-level Uganda shape. The intended
  fix is to read the *same* `config.json` + `g_Locations_XX.csv` pair rather than
  invent a second config format, and generate N dropdowns from `levels`.
