# Uganda Names Pipeline — web app

Replaces the Power Automate flow **and** the AI Builder model. Team members open
a URL, drag in scanned slips, check the extracted rows on screen, and download
the village Excel files ready for `compile_beneficiaries.py`.

- **No installs for the team** — it's a web page.
- **No Claude account for the team** — the API key sits only in the server's settings.
- **No Power Automate, no AI Builder, no Graph API, no webhooks.**

---

## Try it on your own machine first (5 minutes)

```bash
cd web_app
pip install -r requirements.txt
```

Then set the key and run (PowerShell):

```powershell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
streamlit run app.py
```

It opens at `http://localhost:8501`. Upload `Ishanje.pdf`, confirm the table
looks right, download the zip. Once you're happy, deploy it so the team can use it.

---

## Deploy to Azure App Service

You said you have Azure and IT admin access — this uses the plain App Service,
no app registrations or Graph permissions needed.

### 1. Create the app

```bash
az group create --name uganda-pipeline-rg --location westeurope

az appservice plan create --name uganda-pipeline-plan \
  --resource-group uganda-pipeline-rg --sku B1 --is-linux

az webapp create --name uganda-names-pipeline \
  --resource-group uganda-pipeline-rg \
  --plan uganda-pipeline-plan \
  --runtime "PYTHON:3.11"
```

Names are placeholders — use your org's convention. The app name becomes the URL
(`https://uganda-names-pipeline.azurewebsites.net`), so it must be globally unique.

### 2. Tell Azure how to start Streamlit

```bash
az webapp config set --name uganda-names-pipeline \
  --resource-group uganda-pipeline-rg \
  --startup-file "python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0"
```

### 3. Add the API key

```bash
az webapp config appsettings set --name uganda-names-pipeline \
  --resource-group uganda-pipeline-rg \
  --settings ANTHROPIC_API_KEY="sk-ant-..." \
             SCM_DO_BUILD_DURING_DEPLOYMENT=true
```

Generate a **new** key at console.anthropic.com for this app rather than reusing
a personal one, so you can revoke it independently.

### 4. Deploy the code

From inside `web_app/`:

```bash
az webapp up --name uganda-names-pipeline --resource-group uganda-pipeline-rg
```

Make sure `g_Locations_UG.csv` is in this folder when you deploy — the app reads
it to populate the Sub county / Parish / Village dropdowns.

### 5. Lock it down — do not skip this

The app handles beneficiary names, phone numbers and national IDs. Left open,
the URL is public.

Azure Portal → your App Service → **Settings → Authentication** →
**Add identity provider** → **Microsoft** →
- Tenant type: **Workforce**
- Client application requirement: **Allow requests only from this application**
- Restrict access: **Require authentication**
- Unauthenticated requests: **HTTP 302 redirect to login**

Team members then sign in with their existing DelAgua Microsoft 365 account. No
new passwords, no extra admin work per person.

### 6. Share the URL

`https://uganda-names-pipeline.azurewebsites.net`

---

## How the team uses it

1. Open the URL, sign in with their DelAgua account.
2. Drag in one or many scanned PDFs → **Process**.
3. For each slip:
   - **Location** is derived from the file name and matched to
     `g_Locations_UG.csv`. If the file name doesn't match a known village, they
     must pick from the dropdowns — download stays disabled until they do.
   - **Leader / VHT contacts** and the **beneficiary table** are editable —
     they check against the paper and fix anything misread.
4. **Download Excel files (.zip)** → unzip into the Uganda folder → run
   `RUN ME.bat` as usual.

**Name your PDFs after the village** (`Ishanje.pdf`, `Kabirizi.pdf`) — that's what
lets the app resolve Sub county / Parish / Village automatically.

---

## New areas (locations not yet in g_Locations)

Some slips will come from places that aren't in `g_Locations_UG.csv` yet. Tick
**"This is a new area"** on that slip and the dropdowns become free-text boxes.

Every unrecognised location is then collected into **NEW_LOCATIONS_REVIEW.xlsx**,
included at the top of the downloaded zip, with these columns:

| Column | What it tells you |
|---|---|
| Level / New name | Which level is unrecognised, and the name entered |
| Under Sub county / Under Parish | Where it would sit in the hierarchy |
| Closest existing name | The most similar name already in g_Locations |
| Similarity % | How close that match is |
| Likely typo? | `YES` / `no` / `NAME EXISTS` / `blocked` |
| Action | What to do about it |

**Why the similarity column matters.** A new location and a misspelling look
identical to software. `Rweikiriro` is not a new sub-county — it's `Rweikiniro`
typed wrong, and today it created a phantom folder that would have misfiled every
beneficiary under it. The sheet flags anything ≥80% similar as a suspected typo
and sorts those to the top.

The four verdicts:

- **`YES`** — very close to an existing name. Almost certainly a misspelling; fix
  the spelling rather than adding a new location.
- **`no`** — genuinely looks new. Confirm it's real, then `RUN ME.bat` will add it.
- **`NAME EXISTS`** — this exact name already exists, but under a *different*
  parent. The parent is wrong, not this name.
- **`blocked`** — a level above this one is a suspected typo, so this one can't be
  judged yet. Fix the parent and re-run.

**GIDs are deliberately not assigned here.** `compile_beneficiaries.py` already
owns that (Stage 3: it walks the tree, reuses existing parents, and continues GIDs
from the current maximum). Two systems minting GIDs independently would collide.
So the app produces the review sheet, and the compile script stays the single
source of truth.

After new locations are added, remember the existing rule from the pipeline doc:
**load the updated `g_Locations_UG.csv` into PowerSolve before uploading the
beneficiary CSV**, or the VillageID values won't resolve.

---

## Why location comes from the file name, not the scan

Testing the same slip twice, Claude read the handwritten header as
`KABIRIZI ISHANJE` once and `KABURIRA ISHANJE` the next time; the VHT phone came
out `0796327963` then `0763296`. The beneficiary *table* was stable both times —
it's the scrawled header fields that vary.

Since folder and file names must match `g_Locations_UG.csv` **exactly** (a
`Rweikiriro` vs `Rweikiniro` mismatch silently creates a phantom location and
misfiles everyone in it), the app never derives paths from handwriting. It uses
the file name, validates against the CSV, and shows the scanned values beside it
purely as a cross-check.

The same caution applies to the contact and beneficiary fields — **the review
step is not optional**. Handwriting recognition is good here, not perfect, which
is exactly why the table is editable before anything is exported.

---

## Costs

- **App Service B1:** roughly £10–13/month. Drop to F1 (free) for light use, at
  the cost of the app sleeping when idle and a slow first load.
- **Claude API:** a few pence per slip. Track it at console.anthropic.com.
- Nothing to patch or reboot; no AI Builder credits.

## Maintenance

- **When locations change:** replace `g_Locations_UG.csv` in this folder and
  redeploy (`az webapp up ...`) so the dropdowns stay current.
- **Key rotation:** update the `ANTHROPIC_API_KEY` app setting; no code change.
