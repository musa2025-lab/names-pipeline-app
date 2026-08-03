"""
Uganda Names Pipeline - internal web app.

Team members open the URL, drag in scanned beneficiary slips, review/correct
the extracted table on screen, then download the village Excel files in the
folder structure compile_beneficiaries.py expects.

Nobody needs a Claude account or any local install - the API key lives only in
this app's server-side settings.

Run locally:   streamlit run app.py
"""

import os

import pandas as pd
import streamlit as st

from auth import require_access, sign_out_button
from pipeline import (
    BENEFICIARY_FIELDS,
    build_zip,
    classify_location,
    collect_new_locations,
    extract_from_pdf_bytes,
    guess_location,
    load_locations,
    village_output_path,
)

# Column labels shown in the on-screen editor, mapped to the internal keys.
EDITOR_COLUMNS = {
    "full_name": "FULL NAME",
    "slip_number": "SLIP NUMBER",
    "phone_number": "PHONE NUMBER",
    "id_number": "ID NUMBER",
    "gender": "GENDER",
    "stove_question": "STOVE (Y/N)",
}

PICK = "— select —"

st.set_page_config(page_title="Uganda Names Pipeline", page_icon="📋", layout="wide")

# Gate before anything else renders - this app handles beneficiary personal data.
CURRENT_USER = require_access()
sign_out_button()
st.sidebar.caption(f"Signed in as: {CURRENT_USER}")

st.title("Uganda Names Pipeline")
st.caption(
    "Upload scanned beneficiary registration slips. Check the extracted rows, "
    "correct anything the scan got wrong, then download the Excel files."
)

if not os.environ.get("ANTHROPIC_API_KEY"):
    st.error(
        "This app isn't configured yet — ANTHROPIC_API_KEY is missing from the "
        "server settings. Contact whoever set it up."
    )
    st.stop()

LOCATIONS = load_locations()
if not LOCATIONS:
    st.warning(
        "g_Locations_UG.csv wasn't found next to the app, so Sub county / Parish / "
        "Village must be typed by hand. Spelling then has to match the CSV exactly.",
        icon="⚠️",
    )

if "records" not in st.session_state:
    st.session_state.records = {}  # filename -> extracted dict

uploaded = st.file_uploader(
    "Scanned slips (PDF)", type=["pdf"], accept_multiple_files=True
)

col_a, col_b = st.columns([1, 5])
with col_a:
    process_clicked = st.button("Process", type="primary", disabled=not uploaded)
with col_b:
    if st.session_state.records and st.button("Clear all"):
        st.session_state.records = {}
        st.rerun()

if process_clicked:
    progress = st.progress(0.0, text="Starting…")
    for i, uf in enumerate(uploaded):
        progress.progress(i / len(uploaded), text=f"Reading {uf.name}…")
        try:
            st.session_state.records[uf.name] = extract_from_pdf_bytes(uf.getvalue())
        except Exception as exc:  # surface the failure, keep processing the rest
            st.session_state.records[uf.name] = {"__error__": str(exc)}
    progress.progress(1.0, text="Done")
    progress.empty()

if not st.session_state.records:
    st.info("No slips processed yet. Upload one or more PDFs and press **Process**.")
    st.stop()

st.divider()

edited_records: list[dict] = []
blocked = False

for filename, record in st.session_state.records.items():
    if "__error__" in record:
        with st.expander(f"❌  {filename} — could not be read", expanded=True):
            st.error(record["__error__"])
        continue

    beneficiaries = record.get("beneficiaries", [])
    g_sub, g_parish, g_village = guess_location(filename, LOCATIONS)

    with st.expander(
        f"📄  {filename} — {len(beneficiaries)} row(s)", expanded=True
    ):
        st.markdown("##### Location")
        st.caption(
            "Taken from the file name and matched against g_Locations_UG.csv — "
            "not from the handwriting, which is unreliable for these fields."
        )

        is_new = st.checkbox(
            "This is a **new area** — not yet in g_Locations",
            key=f"new_{filename}",
            help=(
                "Only tick this for a genuinely new location. If it's an existing "
                "place spelled differently, leave it unticked and pick from the "
                "dropdowns instead."
            ),
        )

        if LOCATIONS and not is_new:
            sub_options = [PICK] + sorted(LOCATIONS.keys())
            c1, c2, c3 = st.columns(3)
            sub_county = c1.selectbox(
                "Sub county",
                sub_options,
                index=sub_options.index(g_sub) if g_sub in sub_options else 0,
                key=f"sc_{filename}",
            )
            parish_options = [PICK] + sorted(LOCATIONS.get(sub_county, {}).keys())
            parish = c2.selectbox(
                "Parish",
                parish_options,
                index=parish_options.index(g_parish) if g_parish in parish_options else 0,
                key=f"pa_{filename}",
            )
            village_options = [PICK] + sorted(LOCATIONS.get(sub_county, {}).get(parish, []))
            village = c3.selectbox(
                "Village",
                village_options,
                index=village_options.index(g_village) if g_village in village_options else 0,
                key=f"vi_{filename}",
            )
            if PICK in (sub_county, parish, village):
                st.error(
                    "Pick Sub county, Parish and Village before downloading. If this "
                    "really is a new area, tick the box above instead.",
                    icon="🚫",
                )
                blocked = True
        else:
            c1, c2, c3 = st.columns(3)
            sub_county = c1.text_input(
                "Sub county", g_sub or record.get("sub_county", ""), key=f"sc_{filename}"
            )
            parish = c2.text_input(
                "Parish", g_parish or record.get("parish", ""), key=f"pa_{filename}"
            )
            village = c3.text_input("Village", g_village, key=f"vi_{filename}")

            if not (sub_county.strip() and parish.strip() and village.strip()):
                st.error(
                    "Fill in Sub county, Parish and Village.", icon="🚫"
                )
                blocked = True

            # Warn immediately if a "new" name looks like an existing one.
            for finding in classify_location(
                sub_county.strip(), parish.strip(), village.strip(), LOCATIONS
            ):
                if finding["Likely typo?"] == "YES":
                    st.warning(
                        f"**{finding['Level']} \"{finding['New name']}\"** is "
                        f"{finding['Similarity %']}% similar to the existing "
                        f"**\"{finding['Closest existing name']}\"** — is this a "
                        "misspelling rather than a new area?",
                        icon="🔍",
                    )
                else:
                    st.info(
                        f"{finding['Level']} **\"{finding['New name']}\"** will be "
                        "listed as new for review.",
                        icon="🆕",
                    )

        with st.container():
            st.markdown("##### What the scan said (for cross-checking only)")
            st.caption(
                f"Sub county: `{record.get('sub_county','')}`  ·  "
                f"Parish: `{record.get('parish','')}`  ·  "
                f"Village: `{record.get('village','')}`"
            )

        st.markdown("##### Leader & VHT contacts")
        st.caption("Read from the scan — check these against the paper before accepting.")
        h4, h5, h6, h7 = st.columns(4)
        leader_name = h4.text_input("Leader name", record.get("leader_name", ""), key=f"ln_{filename}")
        leader_phone = h5.text_input("Leader phone", record.get("leader_phone", ""), key=f"lp_{filename}")
        vht_name = h6.text_input("VHT name", record.get("vht_name", ""), key=f"vn_{filename}")
        vht_phone = h7.text_input("VHT phone", record.get("vht_phone", ""), key=f"vp_{filename}")

        st.markdown("##### Beneficiaries")
        st.caption("Edit any cell; add or delete rows as needed.")
        df = pd.DataFrame(beneficiaries, columns=BENEFICIARY_FIELDS).rename(
            columns=EDITOR_COLUMNS
        )
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            key=f"ed_{filename}",
            column_config={
                label: st.column_config.TextColumn(label)
                for label in EDITOR_COLUMNS.values()
            },
        )

        reverse_map = {v: k for k, v in EDITOR_COLUMNS.items()}
        edited_beneficiaries = (
            edited_df.rename(columns=reverse_map).fillna("").to_dict("records")
        )

        rebuilt = {
            "sub_county": "" if sub_county == PICK else sub_county.strip(),
            "parish": "" if parish == PICK else parish.strip(),
            "village": "" if village == PICK else village.strip(),
            "leader_name": leader_name,
            "leader_phone": leader_phone,
            "vht_name": vht_name,
            "vht_phone": vht_phone,
            "beneficiaries": edited_beneficiaries,
        }
        edited_records.append(rebuilt)

        if PICK not in (sub_county, parish, village):
            st.caption(f"Will be saved as  `{village_output_path(rebuilt)}`")

st.divider()

if edited_records:
    total_rows = sum(
        1
        for r in edited_records
        for b in r["beneficiaries"]
        if str(b.get("full_name") or "").strip() or str(b.get("slip_number") or "").strip()
    )
    st.subheader(
        f"Ready: {len(edited_records)} village file(s), {total_rows} beneficiary row(s)"
    )

    new_locations = collect_new_locations(edited_records, LOCATIONS)
    if new_locations:
        typos = [f for f in new_locations if f["Likely typo?"] == "YES"]
        st.markdown("##### New areas detected")
        if typos:
            st.warning(
                f"{len(typos)} of these look like **misspellings of existing "
                "locations**, not new areas. Check them before uploading — a wrong "
                "one creates a phantom location and misfiles everyone under it.",
                icon="🔍",
            )
        st.dataframe(pd.DataFrame(new_locations), use_container_width=True)
        st.caption(
            "This table is included in the zip as **NEW_LOCATIONS_REVIEW.xlsx**. "
            "GIDs are not assigned here — `RUN ME.bat` does that, and loading the "
            "updated `g_Locations_UG.csv` into PowerSolve remains a separate step."
        )

    if blocked:
        st.button("Download Excel files (.zip)", disabled=True, type="primary")
        st.caption("Resolve the location warnings above to enable the download.")
    else:
        st.download_button(
            "Download Excel files (.zip)",
            data=build_zip(edited_records, LOCATIONS),
            file_name="Processed_Data.zip",
            mime="application/zip",
            type="primary",
        )
        st.caption(
            "Unzip into your Uganda folder — it recreates "
            "`Processed_Data/{Sub county}/{Parish}/{Village}.xlsx`, then run "
            "**RUN ME.bat** as usual."
        )
