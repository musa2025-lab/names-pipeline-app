"""
Core logic for the Uganda names web app: read a scanned beneficiary slip with
Claude, and build the village Excel file.

Kept separate from app.py (the UI) so this can be unit-tested and reused.
"""

import base64
import csv
import difflib
import io
import json
import re
import zipfile
from pathlib import Path

import anthropic
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

MODEL = "claude-opus-4-8"

BLUE = "174781"
WHITE = "FFFFFF"

# Column order must match what compile_beneficiaries.py expects to read.
BENEFICIARY_HEADERS = [
    "No.",
    "FULL NAME",
    "BENEFICIARY SLIP NUMBER",
    "PHONE NUMBER",
    "ID Number",
    "GENDER",
    "Stove Question (Y / N)",
]

# Keys of the per-beneficiary dict, in the same order as the headers above
# (minus "No.", which is generated).
BENEFICIARY_FIELDS = [
    "full_name",
    "slip_number",
    "phone_number",
    "id_number",
    "gender",
    "stove_question",
]

HEADER_FIELDS = [
    "sub_county",
    "parish",
    "village",
    "leader_name",
    "leader_phone",
    "vht_name",
    "vht_phone",
]

EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        **{f: {"type": "string"} for f in HEADER_FIELDS},
        "beneficiaries": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {f: {"type": "string"} for f in BENEFICIARY_FIELDS},
                "required": BENEFICIARY_FIELDS,
                "additionalProperties": False,
            },
        },
    },
    "required": HEADER_FIELDS + ["beneficiaries"],
    "additionalProperties": False,
}

EXTRACTION_PROMPT = """This is a scanned Uganda CHW beneficiary registration slip.

Extract the header fields (Sub county, Parish, Village, Leader Name, Leader \
Phone, VHT Name, VHT Phone) and every row of the Beneficiaries table (Full \
Name, Beneficiary Slip Number, Phone Number, ID Number, Gender, Stove \
Question Y/N).

Only include a beneficiary row if it has a full name or a slip number written \
on it - skip rows that are entirely blank. If a single field is illegible or \
not filled in, return an empty string for that field rather than guessing. \
Preserve leading zeros on phone numbers exactly as written.
"""


def extract_from_pdf_bytes(pdf_bytes: bytes, api_key: str | None = None) -> dict:
    """Send one scanned slip to Claude and return the structured fields."""
    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    response = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
        output_config={"format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
    )

    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


LOCATIONS_CSV = Path(__file__).resolve().parent / "g_Locations_UG.csv"


def load_locations(csv_path: Path | None = None) -> dict[str, dict[str, list[str]]]:
    """Build {Sub_county: {Parish: [Village, ...]}} from g_Locations_UG.csv.

    This is the authoritative spelling of every location. The app uses it to
    populate dropdowns so a folder name can never drift from what
    compile_beneficiaries.py will look up (a 'Rweikiriro' vs 'Rweikiniro'
    mismatch silently creates a phantom location and misfiles beneficiaries).
    """
    path = csv_path or LOCATIONS_CSV
    if not path.exists():
        return {}

    by_gid: dict[str, dict] = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            gid = (row.get("GID") or "").strip()
            if gid:
                by_gid[gid] = {
                    "type": (row.get("RegionType") or "").strip(),
                    "name": (row.get("RegionName") or "").strip(),
                    "parent": (row.get("ParentID") or "").strip(),
                }

    tree: dict[str, dict[str, list[str]]] = {}
    for node in by_gid.values():
        if node["type"] != "Sub_county":
            continue
        tree.setdefault(node["name"], {})

    for node in by_gid.values():
        if node["type"] != "Parish":
            continue
        parent = by_gid.get(node["parent"])
        if parent and parent["type"] == "Sub_county":
            tree.setdefault(parent["name"], {}).setdefault(node["name"], [])

    for node in by_gid.values():
        if node["type"] != "Village":
            continue
        parish = by_gid.get(node["parent"])
        if not parish or parish["type"] != "Parish":
            continue
        sub = by_gid.get(parish["parent"])
        if not sub or sub["type"] != "Sub_county":
            continue
        tree.setdefault(sub["name"], {}).setdefault(parish["name"], []).append(node["name"])

    for parishes in tree.values():
        for villages in parishes.values():
            villages.sort()
    return tree


def guess_location(filename: str, tree: dict[str, dict[str, list[str]]]) -> tuple[str, str, str]:
    """Best-effort (sub_county, parish, village) for an uploaded file.

    Matches the PDF's filename against the known village list - the filename is
    what the old pipeline used and what compile_beneficiaries.py derives the
    village from, so it is far more reliable than reading handwriting. Returns
    empty strings where there's no confident match, leaving the user to pick.
    """
    stem = Path(filename).stem
    stem = re.sub(r"^\d+[.\s]+", "", stem).strip()
    target = stem.casefold()

    for sub_county, parishes in tree.items():
        for parish, villages in parishes.items():
            for village in villages:
                if village.casefold() == target:
                    return sub_county, parish, village
    return "", "", stem


# A name this close to an existing one is far more likely a typo than a new
# place. 0.80 catches Rweikiriro/Rweikiniro (0.95) and Kabirizi/Kabirisi
# without flagging genuinely distinct names.
TYPO_THRESHOLD = 0.80

NEW_LOCATION_HEADERS = [
    "Level",
    "New name",
    "Under Sub county",
    "Under Parish",
    "Closest existing name",
    "Similarity %",
    "Likely typo?",
    "Action",
]


def _closest(name: str, candidates: list[str]) -> tuple[str, float]:
    """Closest existing name and its similarity ratio (0-1). ('', 0.0) if none."""
    if not name or not candidates:
        return "", 0.0
    best, best_score = "", 0.0
    for cand in candidates:
        score = difflib.SequenceMatcher(
            None, name.casefold(), cand.casefold()
        ).ratio()
        if score > best_score:
            best, best_score = cand, score
    return best, best_score


def classify_location(
    sub_county: str, parish: str, village: str, tree: dict[str, dict[str, list[str]]]
) -> list[dict]:
    """Report which of the three levels don't exist in g_Locations yet.

    Each entry carries the closest existing name so a reviewer can tell a real
    new location from a misspelling of an existing one - the two are
    indistinguishable to the software, and guessing wrong silently misfiles
    every beneficiary under a phantom location.
    """
    findings: list[dict] = []
    if not tree:
        return findings

    # Once a level is a suspected misspelling, its children can't be judged:
    # they may exist perfectly well under the correctly-spelled parent. Report
    # them as blocked rather than claiming they're new.
    parent_misspelled: str | None = None

    def add(
        level: str,
        name: str,
        candidates: list[str],
        all_at_level: list[str],
        under_sub: str,
        under_parish: str,
    ):
        nonlocal parent_misspelled
        match, score = _closest(name, candidates)

        # The name may not exist under the chosen parent but exist elsewhere at
        # the same level - that means the parent is wrong, not that this is new.
        exists_elsewhere = any(c.casefold() == name.casefold() for c in all_at_level)
        if exists_elsewhere and score < 0.999:
            match, score = next(
                c for c in all_at_level if c.casefold() == name.casefold()
            ), 1.0

        pct = round(score * 100)

        if parent_misspelled:
            verdict, action = (
                "blocked",
                f"Fix the {parent_misspelled} spelling first, then recheck.",
            )
        elif score >= 0.999:
            verdict, action = (
                "NAME EXISTS",
                "This exact name already exists under a different parent - the "
                "parent is wrong, not this name.",
            )
            parent_misspelled = parent_misspelled or level
        elif score >= TYPO_THRESHOLD:
            verdict, action = (
                "YES",
                "CHECK SPELLING - looks like the existing name shown here.",
            )
            parent_misspelled = level
        else:
            verdict, action = (
                "no",
                "Confirm it is genuinely new, then let RUN ME.bat add it.",
            )

        findings.append(
            {
                "Level": level,
                "New name": name,
                "Under Sub county": under_sub,
                "Under Parish": under_parish,
                "Closest existing name": match,
                "Similarity %": pct,
                "Likely typo?": verdict,
                "Action": action,
            }
        )

    all_subs = list(tree.keys())
    sub_known = sub_county in tree
    if sub_county and not sub_known:
        add("Sub county", sub_county, all_subs, all_subs, "", "")

    all_parishes = [p for parishes in tree.values() for p in parishes]
    parishes_in_sub = list(tree.get(sub_county, {}).keys()) if sub_known else []
    parish_known = sub_known and parish in tree[sub_county]
    if parish and not parish_known:
        add(
            "Parish",
            parish,
            parishes_in_sub or all_parishes,
            all_parishes,
            sub_county,
            "",
        )

    all_villages = [
        v for parishes in tree.values() for villages in parishes.values() for v in villages
    ]
    villages_in_parish = tree.get(sub_county, {}).get(parish, []) if parish_known else []
    village_known = parish_known and village in villages_in_parish
    if village and not village_known:
        add(
            "Village",
            village,
            villages_in_parish or all_villages,
            all_villages,
            sub_county,
            parish,
        )

    return findings


def collect_new_locations(
    records: list[dict], tree: dict[str, dict[str, list[str]]]
) -> list[dict]:
    """De-duplicated new-location findings across every processed slip."""
    seen: set[tuple] = set()
    out: list[dict] = []
    for record in records:
        for f in classify_location(
            record.get("sub_county", ""),
            record.get("parish", ""),
            record.get("village", ""),
            tree,
        ):
            key = (f["Level"], f["New name"], f["Under Sub county"], f["Under Parish"])
            if key in seen:
                continue
            seen.add(key)
            out.append(f)
    # Show the suspected typos first - they need a decision before upload.
    out.sort(key=lambda f: (f["Likely typo?"] != "YES", f["Level"], f["New name"]))
    return out


def build_new_locations_bytes(findings: list[dict]) -> bytes:
    """Review sheet of locations not yet in g_Locations_UG.csv.

    Deliberately does NOT assign GIDs - compile_beneficiaries.py owns GID
    allocation (Stage 3). Two systems minting GIDs independently would collide.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "New Locations"

    for c_idx, name in enumerate(NEW_LOCATION_HEADERS, 1):
        cell = ws.cell(row=1, column=c_idx, value=name)
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c_idx)].width = 26

    amber = PatternFill("solid", fgColor="FFF2CC")
    for r_idx, finding in enumerate(findings, start=2):
        for c_idx, key in enumerate(NEW_LOCATION_HEADERS, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=finding.get(key, ""))
            if finding.get("Likely typo?") == "YES":
                cell.fill = amber

    last_letter = get_column_letter(len(NEW_LOCATION_HEADERS))
    last_row = max(len(findings) + 1, 2)
    table = Table(displayName="NewLocations", ref=f"A1:{last_letter}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def safe_path_part(value: str, fallback: str) -> str:
    """Make a folder/file name safe, without silently changing spelling.

    Only strips characters Windows/SharePoint reject in names. Does NOT
    normalise case or spelling - folder names must keep matching
    g_Locations_UG.csv exactly (a 'Rweikiriro' vs 'Rweikiniro' mismatch
    creates a phantom location downstream).
    """
    cleaned = re.sub(r'[<>:"/\\|?*]', "", (value or "").strip())
    cleaned = cleaned.strip(". ")
    return cleaned or fallback


def build_village_excel_bytes(beneficiaries: list[dict]) -> tuple[bytes, int]:
    """Build a village Excel file. Returns (xlsx bytes, rows written).

    Rows with neither a name nor a slip number are skipped - same rule the
    validation script uses, applied here so blank scan artefacts never reach
    Master Data.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for c_idx, name in enumerate(BENEFICIARY_HEADERS, 1):
        cell = ws.cell(row=1, column=c_idx, value=name)
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c_idx)].width = 20

    row_count = 0
    for b in beneficiaries:
        full_name = str(b.get("full_name") or "").strip()
        slip = str(b.get("slip_number") or "").strip()
        if not full_name and not slip:
            continue
        row_count += 1
        r = row_count + 1
        ws.cell(row=r, column=1, value=row_count)
        ws.cell(row=r, column=2, value=full_name or None)
        ws.cell(row=r, column=3, value=slip or "-")
        for offset, field in enumerate(
            ["phone_number", "id_number", "gender", "stove_question"], start=4
        ):
            val = str(b.get(field) or "").strip()
            ws.cell(row=r, column=offset, value=val or "-")

    last_letter = get_column_letter(len(BENEFICIARY_HEADERS))
    last_row = max(row_count + 1, 2)
    table = Table(displayName="Table1", ref=f"A1:{last_letter}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), row_count


def village_output_path(record: dict) -> str:
    """Processed_Data/{Sub_county}/{Parish}/{Village}.xlsx for one slip."""
    sub_county = safe_path_part(record.get("sub_county"), "UNKNOWN_SUBCOUNTY")
    parish = safe_path_part(record.get("parish"), "UNKNOWN_PARISH")
    village = safe_path_part(record.get("village"), "UNKNOWN_VILLAGE")
    return f"Processed_Data/{sub_county}/{parish}/{village}.xlsx"


def build_zip(
    records: list[dict], tree: dict[str, dict[str, list[str]]] | None = None
) -> bytes:
    """Bundle every processed slip into one zip, preserving folder structure.

    `records` are dicts with the header fields plus a "beneficiaries" list -
    i.e. the extraction output, optionally after the user edited it on screen.

    If any location isn't in `tree`, a NEW_LOCATIONS_REVIEW.xlsx is added at the
    top of the zip so it can't be missed.
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        used_paths: dict[str, int] = {}
        for record in records:
            path = village_output_path(record)
            # Two slips for the same village would collide inside the zip;
            # suffix rather than silently overwrite one of them.
            if path in used_paths:
                used_paths[path] += 1
                stem, ext = path.rsplit(".", 1)
                path = f"{stem} ({used_paths[path]}).{ext}"
            else:
                used_paths[path] = 0
            xlsx_bytes, _ = build_village_excel_bytes(record.get("beneficiaries", []))
            zf.writestr(path, xlsx_bytes)

        if tree:
            findings = collect_new_locations(records, tree)
            if findings:
                zf.writestr(
                    "NEW_LOCATIONS_REVIEW.xlsx", build_new_locations_bytes(findings)
                )
    return buffer.getvalue()
