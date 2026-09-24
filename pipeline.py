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

EXTRACTION_PROMPT = """These are scanned Uganda CHW beneficiary registration \
sheets. There may be SEVERAL pages, all belonging to the same village - the \
sheets are scanned together into one file.

Work through EVERY page in order and return EVERY beneficiary row from ALL of \
them in a single combined list. Do not stop after the first page. A later page \
usually continues the same numbering, so the list should run unbroken from the \
first row of page 1 to the last row of the final page.

Take the header fields (Sub county, Parish, Village, Leader Name, Leader \
Phone, VHT Name, VHT Phone) from wherever they are filled in - usually the \
first page. For each beneficiary row capture Full Name, Beneficiary Slip \
Number, Phone Number, ID Number, Gender and Stove Question (Y/N).

Only include a beneficiary row if it has a full name or a slip number written \
on it - skip rows that are entirely blank, including the unused rows at the \
end of a sheet. If a single field is illegible or not filled in, return an \
empty string for that field rather than guessing. Preserve leading zeros on \
phone numbers exactly as written.
"""


RENDER_DPI = 220  # ~1870x2420px for Letter - detailed, under the 2576px cap
MAX_EDGE_PX = 2576  # model's high-resolution vision limit
JPEG_QUALITY = 80  # visually lossless for scanned text at this resolution


def _load_pymupdf():
    """Import PyMuPDF at call time, not at module import.

    Importing it at the top meant a missing or broken install took the whole
    app down with an ImportError on "from pipeline import (...)" - and
    Streamlit redacts the underlying message, so the cause was invisible.
    Deferring it lets the app start and report what actually went wrong.
    """
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        try:
            import fitz  # the package's name before 1.24.3
            return fitz
        except ImportError as exc:
            raise RuntimeError(
                "PyMuPDF isn't available on the server, so scanned pages can't "
                f"be converted to images. Underlying error: {exc}. "
                "Check that 'pymupdf' installed during the build."
            ) from exc


def render_pages_to_png(pdf_bytes: bytes) -> list[bytes]:
    """Rasterise every page of the PDF to JPEG.

    Scanners embed their own OCR text layer in these PDFs, and on real scans
    that layer is unreliable - one sample had the sub-county as "PwEl4wIeo" and
    every beneficiary slip number as "NTUO000NN" (letter O) instead of
    "NTU0000NN". Sending the PDF puts that text in front of the model alongside
    the image, and it copies the scanner's mistakes: measured 0/30 slip numbers
    correct from the PDF versus 30/30 from a rendered image of the same page.

    Rasterising throws the text layer away so only the picture is read.

    JPEG rather than PNG: at the same pixel dimensions a page is 674KB instead
    of 3,059KB, and 898KB instead of 4,078KB once base64-encoded for the API.
    PNG cost roughly 7MB of memory per page and pushed the app past Community
    Cloud's limit on a real batch. The dimensions are what the reading accuracy
    depends on, and those are unchanged.
    """
    pymupdf = _load_pymupdf()
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=RENDER_DPI)  # honours the page's /Rotate
            if max(pix.width, pix.height) > MAX_EDGE_PX:
                scale = MAX_EDGE_PX / max(pix.width, pix.height)
                pix = page.get_pixmap(dpi=int(RENDER_DPI * scale))
            pages.append(pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY))
            pix = None  # release the bitmap before rendering the next page
    finally:
        doc.close()
    return pages


def extract_from_pdf_bytes(pdf_bytes: bytes, api_key: str | None = None) -> dict:
    """Send one scanned slip to Claude and return the structured fields."""
    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    pages = render_pages_to_png(pdf_bytes)
    page_count = len(pages)
    content = []
    while pages:  # pop as we encode - never hold the raw and encoded copies
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": base64.standard_b64encode(pages.pop(0)).decode("utf-8"),
            },
        })
    content.append({"type": "text", "text": EXTRACTION_PROMPT})

    # Streamed rather than a plain create(): at this max_tokens the SDK refuses
    # a non-streaming call, since a reply that long could run past its 10-minute
    # ceiling. The result is identical - the whole message is collected before
    # anything is returned.
    with client.messages.stream(
        model=MODEL,
        # Scans arrive merged - one file holds every sheet for a village, six
        # pages and 300 rows is normal. At roughly 70 tokens a row, the old
        # 16000 ceiling truncated the reply mid-list and silently lost the tail.
        max_tokens=48000,
        messages=[{"role": "user", "content": content}],
        output_config={"format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
    ) as stream:
        response = stream.get_final_message()

    content.clear()  # drop the base64 payload before returning

    if response.stop_reason == "max_tokens":
        raise RuntimeError(
            f"This file has more rows than one reply can hold ({page_count} "
            "pages). Some beneficiaries would be missing, so nothing has been "
            "read. Split the PDF into smaller parts and upload them separately "
            "- the app combines sheets for the same village automatically."
        )

    text = next(b.text for b in response.content if b.type == "text")
    result = json.loads(text)
    result["__pages__"] = page_count  # shown in the UI so a merged file is obvious
    return result


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


# Words operators routinely add to a scan's filename that aren't part of the
# village name. "Ishanje village.pdf" and "Ishanje Village scan1.pdf" must
# resolve the same as "Ishanje.pdf" - people name files this way by default.
# NOTE: bare numbers are deliberately NOT stripped. Many villages are
# distinguished only by a trailing number - Migyera 1/2, Kyenjojo 1-4,
# Mutojo 1/2/3 - so removing it would file beneficiaries into a sibling
# village. Only numbers attached to a scan word ("scan2") are dropped.
FILENAME_NOISE = re.compile(
    r"\b(village|villages|cell|parish|sub\s?county|subcounty|district|"
    r"scan(?:ned)?\s*\d*|copy|final|signed|new|form|sheet|slip|slips|names?|pdf)\b",
    re.IGNORECASE,
)

# Below this, a filename is too unlike any known village to auto-fill from.
# Deliberately strict: a wrong guess that the operator accepts is worse than
# no guess at all, because it silently misfiles a whole village.
FILENAME_MATCH_THRESHOLD = 0.86

# If the two best candidates are this close, the filename doesn't identify one
# village and the operator must choose.
AMBIGUOUS_MARGIN = 0.05


def clean_filename_stem(filename: str) -> str:
    """Strip numbering and boilerplate words from a scan's filename."""
    stem = Path(filename).stem
    stem = re.sub(r"^\d+[.\s_-]+", "", stem)  # leading "1. " / "03_"
    # Separators must become spaces *before* the noise words are stripped:
    # "_" is a word character, so \bvillage\b never matches inside
    # "Ishanje_village_final".
    stem = re.sub(r"[_-]+", " ", stem)
    stem = FILENAME_NOISE.sub(" ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def guess_location(filename: str, tree: dict[str, dict[str, list[str]]]) -> tuple[str, str, str]:
    """Best-effort (sub_county, parish, village) for an uploaded file.

    Matches the PDF's filename against the known village list - the filename is
    what the old pipeline used and what compile_beneficiaries.py derives the
    village from, so it is far more reliable than reading handwriting. Returns
    empty strings where there's no confident match, leaving the user to pick.

    Matching is done on a cleaned stem, then falls back to a close-but-inexact
    match, so "Ishanje village.pdf" and "Ishanje Village scan1.pdf" resolve the
    same as "Ishanje.pdf".
    """
    stem = clean_filename_stem(filename)
    target = stem.casefold()
    if not target:
        return "", "", Path(filename).stem

    exact: list[tuple[str, str, str]] = []
    scored: list[tuple[float, tuple[str, str, str]]] = []
    for sub_county, parishes in tree.items():
        for parish, villages in parishes.items():
            for village in villages:
                name = village.casefold()
                if name == target:
                    exact.append((sub_county, parish, village))
                    continue
                scored.append((
                    difflib.SequenceMatcher(None, name, target).ratio(),
                    (sub_county, parish, village),
                ))

    # Some village names exist under more than one parish - NYAKABUNGO sits in
    # both KATOJO and RUHAAMA, RURAMA in both RWENGOMA and Kayenje. The file
    # name alone cannot say which, so the operator must choose rather than the
    # app filing a whole village under the wrong parish on a coin flip.
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        return "", "", stem

    scored.sort(key=lambda x: -x[0])
    if not scored or scored[0][0] < FILENAME_MATCH_THRESHOLD:
        return "", "", stem

    # Refuse to guess between near-equal candidates. "Rwemiyaga.pdf" scores
    # almost identically against "Rwemiyaga 1" and "Rwemiyaga 2"; picking one
    # would quietly file a whole village's beneficiaries under its sibling.
    if len(scored) > 1 and scored[0][0] - scored[1][0] < AMBIGUOUS_MARGIN:
        return "", "", stem
    return scored[0][1]


# --- Field checks -----------------------------------------------------------
#
# Handwriting reading is good on structured fields but unreliable on free text:
# real extractions have produced phone numbers of 7 and 13 digits, and numbers
# with no leading zero. These checks catch values that are *provably* wrong -
# they don't confirm a value is right, only that it can't be. The point is to
# tell the reviewer where to look, not to correct anything automatically.

UG_PHONE = re.compile(r"^0\d{9}$")            # 10 digits, leading zero
SLIP_NUMBER = re.compile(r"^[A-Z]{2,4}\d{4,8}$", re.IGNORECASE)
ID_NUMBER = re.compile(r"^[A-Z0-9]{8,20}$", re.IGNORECASE)
GENDER_VALUES = {"M", "F"}
YES_NO_VALUES = {"Y", "N"}


def check_beneficiary(row: dict) -> dict[str, str]:
    """Return {field: reason} for values that cannot be correct as written.

    Blank is never an error - an illegible field is meant to come back empty
    rather than guessed, and some beneficiaries genuinely have no phone.
    """
    problems: dict[str, str] = {}

    phone = str(row.get("phone_number") or "").strip()
    if phone and not UG_PHONE.match(phone):
        digits = sum(c.isdigit() for c in phone)
        if not phone.startswith("0"):
            problems["phone_number"] = f"{digits} digits, no leading zero"
        else:
            problems["phone_number"] = f"{digits} digits, expected 10"

    slip = str(row.get("slip_number") or "").strip()
    if slip and not SLIP_NUMBER.match(slip):
        problems["slip_number"] = "unexpected format"

    id_no = str(row.get("id_number") or "").strip()
    if id_no and not ID_NUMBER.match(id_no.replace(" ", "")):
        problems["id_number"] = "unexpected format"

    gender = str(row.get("gender") or "").strip().upper()
    if gender and gender not in GENDER_VALUES:
        problems["gender"] = "expected M or F"

    stove = str(row.get("stove_question") or "").strip().upper()
    if stove and stove not in YES_NO_VALUES:
        problems["stove_question"] = "expected Y or N"

    return problems


def check_slip_sequence(rows: list[dict]) -> list[str]:
    """Warn about duplicate or out-of-order slip numbers within one sheet.

    Slip numbers are pre-printed in an unbroken run, so a duplicate or a jump
    is a strong sign a row was misread - the scanner's own OCR turned
    NTU000031 into NTUO00031 on every row of one sample.
    """
    seen: dict[str, int] = {}
    warnings = []
    numbers = []
    prefixes: dict[str, list[int]] = {}
    for i, row in enumerate(rows, 1):
        slip = str(row.get("slip_number") or "").strip().upper()
        if not slip:
            continue
        if slip in seen:
            warnings.append(f"Row {i}: slip {slip} duplicates row {seen[slip]}")
        seen[slip] = i
        prefix = re.match(r"^[A-Z]*", slip).group()
        prefixes.setdefault(prefix, []).append(i)
        digits = re.sub(r"\D", "", slip)
        if digits:
            numbers.append((i, int(digits)))

    # Every slip on a sheet shares one pre-printed prefix, so an odd one out is
    # a misread - "NTUO" for "NTU0" is exactly the mistake the scanner's own
    # OCR made on every row of one sample.
    if len(prefixes) > 1:
        main = max(prefixes, key=lambda p: len(prefixes[p]))
        for prefix, row_numbers in prefixes.items():
            if prefix != main:
                where = ", ".join(str(n) for n in row_numbers[:5])
                warnings.append(
                    f"Row{'s' if len(row_numbers) > 1 else ''} {where}: prefix "
                    f"{prefix!r} differs from {main!r} used by the other rows"
                )

    for (prev_row, prev_n), (row_i, n) in zip(numbers, numbers[1:]):
        if n != prev_n + 1:
            warnings.append(
                f"Row {row_i}: slip jumps from {prev_n} to {n} "
                f"(expected {prev_n + 1})"
            )
    return warnings


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


def merge_by_village(records: list[dict]) -> dict[str, list[dict]]:
    """Group every slip's rows by its village file path, preserving order.

    A village usually runs to more than one sheet, so several uploads share an
    output path. They are merged into one file per village rather than written
    side by side: an earlier version suffixed the second file " (1)", and
    compile_beneficiaries.py then read "KAYENJE (1)" as a village in its own
    right and minted a location for it.

    Rows are appended as they come, never de-duplicated. A person appearing on
    two sheets is a real data problem, and Validation_Report.xlsx catches it
    against the whole dataset - dropping rows here would hide that, and risks
    discarding two people who genuinely share a name.
    """
    merged: dict[str, list[dict]] = {}
    for record in records:
        merged.setdefault(village_output_path(record), []).extend(
            record.get("beneficiaries", [])
        )
    return merged


def build_zip(
    records: list[dict], tree: dict[str, dict[str, list[str]]] | None = None
) -> bytes:
    """Bundle every processed slip into one zip, preserving folder structure.

    `records` are dicts with the header fields plus a "beneficiaries" list -
    i.e. the extraction output, optionally after the user edited it on screen.
    Slips for the same village are combined into a single file.

    If any location isn't in `tree`, a NEW_LOCATIONS_REVIEW.xlsx is added at the
    top of the zip so it can't be missed.
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, beneficiaries in merge_by_village(records).items():
            xlsx_bytes, _ = build_village_excel_bytes(beneficiaries)
            zf.writestr(path, xlsx_bytes)

        if tree:
            findings = collect_new_locations(records, tree)
            if findings:
                zf.writestr(
                    "NEW_LOCATIONS_REVIEW.xlsx", build_new_locations_bytes(findings)
                )
    return buffer.getvalue()
