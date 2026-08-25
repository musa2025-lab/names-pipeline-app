"""
Does sending a rendered image beat sending the PDF?

The scans carry an embedded OCR text layer produced by the scanner, and on the
sample we have that layer is garbage (Sub_County read as "PwEl4wIeo"). When a
PDF is sent, the model sees both the image and that text - so it may be
copying the scanner's mistakes instead of reading the picture.

Metric: the beneficiary slip numbers on this sheet run in an unbroken sequence,
so they can be scored objectively without a manual transcription.

Usage:  python test_image_vs_pdf.py <pdf> <first_slip_no> <row_count> [runs]
"""

import base64
import io
import json
import sys

import anthropic
import pymupdf

from pipeline import EXTRACTION_PROMPT, EXTRACTION_SCHEMA, MODEL

RENDER_DPI = 220  # ~1870x2420px for Letter - detailed, under the 2576px cap


def render_png(pdf_path: str) -> bytes:
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=RENDER_DPI)
    return pix.tobytes("png")


def call(source_block: dict, client, thinking: bool) -> dict:
    kwargs = dict(
        model=MODEL,
        max_tokens=16000,
        messages=[{"role": "user", "content": [
            source_block,
            {"type": "text", "text": EXTRACTION_PROMPT},
        ]}],
        output_config={"format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
    )
    if thinking:
        kwargs["thinking"] = {"type": "adaptive"}
    resp = client.messages.create(**kwargs)
    return json.loads(next(b.text for b in resp.content if b.type == "text"))


def score_slips(result: dict, first: int, count: int) -> tuple[int, int, list[str]]:
    expected = [f"NTU{n:06d}" for n in range(first, first + count)]
    got = [str(b.get("slip_number") or "").strip().upper()
           for b in result.get("beneficiaries", [])]
    hits = 0
    notes = []
    for i, exp in enumerate(expected):
        g = got[i] if i < len(got) else "<missing>"
        if g == exp:
            hits += 1
        elif len(notes) < 5:
            notes.append(f"row{i+1}: got {g!r} want {exp!r}")
    return hits, len(expected), notes


def main() -> None:
    pdf_path = sys.argv[1]
    first = int(sys.argv[2])
    count = int(sys.argv[3])
    runs = int(sys.argv[4]) if len(sys.argv) > 4 else 2

    client = anthropic.Anthropic()
    pdf_b64 = base64.standard_b64encode(open(pdf_path, "rb").read()).decode()
    png_b64 = base64.standard_b64encode(render_png(pdf_path)).decode()
    print(f"rendered PNG at {RENDER_DPI} dpi\n")

    configs = [
        ("PDF   + no thinking", {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}}, False),
        ("IMAGE + no thinking", {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}}, False),
        ("IMAGE + thinking",    {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}}, True),
    ]

    for label, block, thinking in configs:
        print(f"{'='*60}\n{label}\n{'='*60}")
        scores = []
        for i in range(runs):
            try:
                res = call(block, client, thinking)
                hits, total, notes = score_slips(res, first, count)
                scores.append(hits / total)
                print(f"  run {i+1}: slip numbers {hits}/{total} = {hits/total:5.1%}   "
                      f"(rows returned: {len(res.get('beneficiaries', []))})")
                for nte in notes:
                    print(f"        {nte}")
            except Exception as e:
                print(f"  run {i+1}: ERROR {type(e).__name__}: {str(e)[:120]}")
        if scores:
            print(f"  ---- mean: {sum(scores)/len(scores):.1%}\n")


if __name__ == "__main__":
    main()
