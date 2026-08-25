#!/usr/bin/env python3
"""Compare two .docx across EVERY axis that can carry an edit.

Why this exists (2026-08-19): a delivered brief came back hand-edited.
I diffed `doc.paragraphs`, saw 41/41 identical, and reported "prose unchanged -
only the font size differs". Wrong: `doc.paragraphs` does NOT include table cell
text, and the user had also restructured a table (long -> wide pivot with a
merged span header). The user caught it, not me.

"The prose did not change" is not "the content did not change."

Axes compared here:
  1. prose        - doc.paragraphs
  2. tables       - cell text, row/col counts, merge shape
  3. geometry     - column widths, image sizes, section margins
  4. type sizes   - measured from a rendered PDF, not from XML

Axis 4 needs rendering because `w:sz` is frequently ABSENT on runs (size is
inherited from the style), so reading XML reports None and hides a real change.
Only the rendered PDF tells you what the reader actually sees.

Usage:
    python docx_diff.py OLD.docx NEW.docx            # axes 1-3
    python docx_diff.py OLD.docx NEW.docx --render   # + axis 4 (needs Word COM)

Exit code 1 if any axis differs, else 0.
"""
import argparse
import difflib
import os
import sys
import tempfile
from collections import Counter

from docx import Document


# ── helpers ───────────────────────────────────────────────────────────────────
def emu_cm(v):
    return None if v is None else round(v / 360000, 2)


def prose(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def tables(path):
    """Cell text plus dimensions. Merged cells repeat their text, which is what
    makes a merge visible here at all."""
    out = []
    for ti, t in enumerate(Document(path).tables):
        out.append(f"### TABLE {ti}  rows={len(t.rows)} cols={len(t.columns)}")
        for ri, row in enumerate(t.rows):
            cells = [c.text.replace("\n", "\\n").strip() for c in row.cells]
            out.append(f"T{ti} r{ri}: " + " | ".join(cells))
    return out


def geometry(path):
    """Keyed, so it can be compared item by item.

    NOT a unified diff: with 15+ tables the context lines around a hunk look just
    like unchanged items, and a changed table outside the context window is simply
    absent. Both misled a reader into "only T5 changed" when two tables had moved.
    Every key is compared explicitly and only real changes print.
    """
    d = Document(path)
    out = {}
    s = d.sections[0]
    out["page"] = (f"{emu_cm(s.page_width)}x{emu_cm(s.page_height)} "
                   f"margins L{emu_cm(s.left_margin)} R{emu_cm(s.right_margin)} "
                   f"T{emu_cm(s.top_margin)} B{emu_cm(s.bottom_margin)}")
    st = d.styles["Normal"]
    out["Normal style"] = f"{st.font.name} {st.font.size.pt if st.font.size else None}pt"
    # Key on index only. Putting the header text in the key made a renamed column
    # look like one item removed plus another added, instead of one value changed.
    # Header text belongs to the tables axis; this axis is geometry alone.
    for ti, t in enumerate(d.tables):
        out[f"T{ti} colwidths"] = str([emu_cm(c.width) for c in t.rows[0].cells])
    for i, sh in enumerate(d.inline_shapes):
        out[f"img{i} size"] = f"{emu_cm(sh.width)}x{emu_cm(sh.height)} cm"
    return out


def render_pdf(docx_path, out_pdf):
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(out_pdf), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()


def type_profile(pdf_path):
    """Chars rendered at each point size - the only trustworthy size measurement."""
    import fitz
    d = fitz.open(pdf_path)
    c = Counter()
    for page in d:
        for blk in page.get_text("dict")["blocks"]:
            for line in blk.get("lines", []):
                for sp in line["spans"]:
                    if sp["text"].strip():
                        c[round(sp["size"], 1)] += len(sp["text"].strip())
    pages = d.page_count
    d.close()
    return c, pages


def report_seq(name, a, b):
    """Line sequences (prose, table cells) -> unified diff. Returns changed count."""
    if a == b:
        print(f"  [SAME] {name}")
        return 0
    changed = 0
    lines = []
    for line in difflib.unified_diff(a, b, "OLD", "NEW", n=1, lineterm=""):
        if line.startswith(("---", "+++", "@@")):
            continue
        if line.startswith(("+", "-")):
            changed += 1
        lines.append("      " + line[:170])
    print(f"  [DIFF] {name} - {changed} changed line(s)")
    for l in lines:
        print(l)
    return changed


def report_map(name, a, b):
    """Keyed items -> explicit per-key comparison. Returns changed count."""
    keys = list(a) + [k for k in b if k not in a]
    diffs = [(k, a.get(k), b.get(k)) for k in keys if a.get(k) != b.get(k)]
    if not diffs:
        print(f"  [SAME] {name} ({len(keys)} items)")
        return 0
    print(f"  [DIFF] {name} - {len(diffs)} of {len(keys)} item(s) changed")
    for k, va, vb in diffs:
        print(f"      {k}")
        print(f"          OLD {va}")
        print(f"          NEW {vb}")
    return len(diffs)


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--render", action="store_true",
                    help="also measure rendered type sizes (Word COM + PyMuPDF)")
    args = ap.parse_args()

    for p in (args.old, args.new):
        if not os.path.exists(p):
            sys.exit(f"missing: {p}")

    print(f"OLD {args.old}\nNEW {args.new}\n")
    tally = {}

    print("1. prose (doc.paragraphs)")
    tally["prose"] = report_seq("prose", prose(args.old), prose(args.new))

    print("\n2. tables (cell text + dimensions)")
    tally["tables"] = report_seq("tables", tables(args.old), tables(args.new))

    print("\n3. geometry (widths, images, margins)")
    tally["geometry"] = report_map("geometry", geometry(args.old), geometry(args.new))

    print("\n4. rendered type sizes")
    tally["type sizes"] = None
    if not args.render:
        print("  [SKIP] pass --render to measure (XML w:sz is inherited and unreliable)")
    else:
        try:
            with tempfile.TemporaryDirectory() as td:
                pa, pb = os.path.join(td, "a.pdf"), os.path.join(td, "b.pdf")
                render_pdf(args.old, pa)
                render_pdf(args.new, pb)
                ca, na = type_profile(pa)
                cb, nb = type_profile(pb)
            print(f"      pages {na} -> {nb}")
            print(f"      {'pt':>6} {'OLD':>8} {'NEW':>8}")
            for k in sorted(set(ca) | set(cb)):
                flag = "  <<" if ca.get(k, 0) != cb.get(k, 0) else ""
                print(f"      {k:>6} {ca.get(k,0):>8} {cb.get(k,0):>8}{flag}")
            n_diff = sum(1 for k in set(ca) | set(cb) if ca.get(k, 0) != cb.get(k, 0))
            tally["type sizes"] = n_diff + (na != nb)
            print(f"  [{'DIFF' if tally['type sizes'] else 'SAME'}] type sizes")
        except Exception as e:            # rendering is best-effort, never fatal
            print(f"  [SKIP] render failed: {e}")

    # Summary last, so a filtered or truncated read of this output still cannot
    # miss that an axis changed. "Only X changed" must be read off these counts,
    # never off the body above.
    print("\n" + "=" * 46)
    print("SUMMARY  (changes per axis)")
    for k, v in tally.items():
        print(f"  {k:<12} {'not measured' if v is None else v}")
    print("=" * 46)
    changed = any(v for v in tally.values() if v)
    print("DIFFERENCES FOUND" if changed else "IDENTICAL on all measured axes")
    sys.exit(1 if changed else 0)


if __name__ == "__main__":
    main()
