"""Render a manuscript .docx and check its tables the way a reader would see them.

Structural inspection of a .docx does NOT catch the defects that matter - cells
that wrap, tables broken across a page, a rule drawn where none belongs. Those
only appear once the document is rendered. This renders via Word COM to PDF and
audits the drawn page.

    python verify_tables.py manuscript.docx
    python verify_tables.py manuscript.docx --xlsx-dir output/tables/publication
    python verify_tables.py manuscript.docx --keep-png out/     # eyeball them too

Checks
  1. vertical rules            must be zero - journals do not print them
  2. horizontal rules          should be 3 per table (x columns, since Word draws
                               each rule cell by cell); more means inside rules
  3. split tables              a table continuing across a page break
  4. blank pages               usually a page break inside a skipped section
  5. numbers                   every numeric token in each source xlsx must appear
                               in the rendered document (thousands separators
                               normalised - "19,959" is one number, not two)
  6. body vs tables            every "x.xx [lo-hi]" estimate cited in the text must
                               exist as "x.xx (lo-hi)" in some table

Exit code is non-zero if any check fails, so a build script can gate on it.
Rules and rationale: agent/feedback/figure_table_to_word.md
"""
import argparse
import glob
import os
import re
import sys
from collections import Counter

NUM = re.compile(r"-?\d+\.?\d*")
BRACKET_EST = re.compile(r"(\d+\.\d{2}) \[(\d+\.\d{2})[-–](\d+\.\d{2})\]")
PAREN_EST = re.compile(r"(\d+\.\d{2}) \((\d+\.\d{2})[-–](\d+\.\d{2})\)")


def render_pdf(docx_path, out_dir):
    """Word COM -> PDF. Word must be installed; this is Windows-only."""
    import win32com.client as win32
    os.makedirs(out_dir, exist_ok=True)
    pdf = os.path.join(out_dir, "render.pdf")
    if os.path.exists(pdf):
        try:
            os.remove(pdf)          # Word refuses to overwrite a PDF still held open
        except OSError:
            pdf = os.path.join(out_dir, "render_%d.pdf" % os.getpid())
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs(pdf, FileFormat=17)      # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()
    return pdf


def count_rules(page):
    """Vertical and horizontal rule segments drawn on one page."""
    v = h = 0
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "re":
                r = item[1]
                if r.width < 2 and r.height > 6:
                    v += 1
                elif r.height < 2 and r.width > 6:
                    h += 1
            elif item[0] == "l":
                a, b = item[1], item[2]
                if abs(a.x - b.x) < 0.5 and abs(a.y - b.y) > 6:
                    v += 1
                elif abs(a.y - b.y) < 0.5 and abs(a.x - b.x) > 6:
                    h += 1
    return v, h


def xlsx_tokens(path):
    """Numeric tokens in a sheet, excluding column 0 (row labels)."""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    out = Counter()
    for row in ws.iter_rows(values_only=True):
        for j, c in enumerate(row):
            if c is None or j == 0:
                continue
            # strip thousands separators FIRST or "19,959" tokenises as 19 + 959
            for t in NUM.findall(str(c).replace(",", "")):
                out[t] += 1
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--xlsx-dir", help="source tables; every number in them must appear")
    ap.add_argument("--xlsx-glob", default="*.xlsx")
    ap.add_argument("--out", default=None, help="where to write the render")
    ap.add_argument("--keep-png", metavar="DIR", help="also write page PNGs to eyeball")
    ap.add_argument("--sources", nargs="*", metavar="GLOB",
                    help="extra globs whose files must predate the build "
                         "(the xlsx dir and the builders beside the .docx are checked anyway)")
    ap.add_argument("--skip-freshness", action="store_true",
                    help="do not fail when the build is older than its inputs")
    args = ap.parse_args()

    import fitz

    out_dir = args.out or os.path.join(os.path.dirname(os.path.abspath(args.docx)),
                                       "_verify")
    pdf = render_pdf(args.docx, out_dir)
    doc = fitz.open(pdf)
    failures = []

    print("%s  (%d pages)" % (os.path.basename(args.docx), doc.page_count))
    print("-" * 68)

    # 0. freshness - is this document actually built from the current inputs?
    if not args.skip_freshness:
        built = os.path.getmtime(args.docx)
        here = os.path.dirname(os.path.abspath(args.docx))
        srcs = (glob.glob(os.path.join(here, "*.py")) + glob.glob(os.path.join(here, "*.md")))
        if args.xlsx_dir:
            srcs += glob.glob(os.path.join(args.xlsx_dir, args.xlsx_glob))
        for g in (args.sources or []):
            srcs += glob.glob(g)
        # A build writes siblings too (the markdown twin, a generated index). Anything
        # stamped within a few seconds of the .docx came out of the same run, so it is
        # an output, not an input that the build failed to pick up.
        srcs = [f for f in set(srcs)
                if os.path.abspath(f) != os.path.abspath(args.docx)
                and abs(os.path.getmtime(f) - built) > 5]
        newer = sorted((f for f in srcs if os.path.getmtime(f) > built),
                       key=os.path.getmtime, reverse=True)
        if newer:
            print("0. freshness: STALE - %d input(s) changed after the build" % len(newer))
            for f in newer[:6]:
                print("     %s" % os.path.basename(f))
            failures.append("the document is older than %d of its inputs - rebuild before "
                            "trusting anything below" % len(newer))
        else:
            print("0. freshness: ok (%d inputs all predate the build)" % len(srcs))

    # 1-2. rules
    total_v = 0
    per_page = []
    for i, page in enumerate(doc):
        v, h = count_rules(page)
        total_v += v
        per_page.append((i + 1, v, h))
    print("1. vertical rules: %d" % total_v)
    if total_v:
        bad = [p for p, v, _ in per_page if v]
        failures.append("vertical rules on pages %s - journals print none" % bad)
        print("   FAIL on pages %s" % bad)
    print("2. horizontal rules per page (3 x columns per table): %s"
          % [(p, h) for p, _, h in per_page if h])

    # 3. split tables. A long Table 1 legitimately spans pages - what must never
    #    happen is a continuation page opening with bare numbers. So a split is
    #    fine if the header row repeats (w:tblHeader), and a defect if it does not.
    split_ok, split_bad = [], []
    for i in range(doc.page_count - 1):
        this_txt, next_txt = doc[i].get_text().strip(), doc[i + 1].get_text().strip()
        _, hr_this = count_rules(doc[i])
        _, hr_next = count_rules(doc[i + 1])
        if not (hr_this and hr_next and this_txt and next_txt):
            continue
        lines_next = [l.strip() for l in next_txt.split("\n") if l.strip()]
        if not lines_next or re.match(r"(Table|eTable|Supplementary)", lines_next[0], re.I):
            continue                              # a new table, not a continuation
        blocks = doc[i + 1].get_text("dict")["blocks"]
        top = min((b["bbox"][1] for b in blocks if b.get("lines")), default=1e9)
        if top >= 130:
            continue                              # not flush to the top of the page
        prev_lines = {l.strip() for l in this_txt.split("\n")}
        (split_ok if lines_next[0] in prev_lines else split_bad).append(i + 2)
    print("3. continuation pages: %s repeat the header, %s do not"
          % (split_ok or "none", split_bad or "none"))
    if split_bad:
        failures.append("table continues on page(s) %s with no repeated header - "
                        "the reader sees bare numbers" % split_bad)

    # 4. blank pages
    blank = [i + 1 for i, p in enumerate(doc) if not p.get_text().strip()]
    print("4. blank pages: %s" % (blank or "none"))
    if blank:
        failures.append("blank page(s) %s" % blank)

    full = " ".join(p.get_text().replace("\n", " ") for p in doc)
    flat = full.replace(",", "")

    # 5. every source number present
    if args.xlsx_dir:
        print("5. source numbers")
        rendered = Counter(NUM.findall(flat))
        for path in sorted(glob.glob(os.path.join(args.xlsx_dir, args.xlsx_glob))):
            want = xlsx_tokens(path)
            missing = {t: n for t, n in want.items() if rendered[t] < 1}
            name = os.path.basename(path)
            if missing:
                print("   MISSING from %s: %s" % (name, sorted(missing)[:12]))
                failures.append("%s has numbers absent from the document" % name)
            else:
                print("   ok  %s (%d tokens)" % (name, sum(want.values())))
    else:
        print("5. source numbers: skipped (pass --xlsx-dir)")

    # 6. body estimates backed by a table
    in_text = set(BRACKET_EST.findall(full))
    in_tables = set(PAREN_EST.findall(full))
    orphans = sorted(in_text - in_tables)
    print("6. text estimates with no matching table value: %s" % (orphans or "none"))
    if orphans:
        failures.append("estimates cited in the text but in no table: %s" % orphans)

    if args.keep_png:
        os.makedirs(args.keep_png, exist_ok=True)
        for i, page in enumerate(doc):
            page.get_pixmap(dpi=150).save(
                os.path.join(args.keep_png, "page%02d.png" % (i + 1)))
        print("\npage images: %s" % args.keep_png)

    print("-" * 68)
    if failures:
        print("FAIL (%d)" % len(failures))
        for f in failures:
            print("  - %s" % f)
        return 1
    print("PASS - but still look at the pages; wrapped cells and ugly breaks are "
          "not something a counter can see.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
