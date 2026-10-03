"""Render a manuscript .docx and check its tables the way a reader would see them.

Structural inspection of a .docx does NOT catch the defects that matter - cells
that wrap, tables broken across a page, a rule drawn where none belongs. Those
only appear once the document is rendered. This renders via Word COM to PDF and
audits the drawn page.

    python verify_tables.py manuscript.docx
    python verify_tables.py manuscript.docx --xlsx-dir output/tables/publication

⚠ --xlsx-dir 에는 «원고 표가 된» 워크북만 준다. output/tables 전체를 주면 중간 산출물
  (전체 정밀도 소수, 표본흐름의 시작 행수 등)이 전부 「문서에 없다」로 신고돼 진짜 결함이
  그 소음에 묻힌다 — 2026-09-08 실측: 9건 FAIL 중 진짜는 0건이었고, 대상을 좁히니
  「source numbers: ok」가 됐다.
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


## ⭐ 아래 둘은 «순수함수»다 — Word COM·PDF 없이 시험된다(`verify_tables_selftest.py`).
##    main() 안에 묻혀 있던 것을 2026-09-10 에 올렸다. 실제로 물린 곳이 렌더가 아니라
##    여기였고, 묻혀 있는 동안은 시험할 방법이 없었다.

## ⛔ 저널은 "Include line and page numbering" 을 «요구»하고, 그 번호가 PDF 텍스트의
##    첫 줄로 나온다. 거르지 않으면 「연속 페이지 첫 줄이 앞 페이지에도 있는가」가
##    **쪽번호끼리 비교**하게 되고, 쪽번호는 페이지마다 다르므로 절대 일치하지 않는다 —
##    규정을 지킨 «모든» 원고에서 「헤더 반복 안 됨」 위양성이 난다.
##    실측 2026-09-08: '26 / 27' 로 비교돼 FAIL, 거르니 'Item' 이 되어 일치(정상).
PAGENUM = re.compile(r"^\s*\d+\s*(?:/|of)\s*\d+\s*$|^\s*\d{1,4}\s*$")


def first_real(txt):
    """쪽번호·줄번호를 «맨 앞에서 연속으로만» 건너뛴 실제 첫 줄들.

    ⚠ 줄번호는 1~4자리 «맨숫자 한 줄»이라 표의 정수 셀과 모양이 같다. 그래서 맨 앞의
       연속분만 건너뛴다 — 표 «안»의 숫자 줄은 지우지 않는다.
    ⚠ 전부 번호뿐이면 원래 줄들을 돌려준다(빈 리스트를 주면 호출부가 IndexError).
    """
    out = [l.strip() for l in txt.split("\n") if l.strip()]
    i = 0
    while i < len(out) and PAGENUM.match(out[i]):
        i += 1
    return out[i:] or out


def norm_numeric(s):
    """숫자 비교를 위한 정규화 — 천단위 쉼표와 «진짜 마이너스»만 손댄다.

    ⛔ 원고는 U+2212(−)를 쓰고 xlsx 는 하이픈(-)을 쓴다. 정규화하지 않으면 «음수가 전부
       「문서에 없다」»로 뜬다 — 실측 2026-09-08: p2_table1 의 -3.6 이 MISSING 으로
       신고됐는데 원고에는 −3.6 이 세 곳에 있었다.
    ⛔ en-dash(–)는 «건드리지 않는다» — 범위 표기(32.6–49.2)라 하이픈으로 바꾸면
       「32.6」과 「-49.2」로 쪼개져 «없는 음수»가 생긴다.
    ⚠ 이 함수는 두 곳(문서 텍스트·xlsx 셀)에서 쓴다. 전에는 같은 정규화가 두 군데
       따로 적혀 있었고 en-dash 단서는 한 곳에만 있었다 — 한쪽만 낡는 형태였다.
    """
    return s.replace(",", "").replace("\u2212", "-")


def _sheet_rows(path):
    """모든 시트의 행을 돌려준다.

    ⚠ openpyxl 은 R `openxlsx` 가 쓴 파일에서 죽는다 -- drawing 관계는 선언돼 있는데
       `xl/drawings/drawing1.xml` 이 없어 KeyError 로 터진다(2026-09-11 실측,
       한 프로젝트의 산출 워크북 전부). 그러면 표 대조가 통째로 안 돌아가고,
       「감사 통과」라고 말할 수 없는 상태가 조용히 계속된다. calamine 을 먼저 쓰고
       openpyxl 은 폴백으로 남긴다.
    ⚠ 옛 판은 `.active`(첫 시트)만 읽었다. 이 랩의 산출 워크북은 대개 여러 시트이고
       본문 표가 첫 시트가 아닌 경우가 많다 -- 그러면 대조할 것이 없는데도 통과한다.
    """
    try:
        import pandas as pd
        for _, df in pd.read_excel(path, sheet_name=None, engine="calamine",
                                   header=None).items():
            for row in df.itertuples(index=False, name=None):
                yield row
        return
    except Exception:
        pass
    import openpyxl
    for ws in openpyxl.load_workbook(path, data_only=True).worksheets:
        for row in ws.iter_rows(values_only=True):
            yield row


def xlsx_tokens(path):
    """Numeric tokens in a workbook, excluding column 0 (row labels)."""
    out = Counter()
    for row in _sheet_rows(path):
        for j, c in enumerate(row):
            if c is None or j == 0 or (isinstance(c, float) and c != c):
                continue
            # strip thousands separators FIRST or "19,959" tokenises as 19 + 959
            for t in NUM.findall(norm_numeric(str(c))):
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
        lines_next = first_real(next_txt)   # 쪽·줄번호를 거른다 — 그 함수 머리말 참조
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
    # A page carrying only a page number or running head IS blank (2026-09-21: a footer "14" hid a
    # blank page from this check). Look at the body band only, and at images/drawings too.
    def _body_empty(p):
        H = p.rect.height
        body = [b for b in p.get_text("blocks") if b[4].strip() and H * 0.07 < (b[1] + b[3]) / 2 < H * 0.92]
        return not body and not p.get_images(full=True) and not p.get_drawings()
    blank = [i + 1 for i, p in enumerate(doc) if _body_empty(p)]
    print("4. blank pages: %s" % (blank or "none"))
    if blank:
        failures.append("blank page(s) %s" % blank)

    full = " ".join(p.get_text().replace("\n", " ") for p in doc)
    flat = norm_numeric(full)      # 근거는 norm_numeric() 머리말에 있다

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
