#!/usr/bin/env python3
"""
surface_leak_scan.py — general-purpose surface-leak / meta-commentary scanner.

Generalizes pptx_kit.check_surface_leaks (which only works on python-pptx
Presentation text frames) to any text-bearing deliverable: .tex, .md, .txt
source, rendered .pdf, or .docx (paragraphs + table cells). Same philosophy,
explicitly carried over:

  `terms` has NO default and is not itself shared: what counts as a "leak" is
  register- and project-specific. Each caller supplies its own list, informed
  by output_surface.md's two-question register test:
    1. 청중이 주제에 대해 알게 되는 것인가, 내 편집의 해명인가?
    2. 이전 버전 없이 처음 만들었어도 이 문장이 여기 있을까?
  A shared default list produces false positives in one register and misses
  in another — do not add one here.

  하나의 예외가 `--rhetoric`이고, 그건 기본값이 아니라 opt-in이다. 왜 그 하나만인지는
  rhetoric_rules()의 실측이 답한다(원고 429줄: rhetoric 오탐 0건, 나머지 4범주 30건).

Gap this closes (2026-08-10, compreg_meeting_memo_260809.tex): the pptx-only
scope of check_surface_leaks meant a LaTeX meeting memo had no automated
surface-leak gate at all — two passes of manual reading each missed real
"this document is N files merged" / "an internal audit flagged this" leaks
that a caller-supplied \\-phrase list here would have caught mechanically on
the first pass.

Usage:
    python surface_leak_scan.py FILE --terms "이 메모는" "사용자가" "이전 버전"
    python surface_leak_scan.py FILE --terms-file terms.txt
    # exits 1 and prints hits if any found; exits 0 (silent) if clean

As a library:
    from surface_leak_scan import scan_text, scan_file
    hits = scan_file("memo.tex", ["사용자가", "세션에서"])
    # -> [(line_or_page, term, snippet), ...]

For .pptx specifically, don't route through here — use the pptx-editing skill's own
`audit_surface_text.py` instead. It additionally ships a small register-agnostic baseline
(mechanical English academic-writing tells, not project-specific) on top of the same
caller-supplied-terms idea via --extra-pattern, and it also catches figure/table number drift,
which this generic scanner has no concept of.
"""
import argparse
import re
from pathlib import Path
import sys

# Windows console default codepage (cp949/cp1252) can't encode arbitrary Korean/
# accented output; reconfigure stdout/stderr to UTF-8 so this never crashes on
# what it's printing rather than what it found.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


def scan_text(text, terms, unit="line"):
    """Scan `text` for every string in `terms`. Returns [(line_no, term, snippet), ...].

    Matching is substring, case-sensitive (Korean/mixed-script text has no
    useful case-fold; if a project needs case-insensitive matching for an
    English-heavy document, lower() both sides before calling).
    """
    hits = []
    for i, line in enumerate(text.splitlines(), start=1):
        for term in terms:
            if term in line:
                snippet = line.strip()
                if len(snippet) > 100:
                    idx = snippet.find(term)
                    lo = max(0, idx - 30)
                    snippet = ("…" if lo > 0 else "") + snippet[lo:idx + len(term) + 30] + "…"
                hits.append((i, term, snippet))
    return hits


def _extract_pdf_text_by_page(path):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise SystemExit("PDF scanning needs PyMuPDF: pip install pymupdf")
    doc = fitz.open(path)
    return [page.get_text() for page in doc]


def _extract_docx_blocks(path):
    """[(location, text), ...] -- paragraphs, then table cells. Added 2026-08-11: before this,
    scan_file() opened every non-PDF path as UTF-8 text, which crashes immediately on .docx
    (a zip, not text) -- the manuscript/brief genres in agent/tools/audit.py's dispatcher were
    handing out a --terms hint that errored the moment anyone actually ran it."""
    from docx import Document
    doc = Document(path)
    out = []
    for i, p in enumerate(doc.paragraphs, start=1):
        if p.text.strip():
            out.append(("paragraph %d" % i, p.text))
    for ti, table in enumerate(doc.tables, start=1):
        for ri, row in enumerate(table.rows, start=1):
            for ci, cell in enumerate(row.cells, start=1):
                t = "\n".join(cp.text for cp in cell.paragraphs)
                if t.strip():
                    out.append(("table %d row %d col %d" % (ti, ri, ci), t))
    return out


def rhetoric_rules():
    """«수사» 규칙 하나만 형식을 넘어 공유한다 — 정본은 pptx 스킬에 있고 여기선 «빌려 쓴다».

    2026-08-20 실측이 이 경계를 정했다. audit_surface_text의 5범주를 실제 원고 429줄에 그대로
    걸어보니 meta 25건·nav 1건·provenance 4건이 **전부 정상 산문**이었다("Sensitivity analyses
    addressed three threats…"는 원고의 표준 문장이고, 상호참조와 출처 표기도 마찬가지다).
    **rhetoric만 0건**이었다.
    → 그래서 통짜 탑재는 정확도를 «떨어뜨린다». 슬라이드에서 결함인 말이 원고에서는 정상이고,
      우는 게이트는 곧 꺼져서 신호가 0이 된다. 공유할 것은 **보편적인 한 범주뿐**이다.

    이 파일의 «기본 목록 없음» 설계는 그대로다 — 이건 기본값이 아니라 `--rhetoric` opt-in이다.
    """
    import os
    for cand in (os.path.join(os.path.expanduser("~"), ".claude", "skills", "pptx-editing", "scripts"),):
        if os.path.isdir(cand):
            sys.path.insert(0, cand)
            break
    import audit_surface_text as A
    return A.DEFAULT["rhetoric"], A.caps_emphasis


def _extract_hwpx_blocks(path):
    """[(location, text), ...] for .hwpx -- 정부·기관 보고서.

    추출기를 새로 쓰지 않고 `audit_text_consistency.read_hwpx()`를 재사용한다. 그쪽이 이미
    `hwpxlib`로 표 셀의 **문단 경계를 보존**해서 읽고, 그 규칙(셀 안 여러 문단 = 줄바꿈)은
    두 도구가 같아야 한다 -- 사본을 두면 조용히 갈린다.

    Added 2026-08-20. 그 전까지 hwpx는 «표면 규율» 축에 검사가 아예 없었다 --
    [[report-content-discipline]]이 통째로 그 얘기인데도. 좌표(check × DOC)로 보면
    한 칸이 비어 있던 것이고, 형식이 하나 빠진 것을 «없다»고 말할 방법이 없었다.
    """
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from audit_text_consistency import read_hwpx
    body, cells = read_hwpx(path)          # (본문 문단들, 표 셀들)
    out = []
    for i, t in enumerate(body, start=1):
        if t.strip():
            out.append(("paragraph %d" % i, t))
    for i, t in enumerate(cells, start=1):
        if t.strip():
            out.append(("table cell %d" % i, t))
    return out


def _extract_xlsx_blocks(path):
    """[(location, text), ...] -- 시트명과 «열 제목»만. 셀 값은 읽지 않는다.

    왜 값을 안 읽나. 이 도구가 보는 것은 «사람이 읽는 표면»이고, 표에서 그것은 시트명과
    머리행이다. 값의 정합은 파이프라인의 숫자 게이트가 본다. 값까지 넣으면 오탐이 폭발해
    도구가 소음이 되고, 소음이 된 게이트는 아무도 안 본다.

    Gap this closes (2026-09-02, 한 원고): 표의 열 이름이 «다른 양»을 가리키는 것을
    (「변화 중앙값」인데 값은 중앙값의 차) 어떤 자동 검사도 못 봤다. 원고·범례·PDF 는 검사
    대상인데 표만 빠져 있었고, 그 결함은 사람이 정독해서야 나왔다.
    """
    try:
        import openpyxl
    except ImportError:
        return [("", "")]                       # 없으면 조용히 건너뛴다 -- 다른 형식은 계속 돈다
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for name in wb.sheetnames:
        out.append(("sheet name", name))
        ws = wb[name]
        hdr = next(ws.iter_rows(max_row=1, values_only=True), ())
        cells = [str(c) for c in hdr if c is not None and str(c).strip()]
        if cells:
            out.append(("%s: header" % name, " | ".join(cells)))
    wb.close()
    return out or [("", "")]


def blocks_of(path):
    """[(location, text), ...] for any supported format -- 읽기와 «규칙»을 가르는 지점.

    scan_file()이 형식별로 흩어져 있던 추출 분기를 여기로 모았다. 규칙(--terms, --rhetoric)은
    이제 블록만 받으므로, 새 규칙을 붙일 때 형식 분기를 다시 쓰지 않는다.
    """
    low = path.lower()
    if low.endswith(".hwpx"):
        return _extract_hwpx_blocks(path)
    if low.endswith(".docx"):
        return _extract_docx_blocks(path)
    if low.endswith(".pptx"):
        # 덱도 «산출물»이다. --rhetoric(AI 티)은 형식을 가리지 않으므로 여기서도 돌아야 한다.
        # 이 갈래가 없으면 .pptx가 평문 갈래로 떨어져 UnicodeDecodeError를 낸다(zip이라서).
        # 정본 추출기는 pptx 스킬에 있다 -- 사본을 만들지 않는다.
        rhetoric_rules()                                    # 스킬 경로를 sys.path에 올린다
        from audit_surface_text import paragraphs
        return [("slide %d" % si, t) for si, t in paragraphs(path)]
    if low.endswith(".xlsx"):
        # 표도 «산출물»이다. 시트명과 열 제목은 독자가 읽는 표면이고, 실측상 가장 자주
        # 검사에서 빠지는 자리다(2026-09-02). ⛔ 셀 값은 읽지 않는다 -- 위 함수 머리말 참조.
        return _extract_xlsx_blocks(path)
    if low.endswith(".pdf"):
        return [("p%d" % n, t) for n, t in enumerate(_extract_pdf_text_by_page(path), start=1)]
    with open(path, encoding="utf-8") as f:
        return [("", f.read())]


def scan_file(path, terms):
    """Scan a file for `terms`. .pdf gets per-page text extraction (location = page
    number); .docx/.hwpx get per-paragraph/per-cell extraction (location = paragraph or
    table-cell label); everything else is read as UTF-8 text (location = line number).
    Returns [(location, term, snippet), ...].
    """
    hits = []
    for loc, block_text in blocks_of(path):
        for line_no, term, snippet in scan_text(block_text, terms):
            if not loc:                                    # 평문: 위치 = 줄번호
                where = line_no
            elif line_no == 1 and "\n" not in block_text:  # 한 줄짜리 블록: 라벨만
                where = loc
            else:
                where = "%s, line %d" % (loc, line_no)
            hits.append((where, term, snippet))
    return hits


def scan_rhetoric(path):
    """수사(rhetoric) 규칙을 형식과 무관하게 적용한다. [(location, term, snippet), ...]."""
    pattern, caps_emphasis = rhetoric_rules()
    rx = re.compile(pattern, re.I)
    blocks = blocks_of(path)
    hits, flat = [], []
    for loc, block_text in blocks:
        for line_no, line in enumerate(block_text.splitlines(), start=1):
            if not line.strip():
                continue
            where = line_no if not loc else ("%s, line %d" % (loc, line_no) if "\n" in block_text else loc)
            flat.append((where, line.strip()))
            m = rx.search(line)
            if m:
                hits.append((where, m.group(0).strip(), line.strip()[:100]))
    # 대문자 강조는 «문서 전체»를 봐야 판정된다 -- 같은 낱말의 소문자형이 어딘가 있으면 약어가 아니다
    for _tag, where, word, why in caps_emphasis(flat):
        hits.append((where, word, why))
    return hits


def scan_premise(path):
    """세울 필요 없는 것을 세운 자리. [(location, term, snippet), ...].

    규칙은 `audit_text_consistency.py` 의 STRAWMAN / PERSON_PREMISE 를 «가져다» 쓴다.
    복제하지 않는다 — 같은 규칙이 두 파일에 손으로 복사되면 한쪽만 고쳐져 조용히 어긋난다.

    왜 여기 있나 (2026-09-05). audit.py 는 text_consistency 를 manuscript/report/text 에만
    돌린다 — slide·poster 는 그 검사를 아예 안 받고 있었다. 그런데 이 실패는 슬라이드에서
    더 세다: 불릿은 시각 앵커라 「A가 아니라 B」를 띄우면 독자 눈에 A 가 남는다.
    이 파일은 blocks_of() 가 .pptx 를 이미 읽으므로 여기가 제자리다.

    ⚠ 「A가 아니라 B」가 늘 잘못은 아니다 — A 가 이미 무대에 있으면 정당한 대조다.
      원고에서는 «경쟁 가설»(심사자가 제기할 것)이, 수업 자료에서는 «학생이 실제로 가진
      오개념»이 그렇다. 그래서 부정당한 말이 산출물 다른 곳에도 있는지를 함께 본다.
    """
    import importlib.util
    here = Path(__file__).resolve().parent
    spec = importlib.util.spec_from_file_location("atc", here / "audit_text_consistency.py")
    atc = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(atc)

    blocks = blocks_of(path)
    whole = " ".join(t for _, t in blocks)
    hits = []
    for loc, block_text in blocks:
        for line_no, line in enumerate(block_text.splitlines(), start=1):
            for pat in atc.STRAWMAN:
                for m in re.finditer(pat, line):
                    if len(re.findall(re.escape(m.group(1)), whole)) <= 1:
                        hits.append((loc or line_no, "세운 적 없는 것을 부정",
                                     line.strip()[:80]))
            for pat, label in atc.PERSON_PREMISE:
                for _ in re.finditer(pat, line):
                    hits.append((loc or line_no, label, line.strip()[:80]))
    return hits



def coined_rules():
    """«내가 지어낸 국문 표현» — 대화에서 만든 말이 산출물로 새는 것.

    왜 이것만 형식·프로젝트를 넘어 공유되나 (2026-09-11, 한 위원회 문서):
      세션 안에서 나와 사용자는 맥락을 쌓으며 말을 만든다. 그 말은 «받는 사람에게 아무 근거가
      없다». 하루에 넷이 샜고, 그중 「한 걸음 두다」는 **한국어에 없는 결합**이었다. 연구자:
      *"'한 걸음' 정말 어색한 한국어다. 너 왜그래."* · *"좀더 공부하고 와."*
      이것은 register 문제가 아니라 «없는 말을 썼다»는 문제라 어느 국문 산출물에서든 결함이다.

    ⚠ 좁게 유지한다. 이 파일의 「기본 목록 없음」 설계를 깨지 않는 것은 opt-in 이기 때문이고,
      rhetoric 이 그랬듯 «실측으로» 자리를 얻어야 한다. 늘릴 때는 오탐을 재고 늘릴 것.
    ⛔ 「완전히·전혀·모두」 같은 절대 수량어는 넣지 않았다 — 정상 산문에 흔해 오탐이 쏟아진다.
      그건 사람이 본다(korean_prose_grafting.md §6).
    """
    return {
        "한 걸음 두": "한국어에 없는 결합. 「덧붙였다」·「연결고리를 놓았다」",
        "조건을 다는": "이 대화에서만 통하는 말. 무엇을 하는 문단인지 그대로 쓴다",
        "가리키는 한 문장": "압축이 심해 뜻이 안 온다",
        "순서를 내": "「순서를 내다」는 어색한 결합. 「순위가 달랐다」",
    }


def scan_coined(path):
    """[(location, term, snippet), ...] — 지어낸 표현."""
    hits = []
    for loc, text in blocks_of(path):
        for term, why in coined_rules().items():
            i = text.find(term)
            if i >= 0:
                hits.append((loc, "%s  → %s" % (term, why),
                             text[max(0, i - 30):i + len(term) + 30].replace("\n", " ")))
    return hits


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", nargs="+",
                    help="Files or globs to scan (.tex/.md/.txt/.pdf/.docx/.hwpx/.pptx/.xlsx). "
                         "여러 개를 주면 «표면 집합»으로 한 번에 검사한다 -- 게이트가 보는 "
                         "범위를 독자가 보는 범위에 맞추기 위해서다(2026-09-02).")
    ap.add_argument("--terms", nargs="+", help="Banned phrases (space-separated, quote each)")
    ap.add_argument("--terms-file", help="One banned phrase per line")
    ap.add_argument("--rhetoric", action="store_true",
                    help="«수사» 규칙(카드 제목이 수사적 기능·화살표를 말로·문장 속 대문자 강조)도 함께 검사. "
                         "5범주 중 이것만 형식·register를 넘어 안전하다 -- rhetoric_rules() 참조")
    ap.add_argument("--coined", action="store_true",
                    help="내가 지어낸 국문 표현 (대화용 말이 산출물로 샌 것). "
                         "국문 산출물이면 형식 무관 -- coined_rules() 참조")
    ap.add_argument("--premise", action="store_true",
                    help="«세울 필요 없는 것을 세운 자리»도 검사 — 「A가 아니라 B」에서 A 가 "
                         "산출물 다른 곳에 없거나, 읽는 사람을 규정하는 전제. scan_premise() 참조")
    args = ap.parse_args()

    terms = list(args.terms or [])
    if args.terms_file:
        with open(args.terms_file, encoding="utf-8") as f:
            terms += [ln.strip() for ln in f if ln.strip()]
    if not terms and not args.rhetoric and not args.premise and not args.coined:
        raise SystemExit("no --terms/--terms-file/--rhetoric/--premise/--coined given — nothing to scan for")

    # ⭐ 표면 «집합»으로 돈다. 게이트가 보는 범위를 독자가 보는 범위에 맞추려면 한 파일이
    #    아니라 그 산출물이 내보내는 표면 전부를 한 번에 봐야 한다(2026-09-02).
    import glob as _glob
    paths = []
    for pat in args.file:
        got = sorted(_glob.glob(pat))
        paths.extend(got if got else [pat])          # 글롭이 안 맞으면 그대로 둔다(없으면 아래서 걸린다)

    total, scanned = 0, []
    for path in paths:
        try:
            hits = scan_file(path, terms) if terms else []
            if args.rhetoric:
                hits += scan_rhetoric(path)
            if args.coined:
                hits += scan_coined(path)
            if args.premise:
                hits += scan_premise(path)
        except FileNotFoundError:
            print(f"  없음: {path}")
            continue
        scanned.append(path)
        if hits:
            print(f"[surface leaks] {path}")
            for loc, term, snippet in hits:
                print(f"  {loc}: '{term}' in \"{snippet}\"")
            total += len(hits)

    kinds = sorted({Path(p_).suffix.lstrip(".") or "txt" for p_ in scanned})
    print(f"--- 표면 {len(scanned)}개({', '.join(kinds)}) x {len(terms)} terms"
          f"{', rhetoric' if args.rhetoric else ''}"
          f"{', coined' if args.coined else ''} -> 위반 {total}")
    if total:
        sys.exit(1)


if __name__ == "__main__":
    main()
