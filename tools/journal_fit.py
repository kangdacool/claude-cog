#!/usr/bin/env python3
"""
journal_fit.py — how a draft compares with real papers from the TARGET JOURNAL, in section
LENGTHS and in prose STYLE, from a single fetch of the corpus.

WHY THIS EXISTS (2026-08-16). A manuscript had to come down from 6,681 words to
a 4,000-word limit. A dozen compression passes were spent tightening prose uniformly across
every section, each returning 3-8% and each costing something real — a citation-fairness
contrast, a defence of a pre-registered gate, three cross-references to content that no
longer existed. The question that ended it was "how do actual papers in this journal do it?",
and the answer took one measurement:

    24 AJE methodology papers, PMC full text, the publisher's own section tags
        Introduction 647 (16%)   Methods 1553 (39%)
        Results      618 (16%)   Discussion 1052 (26%)   BODY 3973

The draft was Results-heavy at 37% against a 16% norm. Rewriting that one section took the
body from 4,569 to 3,928 — more than every prior pass combined, without touching a finding.
**Uniform compression is the wrong instrument. The question is which section is out of shape.**

The same corpus then showed 8 of 10 style features outside the range the journal spans, the
em-dash rate at nearly twice the most dash-happy paper.

Run it BEFORE the first cut, not after the tenth.

TWO CAUTIONS, both learned on the run this was built for:

  1. Only what falls OUTSIDE the observed range is flagged. Editing toward a median produces
     prose no human wrote either, and the user's instruction was explicit: take only what is
     worth taking, and taking nothing is a valid outcome.
  2. Fixing a flagged feature can MOVE it rather than remove it. Replacing every em-dash
     pushed 'rather than' into 'X, not Y' and the antithesis total went UP. Re-measure after
     editing, and change the construction, not the character.

AFTER MEASURING — WHAT TO CUT, IN THIS ORDER (한 원고 2026-09-03)
Measuring names the section. It does not name the sentences. Doing that out of order costs
findings. This order took a Results section from 1,335 to 1,126 words without losing a number:

  0. READ IT FIRST. Duplication and contradiction are defects, not compression targets. Two
     consecutive sentences repeated the same interaction with the same figures (0.96, P=0.921;
     3.73, P=0.037) — 55 words no compression pass would have found, because each sentence was
     individually fine.
  1. Measure the TARGET journal and one ABOVE it; cut only where both point. Here Discussion
     ran long against the target and was normal against the upper tier — that tier writes long
     discussions — so Results was the only section genuinely out of shape.
  2. Split any aggregate style metric BY SECTION before acting on it. Hedging read far below
     both journals; by section, Discussion sat at the median and the low aggregate was Methods
     and Results, where hedging is wrong. The aggregate would have blurred the flat section.
  3. Inside the flagged section, cut what a TABLE OR FIGURE already carries — this is lossless.
     Wave-by-wave counts that table 1 and figure 1 hold; six item-level estimates a figure and
     a supplementary table hold. Replace with a range or a pointer, never with more prose.
  4. Only then cut a BRANCH THAT CARRIES NO CLAIM. Test: if this result vanished, is the
     argument weaker? A borderline secondary outcome (p=0.080) whose job the negative control
     already did failed that test and came out whole.
  5. Whatever you cut, KEEP THE ANALYSIS — delete it from the manuscript, not the pipeline, and
     leave a comment naming every place that must be restored together. It is the answer when a
     reviewer asks the question it was built for.
  6. Cutting an item BREAKS THE SENTENCES THAT COUNT IT. "Three observations bear on it" became
     "Two". Gates pass a sentence that has silently become false — fix counts and connectives
     in the same edit.
  7. If you ADD, PAY FOR IT in the same pass. 175 words of new argument put Discussion back out
     of shape; that is a debt, not a windfall.

Usage:
    python journal_fit.py --journal "Am J Epidemiol" --terms "methodology OR simulation" \\
        --draft MS.md --start "## Introduction" --end "## Display items"
    python journal_fit.py --journal "Stat Med" --n 40 --draft MS.md --only shape

Supersedes journal_shape.py and journal_style.py, which fetched the same corpus separately.
"""
import argparse
import io
import json
import re
import sys
import urllib.parse
import urllib.request

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
## ⚠ 2026-09-08 — stderr 도 같이 바꿈. NOTE 를 stderr 로 보내는데 호출측이
##    encoding="utf-8" 로 캐처하면 cp949 바이트에서 UnicodeDecodeError 가 난다.
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
UA = {"User-Agent": "Mozilla/5.0 (research use)"}
SECT = [("INTRO", "Introduction"), ("METHODS", "Methods"),
        ("RESULTS", "Results"), ("DISCUSS", "Discussion")]
BODY_SECTIONS = ("INTRO", "METHODS", "RESULTS", "DISCUSS")


def _get(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60).read()


def pmids(journal, terms, n):
    q = '"%s"[jour] AND "pubmed pmc open access"[filter]' % journal
    if terms:
        q += " AND (%s)" % terms
    url = ("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed"
           "&retmode=json&sort=date&retmax=%d&term=%s" % (n, urllib.parse.quote(q)))
    return json.loads(_get(url))["esearchresult"]["idlist"]


def measure(pmid):
    """(abstract words, {section: words}, body words, body text) or None."""
    try:
        doc = json.loads(_get("https://www.ncbi.nlm.nih.gov/research/bionlp/RESTful/"
                              "pmcoa.cgi/BioC_json/%s/unicode" % pmid))
    except Exception:
        return None
    if not isinstance(doc, list) or not doc:
        return None
    d = doc[0]["documents"][0]
    counts = {k: 0 for k, _ in SECT}
    abs_w, paras = 0, []
    for p in d["passages"]:
        inf = p.get("infons", {})
        typ = (inf.get("section_type") or "").upper()
        txt = p.get("text", "") or ""
        if typ == "ABSTRACT":
            abs_w += len(txt.split())
        if typ in counts and inf.get("type") in ("paragraph", "abstract"):
            counts[typ] += len(txt.split())
        if typ in BODY_SECTIONS and inf.get("type") == "paragraph":
            paras.append(txt)
    body = sum(counts.values())
    if body < 1500:                       # not a research article, or full text untagged
        return None
    return abs_w, counts, body, "\n\n".join(paras)


def features(text):
    words = text.split()
    n = len(words) or 1
    sents = [s for s in re.split(r"(?<=[.!?])\s+", text) if len(s.split()) > 2]
    paras = [p for p in text.split("\n\n") if len(p.split()) > 20]

    def per10k(count):
        return round(10000.0 * count / n, 1)

    def med_len(items):
        v = sorted(len(x.split()) for x in items)
        return v[len(v) // 2] if v else 0

    return {
        "words": n,
        "median sentence (words)": med_len(sents),
        "long sentences >40w (%)": round(100.0 * sum(len(s.split()) > 40 for s in sents)
                                         / max(1, len(sents))),
        "median paragraph (words)": med_len(paras),
        # ⚠ `--` 를 맨몸으로 세면 마크다운 구분선 `---` 과 표 구분행 `|---|` 이 전부 잡힌다.
        #    2026-08-23 실측: 같은 초안이 도구 103.5 /10k · 산문 실제 20.8 /10k 였다.
        #    산문의 이중 하이픈은 «공백에 둘러싸여» 있다.
        "em-dash /10k": per10k(len(re.findall(r"—|\s--\s", text))),
        "semicolon /10k": per10k(text.count(";")),
        "'rather than' /10k": per10k(len(re.findall(r"\brather than\b", text, re.I))),
        "'X, not Y' /10k": per10k(len(re.findall(r",\s+not\s+\w", text, re.I))),
        "first-person we /10k": per10k(len(re.findall(r"\bwe\b|\bour\b", text, re.I))),
        "hedges /10k": per10k(len(re.findall(
            r"\bmay\b|\bmight\b|\bsuggest\w*\b|\bappear\w*\b|\blikely\b|\bapproximately\b",
            text, re.I))),
        "numeral-opener /10k": per10k(len(re.findall(
            r"(?:^|(?<=[.!?]\s))(One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten)\s", text))),
    }


def med(vals):
    v = sorted(vals)
    return v[len(v) // 2] if v else 0


def quart(vals):
    v = sorted(vals)
    k = len(v)
    return v[k // 4], v[k // 2], v[(3 * k) // 4], v[0], v[-1]


def as_markdown(path):
    """Return the draft as markdown-ish text.

    A .docx is accepted as well as a .md so this can be pointed at the built
    manuscript, which is the artefact that actually goes out. Word headings
    carry their level in the style name, so they map onto '#'/'##' exactly;
    everything else is a paragraph.
    """
    if not path.lower().endswith(".docx"):
        return io.open(path, encoding="utf-8").read()
    from docx import Document
    out = []
    for p in Document(path).paragraphs:
        name = (p.style.name or "")
        m = re.match(r"Heading (\d)", name)
        out.append(("#" * int(m.group(1)) + " " + p.text.strip()) if m else p.text)
    ## 문단 사이는 «빈 줄»로 잇는다. 하류가 split("\n\n") 으로 문단을 가르므로
    ## "\n" 으로 이으면 docx 의 문단 경계가 전부 사라져, 절 하나가 한 문단으로 잡힌다
    ## (2026-09-05 실측: 172어 초과 문단이 0개인 원고를 median 208 로 보고했다).
    return "\n\n".join(out)


def own_shape(path):
    """Word count per IMRaD section of the draft.

    Heading level is not fixed: a manuscript that uses `# Methods` is as common
    as one that uses `## Methods`. This used to match `##` only, and a draft
    written with `#` silently measured 0 words in every section - which the
    report then printed as "out of shape" by the full length of the paper. A
    check that returns zero because it did not look is worse than no check, so
    both levels are tried and a total miss is reported as a miss.
    """
    t = as_markdown(path)

    def find(label):
        """Start offset of a heading, at any level, or None.

        The section NUMBER is optional. The lab's own manuscript style writes
        "## 1. Introduction", and matching only the bare label made every in-house draft
        report all four sections as 0 -- and therefore as "out of shape" by their entire
        length, a deficit that reads like a finding rather than a miss
        (2026-08-29, polypharm_mort). A trailing subtitle after a colon or dash is
        allowed for the same reason.
        """
        for prefix in ("#### ", "### ", "## ", "# "):
            m = re.search(r"^%s(?:\d+[.)]?\s+)?%s\s*(?:[:—-].*)?$"
                          % (re.escape(prefix), re.escape(label)), t, re.MULTILINE)
            if m:
                return m.start()
        return None

    def w(a, b):
        i, j = find(a), find(b)
        if i is None or j is None or j <= i:
            return 0
        s = re.sub(r"<sup>.*?</sup>", " ", t[i:j])
        return len(re.sub(r"[*_`#>|]", " ", s).split())

    # Discussion runs to whatever comes first of the usual back-matter headings.
    # "Conclusion(s)" belongs in this list: many journals carry it as its own short
    # section after Discussion (3 of 6 sampled JECH papers do), and a draft that ends
    # there used to leave Discussion measured as 0 and then reported as out of shape by
    # its entire length - the same failure the heading-level fix above was written for.
    tail = next((h for h in ("Conclusion", "Conclusions", "Display items",
                             "Ethics approval", "References", "Acknowledgements")
                 if find(h) is not None), "References")
    out = {"INTRO": w("Introduction", "Methods"),
           "METHODS": w("Methods", "Results"),
           "RESULTS": w("Results", "Discussion"),
           "DISCUSS": w("Discussion", tail)}
    if not any(out.values()):
        print("  !! no IMRaD headings found in the draft (tried '# X' and '## X' for\n"
              "     Introduction/Methods/Results/Discussion) - shape NOT measured.\n")
    elif not all(out.values()):
        # A single zero is the dangerous case: the report prints it as a deficit the
        # size of the whole section, and that reads like a finding rather than a miss.
        miss = [k for k, v in out.items() if not v]
        print("  !! %s measured as 0 words - almost certainly NOT FOUND rather than\n"
              "     absent. Check the heading spelling and what follows the section;\n"
              "     do not read the deficit below as a finding.\n" % ", ".join(miss))
    return out


def strip_working_text(t):
    """Remove what an author would never submit: HTML working comments and non-Latin notes.

    Why (2026-08-24, 한 원고): a manuscript source carried Korean working notes inside
    `<!-- ... -->`. Counting them put em-dash at 29.7/10k and reported "ABOVE every paper" when
    the body held exactly one em-dash, itself inside a comment. Semicolons, quote marks and
    sentence length are contaminated the same way. Returns (clean_text, dropped_fraction).
    """
    n0 = len(t)
    t = re.sub(r"<!--.*?-->", " ", t, flags=re.S)
    keep = []
    for line in t.split("\n"):
        cjk = sum(1 for ch in line if "\uac00" <= ch <= "\ud7a3" or "\u4e00" <= ch <= "\u9fff")
        if cjk:
            continue          # any CJK on the line means a working note, not submitted prose
        keep.append(line)
    return "\n".join(keep), (1.0 - len("\n".join(keep)) / float(n0 or 1))


## ⭐ 2026-09-08 — --start/--end 가 «없으면» 본문 구간을 스스로 잡는다.
##    그 전에는 문서 «전체»를 쟀다. audit.py 가 .docx 를 통째로 넘기면 제목면·초록·표·
##    참고문헌이 다 섞여 들어가 세미콜론이 37 대신 94 로, 문장 길이가 21 대신 15 로 나왔다.
##    ⛔ 틀린 숫자를 «조용히» 보여주는 것이 안 보여주는 것보다 나쁘다 — 그래서 무엇을 쟀는지
##       stderr 에 인쇄한다.
##    ⚠ 구조화 초록이 Background/Methods/Results/Conclusions 를 «먼저» 반복하므로 머리를
##       첫 등장으로 잡으면 초록이 본문으로 딸려 온다. Methods «앞»에서 뒤로 훑어 잡는다.
TAIL = ("\nReferences", "\nDeclarations", "\nList of abbreviations",
        "\nAcknowledgements", "\nManuscript tables", "\nFigure legends")
HEAD = ("\nIntroduction", "\nBackground")


def auto_span(t):
    """(잘린 텍스트, 설명) — 못 잡으면 원문 그대로."""
    end_i = min([t.find(k) for k in TAIL if t.find(k) > 0] or [len(t)])
    body = t[:end_i]
    m_i = body.rfind("\nMethods")
    if m_i > 0:
        cand = [body.rfind(h, 0, m_i) for h in HEAD]
        h_i = max(cand)
        if h_i > 0:
            return body[h_i:], "auto: %s … %s" % (
                body[h_i:h_i + 14].strip(), ("cut at " + t[end_i:end_i + 14].strip())
                if end_i < len(t) else "end")
    return body, ("auto: whole document" if end_i >= len(t)
                  else "auto: … cut at %s" % t[end_i:end_i + 14].strip())


def draft_body(path, start, end):
    t = as_markdown(path)
    if start and start in t:
        t = t[t.index(start):]
    elif end is None and start is None:
        t, why = auto_span(t)
        sys.stderr.write("  NOTE body span %s (%d words). --start/--end 로 덮어쓸 수 있다.\n"
                         % (why, len(t.split())))
    if end and end in t:
        t = t[:t.index(end)]
    t, dropped = strip_working_text(t)
    if dropped > 0.05:
        sys.stderr.write(
            "  NOTE %.0f%% of the input was working notes (HTML comments / CJK lines) and was "
            "not measured.\n       Prefer the assembled submission copy.\n" % (100 * dropped))
    t = re.sub(r"<sup>.*?</sup>", " ", t)
    t = re.sub(r"^#+ .*$", "", t, flags=re.M)
    return re.sub(r"[*_`|]", "", t)


def report_shape(rows, draft_path):
    b = med([r[2] for r in rows])
    mine = own_shape(draft_path)
    print("SECTION SHAPE")
    print("  %-14s %8s %7s %9s %8s" % ("", "median", "share", "yours", "vs"))
    for key, label in SECT:
        m = med([r[1][key] for r in rows])
        diff = mine[key] - m
        flag = "  <-- out of shape" if m and abs(diff) > 0.5 * m else ""
        print("  %-14s %8d %6d%% %9d %+8d%s"
              % (label, m, round(100.0 * m / b), mine[key], diff, flag))
    total = sum(mine.values())
    print("  %-14s %8d %7s %9d %+8d" % ("BODY", b, "", total, total - b))
    print("  abstract median %d words\n" % med([r[0] for r in rows]))


def corpus_stats(rows):
    """코퍼스의 특징별 (q1, median, q3, min, max) — 저장·재사용 가능한 형태."""
    feats = [features(r[3]) for r in rows]
    out = {}
    for k in feats[0]:
        if k == "words":
            continue
        q1, md, q3, lo, hi = quart([f[k] for f in feats])
        out[k] = {"q1": q1, "median": md, "q3": q3, "min": lo, "max": hi}
    return {"n": len(feats), "features": out}


def report_style(rows, draft_path, start, end, stats=None):
    """rows 대신 stats(저장된 코퍼스)로도 채점한다. 반환: 범위를 벗어난 특징 목록."""
    if stats is None:
        stats = corpus_stats(rows)
    S = stats["features"]
    mine = features(draft_body(draft_path, start, end))
    print("PROSE STYLE  (draft body %d words, corpus n=%d)" % (mine["words"], stats["n"]))
    print("  %-26s %7s %7s %7s %9s" % ("feature", "Q1", "median", "Q3", "yours"))
    outside = []
    for k in mine:
        if k == "words" or k not in S:
            continue
        s = S[k]
        mark = ""
        if mine[k] > s["max"]:
            mark = "  ABOVE every paper (max %g)" % s["max"]
            outside.append(k)
        elif mine[k] < s["min"]:
            mark = "  BELOW every paper (min %g)" % s["min"]
            outside.append(k)
        elif mine[k] > s["q3"] or mine[k] < s["q1"]:
            mark = "  outside the middle half"
        print("  %-26s %7g %7g %7g %9g%s"
              % (k, s["q1"], s["median"], s["q3"], mine[k], mark))
    print()
    if outside:
        print("Outside the range entirely: %s" % ", ".join(outside))
        print("Change the construction, not the character, and re-measure: a fix can move a")
        print("tell into another feature instead of removing it.")
        ## ⭐ 2026-09-08 — audit.py 가 «줄머리 [의심]»을 세어 [주의] 로 올린다.
        ##    이걸 안 찍으면 검사가 돌고도 [PASS] 로 찍혀 결과가 사라진다 — 실제로 그랬다.
        ##    ⛔ exit code 는 바꾸지 않는다(문체 이상치는 결함이 아니고, 늘 FAIL 인 검사는
        ##       읽히지 않는다). 「사람이 볼 후보」로만 올린다.
        for k in outside:
            s = S[k]
            print("  [의심] %s = %g — 코퍼스 범위 밖 (min %g / max %g). 가드 문장인지 먼저 볼 것"
                  % (k, mine[k], s["min"], s["max"]))
    else:
        print("Every feature falls inside the range these papers span.")
    return outside


def _report_style_legacy(rows, draft_path, start, end):
    feats = [features(r[3]) for r in rows]
    mine = features(draft_body(draft_path, start, end))
    print("PROSE STYLE  (draft body %d words)" % mine["words"])
    print("  %-26s %7s %7s %7s %9s" % ("feature", "Q1", "median", "Q3", "yours"))
    outside = []
    for k in mine:
        if k == "words":
            continue
        q1, _, q3, lo, hi = quart([f[k] for f in feats])
        mark = ""
        if mine[k] > hi:
            mark = "  ABOVE every paper (max %g)" % hi
            outside.append(k)
        elif mine[k] < lo:
            mark = "  BELOW every paper (min %g)" % lo
            outside.append(k)
        elif mine[k] > q3 or mine[k] < q1:
            mark = "  outside the middle half"
        print("  %-26s %7g %7g %7g %9g%s"
              % (k, q1, med([f[k] for f in feats]), q3, mine[k], mark))
    print()
    if outside:
        print("Outside the range entirely: %s" % ", ".join(outside))
        print("Change the construction, not the character, and re-measure: a fix can move a")
        print("tell into another feature instead of removing it.")
    else:
        print("Every feature falls inside the range these papers span.")


def main():
    ap = argparse.ArgumentParser()
    ## --corpus 를 주면 저널명이 캐시 안에 있으므로 필수가 아니다(2026-09-08).
    ap.add_argument("--journal", help='e.g. "Am J Epidemiol" — --corpus 를 쓰면 생략 가능')
    ap.add_argument("--terms", default="", help='e.g. "methodology OR simulation"')
    # 25 로 두었다가 실제 측정 22편짜리 코퍼스가 만들어졌고, 그 분위수가 판정을 두
    # 군데서 뒤집었다(2026-09-12). 검색 수를 넉넉히 잡는다 -- OA 비율이 낮아 검색 수의
    # 절반 이하만 측정된다.
    ap.add_argument("--n", type=int, default=300,
                    help="검색할 논문 수. 이 중 OA 전문이 열리는 것만 측정된다")
    # Positional so audit.py can call it the way it calls every other check
    # (tool.py FILE ...); --draft kept for the existing invocations.
    ap.add_argument("draft", nargs="?", help="draft .md or the built .docx")
    ap.add_argument("--draft", dest="draft_opt", help="same path, as a flag")
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--only", choices=["shape", "style"], help="default: both")
    ## ⭐ 2026-09-08 — «빌드 게이트»로 쓰기 위해 셋을 더했다.
    ##    이 도구는 매 실행 PubMed 에서 코퍼스를 새로 받는다. 느리고 네트워크에 의존해
    ##    파이프라인 게이트가 될 수 없었고, 그래서 «측정은 있는데 아무도 안 보는» 상태가 됐다.
    ##    실측 사고: 문체 이상치 하나(‘X, not Y’ 가 코퍼스 최대의 3.8배)가 측정돼 있었는데
    ##    투고 «결정 시점»에 사람 앞에 놓이지 않았다. 탐지가 아니라 «게시»의 실패였다.
    ap.add_argument("--save-corpus", metavar="PATH",
                    help="측정한 코퍼스 분포를 JSON 으로 저장(한 번만 받아 두고 재사용)")
    ap.add_argument("--corpus", metavar="PATH",
                    help="저장된 코퍼스로 채점 — 네트워크를 쓰지 않는다(게이트용)")
    ap.add_argument("--fail-outside", action="store_true",
                    help="범위를 벗어난 특징이 하나라도 있으면 exit 1 (게이트용)")
    a = ap.parse_args()
    if not a.journal and not a.corpus:
        ap.error("--journal 이 필요하다 (또는 --corpus 로 저장된 코퍼스를 줄 것)")
    a.draft = a.draft or a.draft_opt
    if not a.draft:
        ap.error("give a draft: a positional path, or --draft PATH")

    ## 저장된 코퍼스로 채점 — 오프라인 경로
    if a.corpus:
        with io.open(a.corpus, encoding="utf-8") as fh:
            stats = json.load(fh)
        print("%s — cached corpus (n=%d, saved %s)\n"
              % (stats.get("journal") or a.journal or "?", stats["n"], stats.get("saved", "?")))
        outside = report_style(None, a.draft, a.start, a.end, stats=stats)
        return 1 if (a.fail_outside and outside) else 0

    ids = pmids(a.journal, a.terms, a.n)
    rows = [r for r in (measure(p) for p in ids) if r]
    if not rows:
        print("no open-access full text matched — widen --terms or raise --n")
        return 1

    if len(rows) < 50:
        print("\n⚠ 코퍼스가 %d편뿐이다 — 분위수를 믿지 말 것. 얇은 코퍼스는 «없는 문제»를\n"
              "   만들고 «있는 문제»를 가린다(2026-09-12 실측: n=22 가 em-dash Q3 를 1.7 로,\n"
              "   first-person Q1 을 25.6 으로 잡았다. n=426 에서는 3.8 과 33.5 였다).\n"
              "   --n 을 올리거나 --start 를 앞당겨 최소 50편을 받을 것.\n" % len(rows))

    print("%s — %d open-access articles measured (of %d searched)\n"
          % (a.journal, len(rows), len(ids)))
    if a.only != "style":
        report_shape(rows, a.draft)
    outside = []
    if a.only != "shape":
        outside = report_style(rows, a.draft, a.start, a.end)

    if a.save_corpus:
        import datetime
        stats = corpus_stats(rows)
        stats["journal"] = a.journal
        stats["saved"] = datetime.date.today().isoformat()
        with io.open(a.save_corpus, "w", encoding="utf-8") as fh:
            json.dump(stats, fh, ensure_ascii=False, indent=1)
        print("\ncorpus saved -> %s  (재사용: --corpus %s)" % (a.save_corpus, a.save_corpus))

    print("\nOpen-access only: this is what the journal's OA papers look like, not a random")
    print("sample of it. Use it to find what is out of shape, not as a target to hit exactly.")
    return 1 if (a.fail_outside and outside) else 0


if __name__ == "__main__":
    sys.exit(main())
