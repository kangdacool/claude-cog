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
    return "\n".join(out)


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
        """Start offset of a heading, at either level, or None."""
        for prefix in ("## ", "# "):
            m = re.search(r"^%s%s\s*$" % (re.escape(prefix), re.escape(label)),
                          t, re.MULTILINE)
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

    Why (2026-08-24, teacher_ohs): a manuscript source carried Korean working notes inside
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


def draft_body(path, start, end):
    t = as_markdown(path)
    if start and start in t:
        t = t[t.index(start):]
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


def report_style(rows, draft_path, start, end):
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
    ap.add_argument("--journal", required=True, help='e.g. "Am J Epidemiol"')
    ap.add_argument("--terms", default="", help='e.g. "methodology OR simulation"')
    ap.add_argument("--n", type=int, default=25)
    # Positional so audit.py can call it the way it calls every other check
    # (tool.py FILE ...); --draft kept for the existing invocations.
    ap.add_argument("draft", nargs="?", help="draft .md or the built .docx")
    ap.add_argument("--draft", dest="draft_opt", help="same path, as a flag")
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--only", choices=["shape", "style"], help="default: both")
    a = ap.parse_args()
    a.draft = a.draft or a.draft_opt
    if not a.draft:
        ap.error("give a draft: a positional path, or --draft PATH")

    ids = pmids(a.journal, a.terms, a.n)
    rows = [r for r in (measure(p) for p in ids) if r]
    if not rows:
        print("no open-access full text matched — widen --terms or raise --n")
        return 1

    print("%s — %d open-access articles measured (of %d searched)\n"
          % (a.journal, len(rows), len(ids)))
    if a.only != "style":
        report_shape(rows, a.draft)
    if a.only != "shape":
        report_style(rows, a.draft, a.start, a.end)
    print("\nOpen-access only: this is what the journal's OA papers look like, not a random")
    print("sample of it. Use it to find what is out of shape, not as a target to hit exactly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
