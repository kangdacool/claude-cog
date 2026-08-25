#!/usr/bin/env python3
"""journal_display_census.py — 목표 지면이 «실제로» 싣는 표·그림을 센다.

[[manuscript-rules]]: 「서식은 추측하지 말고 그 저널의 게재 논문 실물로 확인한다」.
`journal_fit.py`가 «산문의 형태»(절별 단어 수·문체)를 재는 짝이라면, 이쪽은
**«전시물의 형태»** — 표 몇 개, 그림 몇 개, 그 캡션이 무엇인가 — 를 잰다.
원고를 쓰기 «전에» 돌린다. journal_fit은 원고를 요구하지만 이건 요구하지 않는다.

    python journal_display_census.py        # 아래 TERM을 고쳐 쓴다

⚠️ **설계를 갈라 봐야 한다.** 리뷰는 PRISMA 흐름도·비뚤림 표·GRADE 표로, 시험은
   CONSORT로 부풀려져 코호트 규범을 흐린다. 섞어 세면 「표 6개가 보통」 같은 틀린
   기준이 나온다 -- 2026-08-24 실측: 전부 섞으면 표 중앙 6·그림 4, **코호트만 보면
   표 3·그림 2**였다.

⚠️ **개수를 `<back>` 앞으로 잘라 세지 마라.** 출판사마다 표를 `<floats-group>`이나
   `<back>` 뒤에 둔다 -- 그렇게 셌다가 캡션이 4개인 논문이 「0표」로 찍혔다(실측).
   캡션 목록에서 세고 «보충자료» 표시만 뺀다.

⚠️ **저널 하나로 좁히면 PMC 공개분이 없을 수 있다**(구독지. Int Psychogeriatr는 1편).
   그럴 땐 저널이 아니라 «같은 설계»로 검색해 분야 규범을 잰다 -- 어차피 물어야 할 것은
   「이런 연구는 표·그림을 몇 개로 내는가」이기 때문이다.
"""
import json
import re
import sys
import time
import urllib.parse
import urllib.request

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
UA = {"User-Agent": "Mozilla/5.0 (research use)"}
E = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


def get(url):
    for _ in range(3):
        try:
            return urllib.request.urlopen(
                urllib.request.Request(url, headers=UA), timeout=60).read()
        except Exception:
            time.sleep(2)
    return b""


def esearch(term, db="pmc", n=40):
    u = f"{E}/esearch.fcgi?db={db}&retmode=json&retmax={n}&term={urllib.parse.quote(term)}"
    d = json.loads(get(u) or b"{}")
    r = d.get("esearchresult", {})
    return r.get("count", "0"), r.get("idlist", [])


def census(pmcid):
    xml = get(f"{E}/efetch.fcgi?db=pmc&id={pmcid}&retmode=xml").decode("utf-8", "replace")
    if "<article" not in xml:
        return None
    title = re.search(r"<article-title>(.*?)</article-title>", xml, re.S)
    title = re.sub(r"<[^>]+>", "", title.group(1)).strip() if title else "?"
    jr = re.search(r"<journal-title>(.*?)</journal-title>", xml, re.S)
    jr = re.sub(r"<[^>]+>", "", jr.group(1)).strip() if jr else "?"

    def caps(tag):
        out = []
        for m in re.finditer(rf"<{tag}\b.*?</{tag}>", xml, re.S):
            block = m.group(0)
            lab = re.search(r"<label>(.*?)</label>", block, re.S)
            cap = re.search(r"<caption>(.*?)</caption>", block, re.S)
            lab = re.sub(r"<[^>]+>", "", lab.group(1)).strip() if lab else ""
            cap = re.sub(r"<[^>]+>", " ", cap.group(1)) if cap else ""
            cap = re.sub(r"\s+", " ", cap).strip()
            out.append(f"{lab}: {cap[:110]}")
        return out

    ## ⚠️ <back> 앞으로 자르는 방식은 틀렸다 -- 출판사마다 표를 <floats-group>이나
    ##    <back> 뒤에 두어, 캡션이 4개인 논문이 「0표」로 찍혔다(실측).
    ##    캡션 목록에서 세되 «보충자료»로 표시된 것만 뺀다.
    def is_supp(s):
        return bool(re.search(r"supplement|appendix|eTable|eFigure|보충", s, re.I))
    tabs = [t for t in caps("table-wrap") if not is_supp(t)]
    figs = [f for f in caps("fig") if not is_supp(f)]
    ## 설계 판정: 리뷰·메타분석·RCT는 표 구성이 달라 코호트 규범을 흐린다.
    kind = "cohort"
    if re.search(r"meta-analys|systematic review|PRISMA", xml[:40000], re.I):
        kind = "review"
    elif re.search(r"randomi[sz]ed|randomised controlled|CONSORT", xml[:40000], re.I):
        kind = "trial"
    return {"journal": jr, "title": title, "kind": kind,
            "n_tab": len(tabs), "n_fig": len(figs), "tables": tabs, "figs": figs}


## ⚠️ 저널 하나로 좁히면 PMC 공개분이 1편뿐이다(Int Psychogeriatr는 구독지).
##    「저널의 평균 수준」을 재려면 «같은 설계를 싣는 지면» 전체를 봐야 한다.
TERM = ('("mild cognitive impairment"[tiab] OR MCI[tiab]) AND depress*[tiab] '
        'AND ("progression to dementia"[tiab] OR "conversion to dementia"[tiab] '
        'OR "incident dementia"[tiab]) AND (cohort[tiab] OR longitudinal[tiab] '
        'OR prospective[tiab]) AND 2019:2026[dp]')
cnt, ids = esearch(TERM, n=40)
print(f"MCI·우울·치매전환 코호트 (PMC 전문) : {cnt}건 중 {len(ids)}건 조회\n")

rows = []
for i in ids[:20]:
    r = census(i)
    if not r or r["n_tab"] + r["n_fig"] == 0:
        continue
    rows.append(r)
    print(f"[{r['kind']:6s} {r['n_tab']}표 {r['n_fig']}그림] {r['title'][:80]}")
    if r["kind"] == "cohort":
        for t in r["tables"]:
            print(f"        T {t}")
        for f in r["figs"]:
            print(f"        F {f}")
    print()

if rows:
    import statistics as st
    print("=" * 74)
    for kind in ("cohort", "review", "trial"):
        g = [r for r in rows if r["kind"] == kind]
        if not g:
            continue
        tt = [r["n_tab"] for r in g]
        ff = [r["n_fig"] for r in g]
        print(f"{kind:7s} {len(g):2d}편 | 표 중앙 {st.median(tt):.0f} ({min(tt)}~{max(tt)})"
              f" | 그림 중앙 {st.median(ff):.0f} ({min(ff)}~{max(ff)})"
              f" | 합 중앙 {st.median([a + b for a, b in zip(tt, ff)]):.0f}")
    print("\n⚠️ 코호트 원저만 우리 규범이다 -- 리뷰는 PRISMA·비뚤림표로, 시험은 CONSORT로 부풀려진다.")
