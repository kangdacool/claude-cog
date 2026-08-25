#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_display_items.py — 표·그림의 «번호 체계»가 성립하는지. 본문을 읽어서 판정한다.

    python audit_display_items.py MS.docx [MS.md ...]

잡는 것 (전부 기계적으로 확실한 것만)
------------------------------------
  순서       번호가 «본문 첫 인용 순서»와 어긋남   ← 저널 관례이자 조판 단계에서 잡히는 것
  고아       문서에 실렸는데 본문에서 한 번도 인용되지 않은 표·그림
  유령       본문이 인용하는데 그 번호의 표·그림이 문서에 없음
  결번       Table S1, S3 은 있는데 S2 가 없음
  [의심] 각주 길이가 형제들보다 튐 (문서 자신의 각주 분포 대비)

왜 있나 (2026-08-20 한 원고에서)
-----------------------------
원고에 보충 표를 하나 더하면서 «순서가 맞나»를 손으로 확인했더니, 새로 넣은 표가 아니라
**원래 있던 Table S1 과 S2 가 뒤바뀌어 있었다** — 외적타당도 표는 Results 앞쪽(L78)에서,
원인별사망 표는 뒤쪽(L82)에서 처음 인용되는데 번호는 반대였다. 사람이 물어봐 줘야만
드러나는 종류이고, 기존 검사 어느 것도 이 축을 보지 않았다
(`audit_text_consistency` 는 «참고문헌» 번호를 보지 표·그림 번호를 보지 않는다).

같은 날 각주 하나가 498자로 부풀어 있었다. 그 문서의 다른 각주는 158-390자였다.
**길이 자체가 아니라 «형제 대비 이탈»이 신호**다 — 명세만 적으면 그 문서의 다른 각주와
비슷한 길이가 되고, 변론이 섞이면 튄다. 그래서 절대 기준을 두지 않고 분포로 본다.

⚠️ 각주 길이는 [의심]이다. 표가 복잡하면 각주도 길어지는 게 정상이라 실패로 치지 않는다.
종료코드 1 은 «확실한 결함»(순서·고아·유령·결번)에만.

읽는 형식: .docx · .md/.txt · .hwpx  (`audit_text_consistency.read_any` 재사용)
"""
import argparse
import collections
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from audit_text_consistency import read_any            # noqa: E402

# "Table 1" / "Table S4" / "Figure S2" / "Figures 1 and 2" 는 첫 토큰만 잡는다.
REF = re.compile(r"\b(Table|Figure|Fig\.)\s+(S?)(\d{1,2})\b")
# 표시물 «자체»의 제목 줄: 문단이 그 이름으로 시작하고 마침표나 줄바꿈이 따라온다.
TITLE = re.compile(r"^\s*(Table|Figure)\s+(S?)(\d{1,2})\s*[.:]")

KIND = {"Fig.": "Figure", "Figure": "Figure", "Table": "Table"}


def key(kind, s, num):
    return "%s %s%d" % (KIND[kind], s, int(num))


def scan(body):
    """본문 문단들에서 «표시물 제목»과 «본문 인용»을 분리해 모은다.

    제목 줄 자체는 인용이 아니다 - 이걸 섞으면 어떤 표든 «자기 자신에게 인용된»
    것이 되어 고아 검사가 통째로 무력해진다.
    """
    titles, cites = {}, []
    for i, para in enumerate(body):
        t = TITLE.match(para)
        if t:
            k = key(t.group(1), t.group(2), t.group(3))
            titles.setdefault(k, i)
            continue                      # 제목 줄에서는 인용을 세지 않는다
        for m in REF.finditer(para):
            cites.append((key(m.group(1), m.group(2), m.group(3)), i))
    return titles, cites


def series(names):
    """'Table S' 같은 계열별로 묶는다 - 계열이 다르면 번호가 겹쳐도 정상이다."""
    out = collections.defaultdict(list)
    for n in names:
        m = re.match(r"(Table|Figure)\s+(S?)(\d+)", n)
        out["%s %s" % (m.group(1), m.group(2))].append(int(m.group(3)))
    return out


def footnote_lengths(path):
    """docx 전용: 본문보다 작은 이탤릭 문단 = 표 각주·그림 범례. 크기별로 묶어
    돌려준다(8pt 각주와 10pt 범례를 같은 자로 재면 안 된다)."""
    if not path.lower().endswith(".docx"):
        return {}
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        return {}
    doc = Document(path)
    base = doc.styles["Normal"].font.size or Pt(12)
    groups = collections.defaultdict(list)
    for p in doc.paragraphs:
        runs = [r for r in p.runs if r.text.strip()]
        if not runs or not all(r.italic for r in runs):
            continue
        sizes = {r.font.size for r in runs if r.font.size}
        if len(sizes) != 1:
            continue
        sz = sizes.pop()
        if sz >= base:
            continue
        text = "".join(r.text for r in runs).strip()
        if len(text) > 40:                       # 한두 단어짜리 이탤릭은 각주가 아니다
            groups[sz.pt].append((len(text), text))
    return groups


def locked(path):
    """Word가 열고 있으면 읽기가 PermissionError로 죽는다 — 그런데 그 상태가 «검토 중»,
    즉 이 검사를 가장 돌리고 싶은 순간이다. 더 나쁜 것은 `zipfile.is_zipfile()`이 그
    예외를 삼키고 False를 돌려준다는 것: 잠긴 파일이 «손상된 파일»로 보고된다.
    잠김을 잠김이라고 말한다."""
    try:
        with open(path, "rb"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def audit(path):
    body, cells = read_any(path)
    titles, cites = scan(body)
    hard, soft = [], []

    present = set(titles)
    cited = collections.OrderedDict()
    for k, i in cites:
        cited.setdefault(k, i)

    # -- 유령·고아 --
    ## ⭐⭐ 보충표(Table S1, Figure S2 …)는 «다른 파일에 사는 것이 정의»다.
    ##    본문 원고에서 보면 문서에 없으니 «유령», 보충 파일에서 보면 그 안에서 인용되지
    ##    않으니 «고아»로 잡힌다 -- 양쪽 다 오탐이고, 그걸 없애려고 문서를 비틀면 진짜
    ##    결함을 못 보게 된다(2026-08-25 psy 원고에서 12건 나왔다).
    ##    본문↔보충의 «대응»은 두 파일을 함께 봐야 하므로 이 도구의 일이 아니다.
    ##    보충을 «만드는» 빌더가 그 검사를 한다(psy: build_supplement.py 자기검사).
    def _is_suppl(k):
        return re.search(r"(?:Table|Figure|표|그림)\s*S\s*\d", k, re.I) is not None

    for k in cited:
        if k not in present:
            if _is_suppl(k):
                soft.append(("보충 인용", "%s -- 보충자료에 있어야 한다(이 파일 밖). "
                                      "대응 검사는 보충 빌더의 몫" % k))
            else:
                hard.append(("유령", "%s 를 본문이 인용하는데 그 표시물이 문서에 없다" % k))
    for k in present:
        if k not in cited:
            if _is_suppl(k):
                soft.append(("보충 표시물", "%s -- 본문 원고 쪽에서 인용된다(이 파일 밖)" % k))
            else:
                hard.append(("고아", "%s 가 실려 있는데 본문에서 한 번도 인용되지 않는다" % k))

    # -- 결번 --
    for name, nums in series(present).items():
        nums = sorted(nums)
        missing = [n for n in range(1, max(nums) + 1) if n not in nums]
        if missing:
            hard.append(("결번", "%s: %s 까지 있는데 %s 가 없다"
                         % (name.strip(), max(nums),
                            ", ".join("%s%d" % (name.split()[1] if len(name.split()) > 1 else "", m)
                                      for m in missing))))

    # -- 순서: 계열 안에서 «첫 인용 위치»가 번호 순이어야 한다 --
    for name, nums in series(set(cited) & present).items():
        seq = sorted(((cited[k], k) for k in cited if k in present
                      and k.startswith(name.rstrip())), key=lambda x: x[0])
        seq = [k for _, k in seq if re.match(r"%s\s*\d" % re.escape(name.rstrip()), k)]
        want = sorted(seq, key=lambda k: int(re.search(r"(\d+)$", k).group(1)))
        if seq != want:
            hard.append(("순서", "%s 의 첫 인용 순서가 %s 인데 번호는 %s 순이다"
                         % (name.strip(), " -> ".join(seq), " -> ".join(want))))

    # -- [의심] 각주 길이 이탈 --
    for pt, items in sorted(footnote_lengths(path).items()):
        if len(items) < 3:
            continue
        lens = sorted(l for l, _ in items)
        med = lens[len(lens) // 2]
        for l, text in items:
            if l > med * 1.6:
                soft.append(("각주 길이", "%.0fpt 각주가 %d자 — 같은 문서의 형제 중앙값 %d자의 "
                                          "%.1f배: \"%s...\"" % (pt, l, med, l / med, text[:60])))

    # inventory: every item that IS in the document, with where the body first
    # cites it (or a loud marker when nothing does)
    inventory = []
    for k in sorted(present, key=lambda s: (s.split()[0], "S" in s, int(re.search(r"\d+$", s).group()))):
        i = cited.get(k)
        inventory.append((k, ("«NEVER CITED»" if i is None
                              else '"%s..."' % body[i].strip()[:52])))
    return hard, soft, inventory


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    a = ap.parse_args()

    total = skipped = 0
    for f in a.files:
        print("\n===== %s =====" % os.path.basename(f))
        if locked(f):
            print("  SKIP  다른 프로그램(대개 Word)이 이 파일을 열고 있어 읽을 수 없다.\n"
                  "        닫고 다시 돌리거나 사본으로 검사할 것 — 검사를 «통과»한 것이 아니다.")
            skipped += 1
            continue
        hard, soft, inventory = audit(f)
        # Print what WAS checked, not just what failed. "No warnings" and "no
        # look" are the same output otherwise, and the question this tool exists
        # to answer - "is anything uncited?" - deserves a positive answer.
        print("  %d display item(s), each cited in the body:" % len(inventory))
        for k, where in inventory:
            print("    %-11s first cited in: %s" % (k, where))
        if not hard and not soft:
            print("  표시물 번호 체계 이상 없음")
        for cat, msg in hard:
            print("  [결함] %-6s %s" % (cat, msg))
        for cat, msg in soft:
            print("  [의심] %-6s %s" % (cat, msg))
        total += len(hard)
    if total:
        print("\n확실한 결함 %d건" % total)
    if skipped:
        print("잠겨서 못 읽은 파일 %d개 — «검사 안 함»이지 «통과»가 아니다" % skipped)
    sys.exit(1 if total or skipped else 0)


if __name__ == "__main__":
    main()
