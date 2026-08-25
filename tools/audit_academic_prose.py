#!/usr/bin/env python3
"""학술 산문 문체 검사 — 지도교수가 «실제로 지운» 네 부류를 기계로 잡는다.

왜 이 도구가 있는가 (teacher-ohs 2026-08-22):
  연구요약을 지도교수께 올렸더니 문장을 지우기만 하고 구두점을 안 치우신 채 돌아왔다.
  삭제 자리가 그대로 보여(`. 첫째,` / `…이다. , 그해` / `…)..`) 원본과 대조해 10건을
  복원했고, 네 부류로 갈렸다. 지적 원문:
    "문체가 구어적이고 학술적인 발표로 보기에 상당히 어색합니다. … 일반적인 원칙을
     제시하기 힘든데 심하게 느껴지는 것만 몇개 수정했습니다."
  「일반적인 원칙을 제시하기 힘들다」고 하셨지만, 지우신 문장들에는 규칙성이 있었다.
  그 규칙성만 코드로 옮긴다.

  ⭐ 산문 규칙을 feedback에 적어 두면 «읽어야» 작동하고 그 기억은 매 세션 리셋된다.
     [[core]] §5 — 규칙이 반복해 깨지면 코드로 옮길 수 있는지부터 묻는다.

무엇을 잡는가:
  ① 예고·메타 문장   뒤에 나올 내용을 미리 알려주는 문장 (구조는 내용이 스스로 보여준다)
  ② 선택의 «이유»     «방법» 절에서 왜 그렇게 했는지 설명하는 문장 -- 방법은 «한 것»만 적는다
  ③ 저널리즘 표현     사실이 아니라 «분위기»의 서술
  ④ 수치 중복         Key findings·초록에 이미 있는 추정치를 본문에서 되풀이

설계 원칙 (넓히지 않는다):
  · [[core]] §3 — SURFACE_RULES를 넓히면 「남길 것」까지 잡혀 «검사를 끄게 된다».
    그래서 ②는 «방법 절에서만» 본다(고찰의 「때문이다」는 기전 설명이라 정상이다).
  · 완전 자동판정이 아니다. [FAIL]은 근거가 확실한 것만, 나머지는 [의심]으로 사람에게 넘긴다.
  · 국문·영문 둘 다 본다 — 국문만 고치고 영문 초록을 두면 «절반만» 고친 것이다(실측).

USAGE
    python audit_academic_prose.py FILE.docx
    python audit_academic_prose.py FILE.md --strict     # [의심]도 실패로 친다
    python audit_academic_prose.py --selftest           # 규칙이 known-good을 안 잡는지 확인

EXIT
    0 통과 · 1 [FAIL] 있음
"""
import argparse
import os
import re
import sys

# Windows 기본 콘솔은 cp949 라 ⚠·«»·— 에서 UnicodeEncodeError 로 «검사기가 죽는다».
# 그러면 audit.py 가 [FAIL] 로 집계해서, 산문에 문제가 없는데 있는 것처럼 보인다.
# 2026-08-22 실측: face_mci_reversion 원고 감사에서 정확히 그렇게 오인했다 —
# 죽은 자리가 "절을 못 찾았다" 경고문(⚠)을 찍으려던 280 행이었다.
# 같은 폴더의 audit_display_items.py·audit_font_sizes.py 등은 이미 이 가드를 갖고 있었다.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- 절 인식
# 절을 알아야 ②(방법 절 한정)를 좁게 걸 수 있다.
# ⚠ **절을 못 찾으면 ②가 조용히 안 돈다** — 그건 통과가 아니라 «검사 안 함»이다.
#   실측(2026-08-22): 다른 파이프라인 브리프 두 개가 「통과」로 나왔는데 절이 ? 하나뿐이었다.
#   원인은 번호 붙은 표제(`3. 연구 방법`, `2. Methods`, `3.1 자료원`)를 정규식이 못 잡은 것.
#   -> 앞머리 번호를 허용하고, 하위 절(3.1 등)도 상위 절에 귀속시킨다.
#   -> 그리고 절을 하나도 못 찾으면 main()에서 «경고»한다(조용히 통과시키지 않는다).
_NUM = r"(?:\d+(?:\.\d+)*[.)]?\s*)?"          # 「3.」 「3.1」 「2)」 …
SECTION_PAT = [
    ("background", r"^\s*#*\s*" + _NUM + r"(배경|서론|연구\s*배경|Background|Introduction)\b"),
    ("methods",    r"^\s*#*\s*" + _NUM + r"(방법|연구\s*방법|Methods?|Materials and Methods)\b"),
    ("results",    r"^\s*#*\s*" + _NUM + r"(결과|연구\s*결과|Results?|Primary Result)\b"),
    ("discussion", r"^\s*#*\s*" + _NUM + r"(고찰|논의|Discussion)\b"),
    ("conclusion", r"^\s*#*\s*" + _NUM + r"(결론|Conclusions?)\b"),
    ("limits",     r"^\s*#*\s*" + _NUM + r"(제한점|한계|Limitations?|Strengths and)\b"),
    ("key",        r"^\s*#*\s*" + _NUM + r"(Key findings?|핵심\s*결과|주요\s*결과|Study Overview)\b"),
]
# 표제로 오인하면 안 되는 것 — 본문 문장이 「결과적으로…」처럼 시작할 수 있다.
# 표제는 짧다. 40자를 넘으면 본문으로 본다.
MAX_HEADING_LEN = 40

# ---------------------------------------------------------------- ③ 저널리즘 표현
# 사실이 아니라 «분위기». 좁게 — 숫자를 동반한 「증가하였다」 같은 것은 잡지 않는다.
JOURNALISM = [
    (r"전국적\s*(인)?\s*의제가?\s*되", "「전국적 의제가 되었다」 — 분위기 서술"),
    (r"사회적\s*(공분|파장|반향)", "감정어"),
    (r"화두(가|로)\s*(되|떠올)", "「화두가 되었다」"),
    (r"뜨거운\s*감자", "관용구"),
    (r"논란이\s*되(었|고)", "「논란이 되었다」 — 무엇이 논란인지 사실로 쓴다"),
    (r"주목(을)?\s*받(고|았)", "「주목받고 있다」"),
    (r"대두되(었|고)", "「대두되었다」"),
    (r"became a national (issue|agenda)", "journalistic"),
    (r"has become a hot (topic|issue)", "journalistic"),
    (r"(has|have) drawn (public )?attention", "journalistic"),
    (r"sparked (a )?(debate|controversy|outrage)", "journalistic"),
]

# ---------------------------------------------------------------- ① 예고·메타
# 어휘 규칙은 좁게, 나머지는 «구조»로 잡는다(아래 check_preview 참조).
PREVIEW_LEX = [
    (r"밝히고자\s*한\s*것은", "예고 — 뒤에 그 내용이 온다"),
    (r"살펴보고자\s*한다", "예고"),
    (r"다음과\s*같이\s*구성(된다|하였다)", "문서 구조 안내"),
    (r"서로\s*(다른\s*)?(약점|장단점)을\s*가지며", "예고 — 뒤에 각각의 설명이 온다"),
    (r"(상호|서로)\s*보완한다", "예고 — 뒤에 각각의 설명이 온다"),
    (r"두\s*자료원을\s*함께\s*(썼|사용하였)다", "예고 — 뒤에 첫째는·둘째는이 온다"),
    (r"We combined two (sources|datasets)", "preview"),
    (r"(questions?|issues?) remained? open", "preview"),
    (r"we (set out to|aimed? to|sought to)", "preview"),
    (r"This (paper|study) is organi[sz]ed as follows", "document roadmap"),
    (r"complement each other", "preview"),
]

# ---------------------------------------------------------------- ② 선택의 이유 (방법 절 한정)
REASON_IN_METHODS = [
    # 근거를 «지우라»가 아니다 — 「~때문이다」로 매달지 말고 사실로 진술하라는 뜻이다.
    (r"때문이다\s*[.。]", "「~때문이다」 — 근거는 남기되 사실 진술로 바꾼다"),
    # 「하도록」이 아니라 «어떤 동사든» + 도록 하였다 -- 실측: 「먼저 오도록 하였다」를
    # 「하도록」으로 걸었더니 자기시험이 놓쳤다. 「~하였다」로 끝내면 될 것을 의도로 늘린 형태다.
    (r"도록\s*하(였|았)다", "「~도록 하였다」 — 의도 서술. 「~하였다」로 끝낸다"),
    (r"위해서이다", "「~위해서이다」"),
    (r"이유는\s*.{0,30}때문", "이유 설명"),
    (r"This is because", "reason-giving in Methods"),
    (r"The reason (is|for this)", "reason-giving in Methods"),
    (r"in order to ensure that", "intent statement"),
]
# 논증의 «해설» — 절 무관. 결과가 왜 그 주장을 지지하는지 풀어 쓰는 문장.
ARGUMENT_GLOSS = [
    (r"[^.]{5,}(라면|뿐이라면)[^.]{0,60}(이유가\s*없다|달라질\s*이유)", "논증의 해설"),
    (r"which .{0,40}(does not|cannot) explain", "argument gloss"),
    (r"이는\s*.{0,30}(의미한다|뜻한다)\s*[.。]", "재진술 — 앞 문장이 이미 말했다"),
]

NUM_PAT = re.compile(r"\d+\.\d{2}\s*\(95%\s*(?:CI)?\s*[\d.]+[–\-~][\d.]+\)")


# ---------------------------------------------------------------- 읽기
def read_text(path):
    """(절이름, 문단) 목록. .docx / .md / .txt."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        from docx import Document
        paras = [p.text.strip() for p in Document(path).paragraphs]
    else:
        with open(path, encoding="utf-8") as fh:
            paras = [ln.strip() for ln in fh.read().split("\n")]
    out, sec = [], "?"
    for t in paras:
        if not t:
            continue
        hit = None
        if len(t) <= MAX_HEADING_LEN:
            hit = next((n for n, pat in SECTION_PAT if re.match(pat, t, re.I)), None)
        if hit:
            sec = hit
            continue
        out.append((sec, t))
    return out


def sentences(par):
    return [s.strip() for s in re.split(r"(?<=[.。다])\s+", par) if s.strip()]


# ---------------------------------------------------------------- 검사
def check_preview(paras):
    """① 예고·메타. 어휘 + «구조»(문단 첫 문장 뒤에 「첫째」가 오면 그 첫 문장은 예고다)."""
    hits = []
    for sec, par in paras:
        ss = sentences(par)
        for pat, why in PREVIEW_LEX:
            for s in ss:
                if re.search(pat, s, re.I):
                    hits.append(("FAIL", sec, s, "①예고·메타: " + why))
        # 구조 규칙: 첫 문장 다음 문장이 「첫째」로 시작하면 첫 문장은 예고일 가능성이 높다.
        # 첫 문장 자체가 「첫째」면 정상(감수자가 남긴 형태가 그것이다).
        if len(ss) >= 2 and re.match(r"^(첫째|첫\s*번째|First[,\s])", ss[1]) \
                and not re.match(r"^(첫째|첫\s*번째|First[,\s])", ss[0]) \
                and len(ss[0]) < 60:
            hits.append(("의심", sec, ss[0],
                         "①예고·메타(구조): 바로 뒤 문장이 「첫째」로 시작한다"))
    return hits


def check_reason(paras):
    """② 방법 절의 이유 서술 + 절 무관 논증 해설.

    🔴 **감수자 수정본을 known-good으로 쓰지 않는다** (사용자 정정 2026-08-22).
      메일에 「심하게 느껴지는 것만 몇개 수정했습니다」라고 쓰여 있다 — 최소 수정본이지
      최종본이 아니다. 그분이 «남기신» 문장이라고 해서 옳은 문장은 아니다.
      한때 「~때문이다에 수치가 있으면 봐준다」로 좁혔는데, 근거가 「감수자가 안 지웠다」뿐이라
      되돌렸다. 검사기의 기준은 «지우신 것»에서만 뽑는다.

    ⚠ 다만 지적 문구는 정확해야 한다: 이유를 «지우라»는 뜻이 아니다.
      「…넣지 않았다. 표본이 13% 줄기 «때문이다»」 → 「…넣지 않았다. 이들을 넣으면 결측으로
      표본이 13% 준다.」 근거는 남기고 «매다는 형식»만 사실 진술로 바꾼다.
    """
    hits = []
    for sec, par in paras:
        for s in sentences(par):
            if sec == "methods":
                for pat, why in REASON_IN_METHODS:
                    if re.search(pat, s, re.I):
                        hits.append(("FAIL", sec, s, "②선택의 이유: " + why))
            for pat, why in ARGUMENT_GLOSS:
                if re.search(pat, s, re.I):
                    hits.append(("의심", sec, s, "②논증의 해설: " + why))
    return hits


def check_journalism(paras):
    hits = []
    for sec, par in paras:
        for s in sentences(par):
            for pat, why in JOURNALISM:
                if re.search(pat, s, re.I):
                    hits.append(("FAIL", sec, s, "③저널리즘 표현: " + why))
    return hits


def check_duplicate_numbers(paras):
    """④ Key findings·초록에 실린 추정치가 본문에 다시 나오는지."""
    key = {n for sec, par in paras if sec in ("key", "conclusion")
           for n in NUM_PAT.findall(par)}
    hits = []
    seen = set()
    for sec, par in paras:
        if sec in ("key", "conclusion"):
            continue
        for n in NUM_PAT.findall(par):
            if n in key and n not in seen:
                seen.add(n)
                hits.append(("의심", sec, n,
                             "④중복: Key findings/결론에 이미 실린 추정치"))
    return hits


# ---------------------------------------------------------------- 자기시험
SELFTEST_GOOD = [
    ("methods", "성별·연령·교육수준을 보정한 설계가중 로지스틱 회귀를 각 wave 안에서 수행하였다."),
    ("methods", "노출은 직전 wave의 값을 사용하였다."),
    ("background", "학생과 학부모에 의한 폭력은 교사에게 인정된 직업적 유해요인이다."),
    ("background", "첫째, 교사가 다른 직군보다 실제로 더 노출되는가. 둘째, 그 양상이 변했는가."),
    ("discussion", "노출이 많은 교원이 소진에 취약한 특성을 함께 가지고 있기 때문이다."),
    ("limits", "추정치의 신뢰구간이 넓다."),
]
SELFTEST_BAD = [
    ("background", "한국에서 이 문제는 2023년 전국적 의제가 되었고, 법이 개정되었다."),
    ("methods", "두 자료원을 함께 썼다. 첫째는 근로환경조사다."),
    ("methods", "2023년 오즈비만으로는 비교군도 나빠졌는지 알 수 없기 때문이다."),
    ("methods", "노출이 결과보다 먼저 오도록 하였다."),
    ("discussion", "두 자료원은 서로 다른 약점을 가지며 상호 보완한다."),
]


def selftest():
    """known-good을 잡으면 규칙이 틀린 것이다. 넓히기 전에 항상 이걸 돌린다."""
    bad = 0
    for sec, s in SELFTEST_GOOD:
        h = [x for f in (check_preview, check_reason, check_journalism)
             for x in f([(sec, s)]) if x[0] == "FAIL"]
        if h:
            bad += 1
            print("  ✗ known-good을 FAIL로 잡았다: %s\n      -> %s" % (s[:60], h[0][3]))
    for sec, s in SELFTEST_BAD:
        h = [x for f in (check_preview, check_reason, check_journalism)
             for x in f([(sec, s)]) if x[0] == "FAIL"]
        if not h:
            bad += 1
            print("  ✗ 잡아야 할 문장을 놓쳤다: %s" % s[:60])
    print("자기시험: known-good %d · known-bad %d · 실패 %d"
          % (len(SELFTEST_GOOD), len(SELFTEST_BAD), bad))
    return 1 if bad else 0


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", nargs="?")
    ap.add_argument("--strict", action="store_true", help="[의심]도 실패로 친다")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()

    if a.selftest:
        sys.exit(selftest())
    if not a.path:
        ap.error("path 또는 --selftest 가 필요하다")

    paras = read_text(a.path)
    hits = (check_preview(paras) + check_reason(paras)
            + check_journalism(paras) + check_duplicate_numbers(paras))

    secs = {s for s, _ in paras}
    print("학술 산문 문체 검사: %s" % os.path.basename(a.path))
    print("  문단 %d · 절 %s" % (len(paras), ", ".join(sorted(secs))))

    # 🔴 절을 못 찾으면 ②(방법 절 한정)가 «조용히» 안 돈다. 통과로 착각하지 않게 알린다.
    if "methods" not in secs:
        print("  ⚠ 방법 절을 찾지 못했다 — ②「선택의 이유」 검사가 «돌지 않았다».")
        print("    표제가 다른 형태일 수 있다. 통과로 읽지 말 것.")

    if not hits:
        print("  통과 — 네 부류 모두 걸리지 않았다"
              + ("" if "methods" in secs else " (단, ②는 안 돌았다)"))
        return 0

    nf = sum(1 for h in hits if h[0] == "FAIL")
    for lvl, sec, s, why in sorted(hits, key=lambda x: (x[0] != "FAIL", x[1])):
        print("  [%s] (%s) %s\n        %s" % (lvl, sec, why, s[:110]))
    print("\n  FAIL %d · 의심 %d" % (nf, len(hits) - nf))
    print("  ⚠ [의심]은 후보다 — 원문을 열어 사람이 판정한다. 근거 = manuscript_rules.md")
    return 1 if (nf or (a.strict and hits)) else 0


if __name__ == "__main__":
    sys.exit(main())
