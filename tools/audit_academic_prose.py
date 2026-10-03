#!/usr/bin/env python3
"""학술 산문 문체 검사 — 지도교수가 «실제로 지운» 네 부류를 기계로 잡는다.

왜 이 도구가 있는가 (한 원고 2026-08-22):
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
  ⑤ 부재 서술         «방법» 절에서 「안 한 것」을 선언하는 문장 -- 설계가 아니라 해명이다
                      (①과 달리 어휘가 아니라 «구조»로 잡는다. 이유는 ABSENCE_PAT 주석)
  ⑥ 검증·추출 일지    «방법» 절에 파이프라인이 «겪은 일»(표 제목·열 위치·대조 결과 N/N·원자료
                      오류)을 적은 문장 -- 전부 참이라 다른 검사가 못 잡는다. 이유는 LOG_PAT 주석
  ⑦ 고칠 수 있는 제한점 «고찰»의 제한점이 자료의 «시점·연도 불일치»를 고백한다 -- 제한점에 적기 전에
                      「맞출 수 있는 자료가 있는가」를 먼저 확인했는지 묻는다. 이유는 FIXABLE_PAT 주석

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
# 2026-08-22 실측: 한 코호트 결과 덱 원고 감사에서 정확히 그렇게 오인했다 —
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
## 2026-09-16 (KHP T4 세션 실측): 학술대회 양식은 「Ⅱ. 연구방법」처럼 로마숫자 접두어를 쓴다 — 못 읽으면
##   ②·⑤·⑥이 «조용히» 안 돈다. 로마숫자(유니코드·라틴 IVX)도 앞머리로 허용한다.
_NUM = r"(?:(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+|[IVX]+(?=[.)\s])|\d+(?:\.\d+)*)[.)]?\s*)?"   # 「3.」 「3.1」 「2)」 「Ⅱ.」 …
SECTION_PAT = [
    ("background", r"^\s*#*\s*" + _NUM + r"(배경|서론|연구\s*배경|Background|Introduction)\b"),
    ("methods",    r"^\s*#*\s*" + _NUM + r"(방법|연구\s*방법|Methods?|Materials? and [Mm]ethods)\b"),
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
    # ⚠ 앞의 `(?<![가-힣])`가 없으면 «낱말 안»에서 걸린다 — 한국어는 띄어쓰기가 경계가 아니다.
    #   2026-09-11 실측 오탐: 「…역학조사 보고**서로 보완한다**」(보고서 + 조사 '로')가
    #   「서로 보완한다」로 잡혀, 결함 0인 연구계획서가 [FAIL]로 찍혔다.
    (r"(?<![가-힣])(상호|서로)\s*보완한다", "예고 — 뒤에 각각의 설명이 온다"),
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

# ---------------------------------------------------------------- ⑤ 방법 절의 «부재» 서술
# 「안 한 것」은 방법이 아니다. 방법 절에서 부재를 «선언»하는 문장은 설계의 서술이 아니라
# 내 결정에 대한 해명이고, 심사자는 그것을 「누가 물었길래?」로 읽는다 — 없던 의심을 만든다.
#
# 왜 어휘 목록이 아니라 «구조»인가 — 2026-09-05 한 원고 실측.
#   PREVIEW_LEX 에 `We combined two sources` 가 이미 있었고, 형제 빌더 두 곳에도 같은
#   문자열이 금지어로 박혀 있었다. 그런데 OEM 초록에 실린 것은 그 «부정형»이었다 —
#   「The two sources were never combined.」 게이트 셋이 전부 통과시켰고, 그 문장은 세 판
#   (260824 → 260903 → 260904)을 살아남았다. 중간에 «내가 문법까지 다듬었다».
#   ⭐ 문자열을 막으면 부정형·수동형·완곡형이 그대로 통과한다. 그래서 문법으로 잡는다:
#     ① 방법 절이고 ② 주어가 연구 «자산»이고 ③ 그 자산의 술어가 부정이고 ④ 수치가 없다.
#   ④가 결과의 부정(「did not change」·「were indistinguishable」)을 걸러낸다 — 그건
#   «발견»이고 대개 추정치를 동반한다. ③에서 부정을 «자산 바로 뒤»로 좁히는 것이
#   「X do not contribute to the analyses」류 오탐을 없앤다(거기선 자산이 목적어다).
ASSET_EN = r"(?:sources?|datasets?|data|analys[ei]s|models?|samples?|cohorts?|surveys?|panels?)"
ASSET_KO = r"(?:자료원|자료|분석|모형|표본|코호트|조사)"
ABSENCE_PAT = [
    (r"\b" + ASSET_EN + r"\b(?:\s+\w+){0,3}\s+(?:were|was|are|is|have|has|had)"
     r"(?:\s+\w+){0,2}\s+(?:never|not|neither)\b",
     "자산의 술어가 부정이다 — 「안 한 것」은 방법이 아니다"),
    (r"\bno (?:attempt|effort)s?\s+(?:was|were)\s+made\b", "부재의 선언"),
    (ASSET_KO + r"[은는이가][^.。]{0,25}(?:않았|않는|아니었|없었)", "자산의 술어가 부정이다"),
]

NUM_PAT = re.compile(r"\d+\.\d{2}\s*\(95%\s*(?:CI)?\s*[\d.]+[–\-~][\d.]+\)")

# ---------------------------------------------------------------- ⑥ 방법 절의 «검증·추출 일지»
# 2026-09-16 한 원고 실측. 자료원 절에 이런 문장이 네 개 있었다 —
#   「연보에는 제목이 같은 표가 … 함께 실려 있어, 환자 거주지 기준의 표만 사용하였다」
#   「연도판마다 표의 열 위치가 달라 열 제목을 기준으로 값을 추출하였다」
#   「… KOSIS 의 해당 통계와 비교한 결과, 304개 시도-연도 가운데 303개가 일치하였고 …」
#   「일치하지 않은 1건은 … 원자료의 오류였다」
# 전부 «참»이라 숫자 게이트·가드·감사 3회가 통과시켰고, 연구자가 잡았다(「진짜 좋은 학술 논문들에서
# 저런걸 서술하나? ARE YOU SURE?」). 목표 저널 원저 3편에는 같은 서술이 0회. 이것은 [[core]] §2⑧
# 「내 «과정»을 산출물 표면에 적는다」의 방법 절 판본이다 — 파이프라인이 겪은 일은 저장소 README 에.
# 독자에게 필요한 것은 «정의»와 «범위의 이유»뿐이고, 검증은 「공표값과 대조하였다」 한 절로 족하다.
# ⚠ 좁게 건다 — 「추출하였다」「대조하였다」 자체는 방법의 정당한 동사다. 걸리는 것은 «표의 생김새»
#   (열 위치·제목이 같은 표), «대조의 결과 보고»(N개 중 N개 일치·오차 중앙값), «원자료 오류의 서술»이다.
LOG_PAT = [
    (r"열\s*(위치|순서)", "표의 생김새(열 위치) — 저장소 README 의 일이다"),
    (r"제목이\s*같은\s*표", "표의 생김새(같은 제목의 표)"),
    (r"(비교|대조)한\s*결과.{0,40}(일치|불일치|오차)", "대조 «결과» 보고 — 「공표값과 대조하였다」 한 절이면 된다"),
    (r"\d+\s*개\s*(가운데|중)\s*\d+\s*개가?\s*일치", "N개 중 N개 일치 — 검증 일지"),
    (r"오차의\s*(중앙값|평균)", "검증 통계 — 결과가 아니라 일지"),
    (r"(원자료|원본|자료원)의\s*오류", "원자료 오류의 서술"),
    (r"(잘못|작게|크게)\s*실린", "원자료 오류의 서술"),
    (r"(파싱|스크립트|코드로|정규식|시트에서)", "구현 어휘 — 방법이 아니라 구현"),
    (r"column (positions?|order|headers?)", "table layout detail (repository README)"),
    (r"matched (the )?(published|KOSIS|official)[^.]{0,60}\b\d+\s*of\s*\d+", "validation tally in Methods"),
    (r"(error|mistake) in the (source|raw|original) (data|table|file)", "source-error narration"),
]

SELFTEST_LOG_BAD = [
    ("methods", "또한 연도판마다 표의 열 위치가 달라 열 제목을 기준으로 값을 추출하였다."),
    ("methods", "연보에는 제목이 같은 표가 환자 거주지 기준과 요양기관 소재지 기준으로 함께 실려 있어, 환자 거주지 기준의 표만 사용하였다."),
    ("methods", "국가통계포털의 해당 통계와 비교한 결과, 304개 시도-연도 가운데 303개가 일치하였고 오차의 중앙값은 0.00%였다."),
    ("methods", "일치하지 않은 1건은 2009년판에서 경기도 한 시군구의 합계 값이 급여유형별 값의 합보다 작게 실린 원자료의 오류였다."),
    ("methods", "Because column positions differ across editions, values were extracted by column header."),
]
SELFTEST_LOG_GOOD = [
    ("methods", "환자 거주지 기준의 관내·관외 진료비를 사용하였고, 추출한 값을 시도 단위로 합산하여 국가통계포털(KOSIS)의 공표값과 대조하였다."),
    ("methods", "연보는 2006년판부터 발간되었으나 2006·2007년판의 진료비는 약국 진료비를 제외한 기준이어서, 진료비 정의가 같은 2008년판부터 2024년판까지 17개 연도판을 사용하였다."),
    ("methods", "본 연구는 대표성이 확보된 국민건강보험공단 맞춤형 청구자료를 활용하였다."),
    ("methods", "Values were aggregated to the provincial level and compared with the published totals."),
    ("results", "304개 시도-연도 가운데 303개가 일치하였다."),   # 결과 절이면 ⑥의 대상이 아니다
    ("discussion", "원자료의 오류 가능성은 제한점이다."),
]


# ---------------------------------------------------------------- 읽기
def read_text(path):
    """(절이름, 문단) 목록. .hwpx / .docx / .md / .txt.

    ⚠ **.hwpx 를 빠뜨리면 「검사기가 죽은 것」이 「문서가 실패한 것」으로 보고된다.**
    2026-09-11 실측: audit.py 는 이 검사를 `report` 장르에 auto 로 물려 두는데, report 의
    대표 형식이 바로 .hwpx 다. zip+XML 을 UTF-8 텍스트로 열다 UnicodeDecodeError 로
    크래시했고, audit.py 는 종료코드만 보므로 `[FAIL] academic_prose` 로 찍혔다. 그 상태로
    deliverable_guard 가 산출물 전달을 막기까지 했다 — 문서에는 아무 문제가 없었다.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        from docx import Document
        paras = [p.text.strip() for p in Document(path).paragraphs]
    elif ext == ".hwpx":
        # 이미 있는 리더를 쓴다 — 세 번째 hwpx 리더를 짜면 «문단 경계» 규칙이 셋으로 갈린다.
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from audit_text_consistency import read_hwpx
        paras = read_hwpx(path)[0]   # 본문만. 표 셀은 산문이 아니라 자료라 이 규칙들의 대상이 아니다
    else:
        # ⚠ 줄바꿈으로 «접힌» 마크다운을 한 줄씩 보면 문장이 두 조각으로 갈린다. 그러면
        #   문장 단위 규칙이 조용히 못 잡는다 — 실측(2026-09-05): 원고 .md 의
        #   「…so the two sources are」/「not the same people and are never pooled.」 가
        #   두 줄로 갈려 ⑤가 안 걸렸다(같은 문장이 .docx 에서는 걸렸다).
        #   마크다운에서 문단 경계는 «빈 줄»이므로 빈 줄까지 이어 붙인다.
        try:
            with open(path, encoding="utf-8") as fh:
                raw = fh.read()
        except UnicodeDecodeError:
            # 텍스트가 아닌 형식이 여기까지 왔다. 트레이스백 대신 «무엇이 문제인지»를 말한다.
            raise SystemExit(
                "[academic_prose] 텍스트로 읽을 수 없는 형식: %s\n"
                "  지원: .hwpx / .docx / .md / .txt — 리더를 추가하거나 이 검사를 빼십시오."
                % (ext or "(확장자 없음)"))
        paras, buf = [], []
        for ln in raw.split("\n"):
            ln = ln.strip()
            if not ln:
                if buf:
                    paras.append(" ".join(buf))
                    buf = []
                paras.append("")
                continue
            # 표제·표·목록은 접지 않는다 (접으면 표제 인식이 깨진다).
            # ⚠ 「짧으면 표제」로 가르면 «문단의 마지막 줄»이 대개 짧아서 접기가 깨진다.
            #   표제처럼 보이는 것만(SECTION_PAT) 따로 세운다.
            _looks_head = len(ln) <= MAX_HEADING_LEN and any(
                re.match(p, ln, re.I) for _, p in SECTION_PAT)
            if re.match(r"^(#{1,6}\s|[-*+]\s|\d+[.)]\s|\||>)", ln) or _looks_head:
                if buf:
                    paras.append(" ".join(buf))
                    buf = []
                paras.append(ln)
                continue
            buf.append(ln)
        if buf:
            paras.append(" ".join(buf))
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


def check_absence(paras):
    """⑤ 방법 절의 «부재» 서술. 수치가 있으면 «발견»일 수 있으므로 제외한다.

    [FAIL] 이 아니라 [의심] 이다. 부재를 밝혀야 정당한 자리가 있기 때문이다 —
    독자가 없으면 결과를 «틀리게» 읽는 경우(예: 두 자료가 같은 사람이라고 오독).
    가르는 질문은 [[core]] §2⑧ 의 둘: ①없으면 독자가 「왜?」라고 묻거나 오독하는가
    ②남긴다면 «대상의 성질»인가 «내 결정의 서술»인가.
    """
    hits = []
    for sec, par in paras:
        if sec != "methods":
            continue
        for s in sentences(par):
            if re.search(r"\d", s):
                continue
            for pat, why in ABSENCE_PAT:
                if re.search(pat, s, re.I):
                    hits.append(("의심", sec, s, "⑤방법 절의 부재 서술: " + why))
                    break
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


# ⑤는 [의심] 등급이라 위 FAIL 기반 자기시험으로는 검증되지 않는다. 자기 목록을 갖는다.
# GOOD 은 «걸리면 안 되는» 것 — 절이 다르거나(결과의 부정은 발견이다), 자산이 목적어이거나,
# 수치를 동반하는 문장. BAD 는 반드시 걸려야 하는 것.
SELFTEST_ABSENCE_GOOD = [
    ("methods", "Those who left teaching are routed to a separate questionnaire and do not "
                "contribute to the analyses of exposure or exhaustion."),
    ("results", "The two sources were never combined."),
    ("methods", "Exposure was taken from the preceding wave."),
    ("methods", "결측이 있는 관측은 분석에서 제외하였다."),
    ("discussion", "두 자료원은 같은 사람이 아니다."),
]
SELFTEST_ABSENCE_BAD = [
    ("methods", "The two sources were never combined."),
    ("methods", "The schoolteachers are wider than the panel, so the two sources are not the "
                "same people and are never pooled."),
    ("methods", "No attempt was made to pool the surveys."),
    ("methods", "두 자료원은 결합하지 않았다."),
]


# ---------------------------------------------------------------- ⑦ «고칠 수 있는» 제한점
# 2026-09-21 한 원고 실측. 제한점에 「병원 자료(2026년)와 인구 자료(2023년)의 시점도 일치하지 않는다」가
# 있었다. 연구자: 「시점을 그냥 일치하면 되는데 왜 일치를 안 하고 뒤에도 한계점으로 두지? … 제한점으로
# 들어가면 절대 안 된다.」 열어 보니 같은 달(2026-08)의 인구가 통계포털에 이미 있었고, 그것으로 바꾸자
# 행정구역 개편 때문에 빼 두었던 지역까지 되살아났다. 원인은 빌려 온 코드의 주석(「인구 파일의 마지막 해」)을
# 열어 보지 않고 물려받은 것 — [[unverified-premise-compounding]] 의 원고 판본이다.
# ⭐ 규칙: **제한점은 «고칠 수 없는 것»만 적는 자리다.** 자료를 다시 받으면 없어지는 결함은 제한점이 아니라 할 일이다.
# ⚠ 좁게 건다 — 시점·연도·기간의 «불일치»만. 「단면 자료」「추적 기간이 짧다」는 설계의 성질이라 안 건다.
FIXABLE_PAT = [
    (r"(시점|연도|기준\s*연도|기준일|조사\s*시기)[^.。]{0,25}(일치하지|맞지\s*않|다르[다며고]|달랐|차이가\s*있)",
     "자료 시점 불일치를 제한점으로 적었다 — 같은 시점의 자료를 구할 수 없는지 «먼저» 확인했는가"),
    (r"\(\s*(19|20)\d\d\s*년?\s*\)[^.。]{0,30}\(\s*(19|20)\d\d\s*년?\s*\)[^.。]{0,30}(일치|다르|달랐|차이)",
     "두 자료의 연도가 다르다는 고백 — 맞출 수 있는지 확인했는가"),
    (r"(differ(ed|ent)?|mismatch(ed)?|not (the )?same|do(es)? not match)[^.]{0,40}\b(year|time ?point|period|date)s?\b",
     "data-timing mismatch confessed as a limitation — was a same-period source checked first?"),
    (r"\b(year|time ?point|period|date)s?\b[^.]{0,40}(differ(ed)?|mismatch(ed)?|did not match|do not match)",
     "data-timing mismatch confessed as a limitation — was a same-period source checked first?"),
]
SELFTEST_FIXABLE_BAD = [
    ("discussion", "본 연구는 단면 자료를 사용하였으며, 병원 자료(2026년)와 인구 자료(2023년)의 시점도 일치하지 않는다."),
    ("discussion", "둘째, 노출 자료와 결과 자료의 기준 연도가 서로 달랐다."),
    ("discussion", "Second, the reference years of the hospital and population data did not match."),
]
SELFTEST_FIXABLE_GOOD = [
    ("discussion", "첫째, 본 연구는 단면 자료를 사용하였으므로 인과 방향을 판단할 수 없다."),
    ("discussion", "추적 기간이 짧아 장기 효과를 평가하지 못하였다."),
    ("methods", "병원 자료와 인구 자료는 모두 2026년 8월 말 기준이다."),
    ("results", "연도에 따라 추정치가 달랐다."),
    ("discussion", "이 연구는 자료 연도와 공변량도 본 연구와 달랐다."),          # 선행연구와의 비교 — 제한점이 아니다
]


def check_fixable(paras):
    """⑦ 고칠 수 있는 제한점. [의심] — 정말 구할 수 없는 자료일 수 있어 사람이 판정한다."""
    hits = []
    for sec, par in paras:
        if sec != "discussion":
            continue
        for s in sentences(par):
            if re.search(r"(선행|이\s*연구는|의\s*연구|previous|prior stud)", s, re.I):
                continue                      # 남의 연구와 견주는 문장은 대상이 아니다
            for pat, why in FIXABLE_PAT:
                if re.search(pat, s, re.I):
                    hits.append(("의심", sec, s, "⑦고칠 수 있는 제한점: " + why))
                    break
    return hits


def check_log(paras):
    """⑥ 방법 절의 검증·추출 일지. [의심] — 드물게 «정의»의 일부일 수 있어 사람이 판정한다."""
    hits = []
    for sec, par in paras:
        if sec != "methods":
            continue
        for s in sentences(par):
            for pat, why in LOG_PAT:
                if re.search(pat, s, re.I):
                    hits.append(("의심", sec, s, "⑥검증·추출 일지: " + why))
                    break
    return hits


def selftest():
    """known-good을 잡으면 규칙이 틀린 것이다. 넓히기 전에 항상 이걸 돌린다."""
    bad = 0
    for sec, s in SELFTEST_LOG_GOOD:
        if check_log([(sec, s)]):
            bad += 1
            print("  ✗ ⑥이 known-good을 잡았다: %s" % s[:70])
    for sec, s in SELFTEST_LOG_BAD:
        if not check_log([(sec, s)]):
            bad += 1
            print("  ✗ ⑥이 잡아야 할 문장을 놓쳤다: %s" % s[:70])
    for sec, s in SELFTEST_FIXABLE_GOOD:
        if check_fixable([(sec, s)]):
            bad += 1
            print("  ✗ ⑦이 known-good을 잡았다: %s" % s[:70])
    for sec, s in SELFTEST_FIXABLE_BAD:
        if not check_fixable([(sec, s)]):
            bad += 1
            print("  ✗ ⑦이 잡아야 할 문장을 놓쳤다: %s" % s[:70])
    for sec, s in SELFTEST_ABSENCE_GOOD:
        if check_absence([(sec, s)]):
            bad += 1
            print("  ✗ ⑤가 known-good을 잡았다: %s" % s[:70])
    for sec, s in SELFTEST_ABSENCE_BAD:
        if not check_absence([(sec, s)]):
            bad += 1
            print("  ✗ ⑤가 잡아야 할 문장을 놓쳤다: %s" % s[:70])
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
          % (len(SELFTEST_GOOD) + len(SELFTEST_ABSENCE_GOOD) + len(SELFTEST_LOG_GOOD) + len(SELFTEST_FIXABLE_GOOD),
             len(SELFTEST_BAD) + len(SELFTEST_ABSENCE_BAD) + len(SELFTEST_LOG_BAD) + len(SELFTEST_FIXABLE_BAD), bad))
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
            + check_journalism(paras) + check_duplicate_numbers(paras)
            + check_absence(paras) + check_log(paras) + check_fixable(paras))

    secs = {s for s, _ in paras}
    print("학술 산문 문체 검사: %s" % os.path.basename(a.path))
    print("  문단 %d · 절 %s" % (len(paras), ", ".join(sorted(secs))))

    # 🔴 절을 못 찾으면 ②(방법 절 한정)가 «조용히» 안 돈다. 통과로 착각하지 않게 알린다.
    if "methods" not in secs:
        print("  ⚠ 방법 절을 찾지 못했다 — ②「선택의 이유」·⑤「부재 서술」·⑥「검증·추출 일지」 검사가 «돌지 않았다».")
        print("    표제가 다른 형태일 수 있다. 통과로 읽지 말 것.")

    if not hits:
        print("  통과 — 일곱 부류 모두 걸리지 않았다"
              + ("" if "methods" in secs else " (단, ②·⑤·⑥은 안 돌았다)"))
        return 0

    nf = sum(1 for h in hits if h[0] == "FAIL")
    for lvl, sec, s, why in sorted(hits, key=lambda x: (x[0] != "FAIL", x[1])):
        print("  [%s] (%s) %s\n        %s" % (lvl, sec, why, s[:110]))
    print("\n  FAIL %d · 의심 %d" % (nf, len(hits) - nf))
    print("  ⚠ [의심]은 후보다 — 원문을 열어 사람이 판정한다. 근거 = manuscript_rules.md")
    return 1 if (nf or (a.strict and hits)) else 0


if __name__ == "__main__":
    sys.exit(main())
