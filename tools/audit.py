#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit.py — 산출물 하나에 «적용 가능한 감사를 전부» 돌리는 단일 진입점.

왜 있나 (2026-08-10)
  감사 도구가 세 곳에 흩어져 있다 — `agent/tools/`(랩 공용), 각 스킬의 `scripts/`(공개 배포),
  `claude-config/agents/`(감사 에이전트). 어떤 산출물에 어떤 검사가 적용되는지가 **사람 기억과
  산문 문서**에만 있었고, 그래서 실제로 잊혔다: `agent/tools/`의 감사 스크립트들은 pptx-editing
  스킬의 SKILL.md에 **단 한 번도 적혀 있지 않았다**(2026-08-10에야 추가).

  같은 날 산문 드리프트 버그도 두 개 나왔다 — README가 "§1–§10"이라는데 가이드는 12절,
  SKILL.md 워크플로에 "7단계"가 두 번. **어떤 검사가 존재하는가를 산문으로 관리하면 낡는다.**

**장르로 가른다. 확장자가 아니다 (2026-08-11 재설계)**
  초판은 확장자로 검사를 골랐는데 사용자가 바로 짚었다 — *"피피티도 포스터도 원고도 문서요약도
  보고서도 다 같은걸로 잡게되네 그럼."* 맞는 지적이고, 증거가 이미 코드 안에 있었다:
  `.pptx` 안에서 포스터만 빼내려고 `is_poster()` 특례를 넣은 것 자체가 장르가 새어나온 신호였다.

  같은 확장자, 다른 규칙의 실제 사례:
    .pptx  슬라이드(본문 16pt) vs 포스터(본문 24pt). 프로파일을 잘못 고르면 18pt 포스터가 통과한다
           — 2026-08-11까지 실제로 그렇게 돌고 있었다.
    .pptx  회의 덱은 "출처: xxx.csv"가 **바람직한** 표기지만, 학술 포스터에선 내부 잔재다.
           같은 surface 검사가 한쪽에선 오탐이 된다.
    .docx  저널 투고 원고(booktabs·세로줄 금지) vs 국문 브리프(격자·zebra가 **의도된 디자인**).
           CORE.md가 명시적으로 경고하는 지점 — 저널 표 규칙을 브리프에 적용하면 멀쩡한 관습을 망가뜨린다.
    .hwpx  정부 보고서는 논문 문법도 저널 표 규칙도 넘어오지 않는다(report-audit 에이전트 담당).

  그래서 REGISTRY는 **장르**로 키를 잡고, 확장자는 장르 추정의 힌트로만 쓴다. 추정이 애매하면
  가정한 장르를 **찍어서 보여주고** `--genre`로 뒤집을 수 있게 한다. 조용히 추측하지 않는다.

설계 원칙
  - **도구를 감추지 않는다.** 인자가 필요해 자동 실행 못 하는 검사는 «수동»으로 보고하되
    **실행할 정확한 명령줄을 같이 출력한다.**
  - **커버리지를 보이게 한다.** 돌린 것 / 건너뛴 것 / 그 이유를 전부 찍는다.
  - **장르 때문에 끈 검사는 이유를 말한다.** 조용히 빠지면 «검사했다»는 착각이 남는다.
  - **PASS/FAIL 이분법을 쓰지 않는다 — [주의]가 셋째 상태다** (2026-08-22 신설).
    여러 검사기가 문서에 「완전 자동판정이 아니라 [의심] 후보를 추려 사람이 보게 한다」고
    적어 두고 exit 0으로 끝난다(term_collision · display_items · academic_prose). 이분법이면
    그 후보가 [PASS]에 묻혀 «통과 7 · 실패 0»만 남고, 읽는 사람은 볼 것이 남았다는 사실을
    모른다. 도구의 exit code는 건드리지 않고(다른 호출자가 깨진다) **출력 줄머리의 [의심]
    마커를 세어** [주의]로 승격한다. 실패가 아니므로 종료 코드는 0이지만 «남았다»는 보인다.
  - **공개 스킬은 이 파일에 의존하지 않는다.** 의존 방향은 항상 이쪽 → 스킬이다.

사용법
  python agent/tools/audit.py FILE [FILE ...]
  python agent/tools/audit.py FILE --genre brief      # 추정이 틀렸을 때
  python agent/tools/audit.py --list                  # 장르 x 검사 전체 지도
  python agent/tools/audit.py --self-check            # 등록된 도구 실재 확인

종료 코드: 실패한 검사가 하나라도 있으면 1
"""
import argparse
import os
import re
import shutil
import subprocess
import tempfile
import sys

if hasattr(sys.stdout, "buffer"):
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
CLAUDE = os.path.dirname(os.path.dirname(HERE))
SKILLS = os.path.join(CLAUDE, "claude-config", "skills")


def _tool(*parts):
    return os.path.join(HERE, *parts)


def _skill(name, *parts):
    return os.path.join(SKILLS, name, "scripts", *parts)


def _has_powerpoint():
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def _find_ko_corpus(target):
    """국문 «용어» 대조용 목표 저널 코퍼스(`*_ko_corpus.txt.gz`)를 위로 올라가며 찾는다.
    못 찾아도 검사는 돈다 — 교과서 코퍼스만으로도 「미등재·다른 뜻」은 보인다.
    여러 개면 앞의 셋까지 함께 넘긴다(저널 코퍼스는 겹쳐 봐도 해가 없다: 대조 폭만 넓어진다).
    만드는 법은 ko_term_corpus.py 머리말."""
    import glob as _g
    d = os.path.abspath(target if os.path.isdir(target) else os.path.dirname(target))
    for _ in range(6):
        hits = sorted(_g.glob(os.path.join(d, "**", "*_ko_corpus.txt.gz"), recursive=True))
        if hits:
            out = []
            for h in hits[:3]:
                out += ["--venue-file", h]
            return out
        nd = os.path.dirname(d)
        if nd == d:
            break
        d = nd
    return None


def _find_style_corpus(target):
    """대상에서 위로 올라가며 `*_style_corpus.json` 을 찾는다.
    찾으면 journal_fit 에 넘길 인자, 없으면 None.
    ⛔ 「가장 가까운 것 하나」만 쓴다 — 여러 개면 프로젝트가 목표 저널을 안 정한 것이고,
       그때는 조용히 고르지 말고 [수동] 으로 떨어뜨려 사람이 정하게 한다."""
    import glob as _g
    d = os.path.abspath(target if os.path.isdir(target) else os.path.dirname(target))
    for _ in range(6):
        hits = sorted(_g.glob(os.path.join(d, "**", "*_style_corpus.json"), recursive=True))
        if len(hits) == 1:
            return ["--corpus", hits[0]]
        if len(hits) > 1:
            return None          # 모호하면 사람에게
        nd = os.path.dirname(d)
        if nd == d:
            break
        d = nd
    return None


# ---------------------------------------------------------------------------
# GENRES — 장르마다 폰트 프로파일과 «이 장르에선 끄는 검사»가 다르다.
#   off  : {검사이름: 끄는 이유}. 이유를 반드시 적는다 — 조용히 빠지면 검사한 줄 안다.
# ---------------------------------------------------------------------------
GENRES = {
    "slide":      dict(ext=[".pptx"], profile="slide-ko", label="학회 발표 슬라이드", off={}),
    # 내부 회의 덱은 «출처: xxx.csv»가 오히려 바람직한 register다. 학회 덱·포스터에서
    # 같은 문자열은 내부 잔재이므로, 장르로 갈라 provenance만 끈다(검사 자체는 유지).
    "meeting":    dict(ext=[".pptx"], profile="slide-ko", label="내부 회의 덱", off={},
                       extra={"surface_text": ["--skip", "provenance"]}),
    # ⭐ 학생에게 «배포되는» 수업 슬라이드 (2026-09-09 신설). 학회 덱과 둘이 다르다.
    #   ① 내부 파일명이 정상이다 — 학생이 `sample_1500B.csv` 를 직접 쳐야 한다. 안 끄면
    #      한 덱에서 provenance 오탐 12건이 나고, 그만큼 나면 도구를 안 쓰게 된다.
    #   ② 대신 «운영 레지스터»가 샌다. 조교·교수의 말(시험 운영·채점 기제·선행 참조)이
    #      학생 화면에 남는 것이 이 장르 고유의 결함이다. 2026-09-09 한 수업 2주차 덱
    #      피드백 11건 중 4건이 그것이었고 하나는 본문에 박힌 「시험 출제 범위」였다
    #      (사용자: 「학생이 보는 것과 조교나 교수가 보는 것은 엄연히 다르다」).
    #   ⚠ 목록은 좁게 — 「검정은 3주차에 배운다」 같은 정당한 범위 표시는 잡지 않는다.
    "class":      dict(ext=[".pptx"], profile="slide-ko", label="학생 배포 수업 슬라이드",
                       off={},
                       extra={"surface_text": [
                           "--skip", "provenance",
                           "--extra-pattern",
                           "시험 ?출제 ?범위|출제 ?범위|시험에 나온다|과제에 필요한|"
                           "채점은 이 열|칸이 없다|에서 이 순서를 쓴다|에서 배울 예정|"
                           "주차 이론에서"]}),
    "poster":     dict(ext=[".pptx"], profile="poster", label="학술대회 포스터",
                       off={"render_overflow":
                            "하단 푸터 띠가 설계상 존재 — 슬라이드용 «아래 가장자리» 휴리스틱이 항상 FAIL"}),
    "manuscript": dict(ext=[".docx"], profile="doc", label="저널 투고 원고", off={}),
    # ⭐ 읽기용 통합본(totale) — 저자·공저자가 «한 창에서» 읽는 사본 (2026-09-10 신설).
    #   제출본과 «같은 소스, 다른 프로파일»이고 두 가지가 설계상 다르다:
    #   ① 출처 워크북 경로와 빌드 이력을 «일부러» 싣는다(9pt 회색) — 어느 xlsx 에서 온
    #      표인지가 이 사본의 용도다. 그래서 한글·내부경로가 정상적으로 나온다.
    #   ② 표를 워크북에서 렌더하므로 「Table 1.」 «캡션»이 없다. 본문이 Table 1–4 를
    #      인용하는데 그 라벨이 문서에 없어 display_items 가 전부 «유령»으로 신고한다.
    #   ⛔ 이 둘을 실패로 두면 totale 감사가 «항상» 실패 2로 끝난다 — 늘 뜨는 경고는
    #      읽히지 않고, 그러면 진짜 실패가 그 소음에 묻힌다.
    #   ⚠ 끄는 것은 이 둘«뿐»이다. 제출본의 결함은 대개 totale 에도 있으므로 나머지
    #      (폰트·표폭·용어충돌·문자일관성·문체·수사)는 그대로 돌린다.
    "reading":    dict(ext=[".docx"], profile="doc", label="읽기용 통합본(totale)",
                       off={"docx_surface":
                            "totale 은 출처 워크북 경로·빌드 이력을 설계상 싣는다(읽는 사본). "
                            "그 검사는 «제출본»에서 돌린다 — 거기서는 0이어야 한다",
                            "display_items":
                            "totale 은 표를 워크북에서 렌더해 「Table N.」 캡션이 없다. "
                            "본문의 Table 1–4 인용이 전부 «유령»으로 잡힌다(프로젝트 제약 28). "
                            "표시물 대조는 제출본에서 한다"}),
    # 보충자료는 «저널 원고와 같은 규율»을 받되 한 가지가 구조적으로 다르다: 자기 표를
    # 자기가 인용하지 않는다. Table S1~S8 은 «본문 원고»가 부르고, 보충이 "Table 2"라고
    # 쓰면 그건 본문 표를 가리키는 정당한 상호참조다. display_items 는 문서를 자족적인
    # 것으로 보므로 그 둘을 각각 「미인용」·「유령」으로 신고한다 — 2026-08-31 psy 에서
    # 실측: 8개 전부 본문 원고에서 인용되고 있는데 8건 전부 «NEVER CITED»로 떴다.
    # ⚠️ 끄는 것은 display_items «하나»다. 나머지(표 폭·용어 충돌·문자 일관성·폰트)는
    #    그대로 돈다 — 보충도 심사자가 읽는 지면이다.
    "supplement": dict(ext=[".docx"], profile="doc", label="저널 보충자료",
                       off={"display_items":
                            "보충의 표는 «본문 원고»가 인용한다. 자족 문서로 검사하면 전 표가 "
                            "«미인용»으로, 본문 표 참조가 «유령»으로 잡힌다. 실제 인용 여부는 "
                            "build_supplement.py 의 자기검사가 본문과 대조한다"}),
    "brief":      dict(ext=[".docx"], profile="doc", label="국문 브리프·연구요약",
                       off={"verify_tables":
                            "저널 표 검사다. 브리프는 격자·zebra가 의도된 디자인이라 규칙이 다르다"}),
    "report":     dict(ext=[".hwpx"], profile="doc", label="정부·기관 보고서", off={}),
    # ⭐ 창업 갈래(2026-09-11 신설). 연구 산출물과 실패 모드가 다르다 — 상대가 «회사 밖»이고
    #   보낸 것은 안 돌아온다. 저장소 표시(«» ⭐⛔ **굵게**)가 그대로 나가면 그게 첫인상이다.
    "outbound":   dict(ext=[".md", ".txt"], profile="doc", label="대외 문안(메일·공문·신청서)",
                       off={"display_items": "편지에는 표·그림 번호 체계가 없다",
                            "text_consistency": "참고문헌 번호 검사다. 문안에는 해당 없다"}),
    "text":       dict(ext=[".md", ".txt", ".tex", ".pdf"], profile="doc", label="메모·원고 소스", off={}),
    # ⭐ 초록은 «메모»가 아니다 (2026-09-05 신설). 250단어가 본문 3,500단어보다 많이 읽히고,
    #   심사의 첫 관문이며, 형제 산출물(원고·학회초록·포스터)과 숫자가 어긋나기 가장 쉬운
    #   지면이다. 그런데 확장자가 .md 라 text(메모·원고 소스)로 떨어져 «가장 약한 프로파일»을
    #   받고 있었다 — 한 학회 초록이 방어적 문장을 세 판 동안 달고 통과한 자리다.
    #   ext=[] 인 것은 확장자가 아니라 «이름»으로 갈리기 때문이다(detect_genre 참조).
    "abstract":   dict(ext=[], profile="doc", label="저널·학회 초록", off={}),
    # ⭐ 커버레터 (2026-09-08 신설). 그 전에는 .docx 기본값인 manuscript 로 떨어져
    #   «저널 논문 코퍼스»와 문체를 비교당했다 — 실측: em-dash 114.9/10k 가 「코퍼스 범위
    #   밖」으로 떴는데, 편지는 논문이 아니므로 그 비교 자체가 전제 오류다(이 파일이 이미
    #   아는 함정: 전제가 맞지 않는 대상을 검사에 넣으면 정상이 결함으로 찍히고, 그러면
    #   시끄러운 검사가 되어 아무도 안 읽는다).
    #   ⛔ 끄는 것은 «전제가 다른» 셋뿐이다. 표면·용어·문자 일관성·폰트는 그대로 돈다 —
    #      편집자가 «가장 먼저» 읽는 지면이다.
    "letter":     dict(ext=[], profile="doc", label="투고 커버레터",
                       off={"journal_fit":
                            "편지는 논문이 아니다. 저널 논문 코퍼스와 문체를 비교하면 "
                            "정상 편지가 늘 이상치로 뜬다",
                            "verify_tables":
                            "표가 없다(있으면 그것이 이상하다)",
                            "display_items":
                            "표·그림이 없다. 자족 문서로 검사하면 전 인용이 «유령»이 된다"}),
    # 산출물이 아니라 «들어오는 자료원». 다른 장르가 "내가 만든 걸 감사"한다면 이건
    # "남이 준 걸 쓰기 전에 감사"다 -- 2026-08-18 사고로 신설: 열 이름은 정확히 맞는
    # 과거력 체크리스트 31열을 «제외기준의 구조화된 소스»로 반겼는데, 전 6,260행이
    # 미입력(전부 '0')인 죽은 템플릿 열이었다. 그대로 실었으면 "전원 음성"이라는 틀린
    # 데이터가 됐다. 헤더 존재 ≠ 값 존재.
    "data":       dict(ext=[".xlsx", ".csv"], profile="doc", label="입력 자료원(엑셀·CSV)", off={}),
    # ⭐ 수업 문항·케이스 (2026-09-07 신설). 그 전에는 .docx 기본값인 manuscript 로 떨어져
    #   목표 저널·xlsx 표 검사가 매번 [수동]으로 남았고, 그래서 audit.py 가 «항상»
    #   「감사 미완료」를 찍었다 — 늘 뜨는 경고는 읽히지 않는다. 수업자료에는 목표 저널도
    #   output/tables 도 «구조적으로» 없으므로 이유와 함께 끈다.
    #   대신 이 장르에만 appraisal(근거평가 도구 규격)이 붙는다.
    "teaching":   dict(ext=[], profile="doc", label="수업 문항·케이스·모범답안",
                       off={"journal_fit":
                            "투고 원고가 아니다 — 목표 저널이 없다. 대상 독자는 학생·교수다",
                            "verify_tables":
                            "표를 xlsx 로 내는 파이프라인이 아니다(output/tables 가 없다). "
                            "수치는 verify_numbers 류가 «논문 원문»에 직접 대조한다"}),
    # ⭐ 학습 읽기가이드 (2026-09-30 신설) — 연구자가 «배우는» 쪽의 교재·논문 정리본.
    #   그 전에는 .docx 기본값인 manuscript 로 떨어져, 빌드할 때마다 원고 전제의 검사
    #   둘이 «항상» 실패했다(2026-09-29 보건의료경제학 Wk2~9 여덟 편 전부).
    #   ① new_in_conclusion — 첫 절이 「결론부」로 잡혀 원문 요약 문장이 전부 신고됐다.
    #   ② display_items — 원문 그림은 «이미지»로 싣고 캡션 번호가 없다. 본문이 다른 주차
    #      원문의 Figure 1 을 이름으로 부르면 매번 «유령»으로 잡혔다.
    #   ⛔ 늘 뜨는 실패는 안 읽힌다 — 그러면 진짜 결함(작업 표식, 표 폭, 폰트)이 묻힌다.
    #   ⚠ 끄는 것은 이 둘«뿐». 표면·문자 일관성·문체·용어는 그대로 돈다 — 9pt 로 급히
    #     읽는 문서라 표면 결함이 더 비싸다. docx_surface 는 teaching 에 없지만 여기엔 건다
    #     (⭐·⚠ 표식이 실제로 이 장르에서 샜다).
    "study":      dict(ext=[], profile="doc", label="학습 읽기가이드(교재·논문 정리)",
                       off={"new_in_conclusion":
                            "원고의 결론부 구조가 없다. 절마다 원문 요약이라 첫 절이 결론부로 "
                            "잡혀 원문 문장이 신고된다",
                            "display_items":
                            "원문 그림을 이미지로 싣고 캡션 번호가 없다. 다른 주차 원문의 "
                            "Figure 를 이름으로 부르면 «유령»으로 잡힌다"}),
}

# ---------------------------------------------------------------------------
# REGISTRY — 장르별로 «무엇을 돌리는가». 이 표가 정본이다.
# ---------------------------------------------------------------------------
REGISTRY = [
    dict(genres=["slide", "meeting", "class", "poster"], name="text_fit",
         path=_skill("pptx-editing", "audit_text_fit.py"), auto=True, needs="powerpoint",
         catches="텍스트가 슬라이드 밖으로 나가거나 다른 도형 위에 얹힘 (PowerPoint 실측). "
                 "표 셀은 못 본다 -> render_overflow가 담당"),
    dict(genres=["slide", "meeting", "class", "poster"], name="surface_text",
         path=_skill("pptx-editing", "audit_surface_text.py"), auto=True,
         catches="표면에 남은 편집 해명·내비게이션·재진술·내부 파일명, Figure/Table 번호 drift. "
                 "회의 덱에선 provenance가 오탐일 수 있다(출처 표기가 바람직한 register)"),
    # ⭐ 랩미팅 3부 구성의 «논문 발표» 파트 (2026-09-15 신설). 감사 전부가 통과한 덱에서
    #   논문 요약 7장을 사용자가 먼저 봤다 — 넘침·숫자·말투 검사는 «몇 장인가»를 안 본다.
    #   detect_genre 는 .pptx 를 대개 slide 로 잡으므로 slide 에도 건다. 라벨이 없으면 no-op.
    dict(genres=["slide", "meeting"], name="journal_club",
         path=_tool("audit_journal_club.py"), auto=True,
         catches="논문 발표 파트(라벨 Journal club·The paper·저널클럽…)가 표지 캡처 한 장인가 — "
                 "논문의 설계·결과·표를 슬라이드로 옮긴 것. 논문은 원문 PDF 로 발표한다"),
    # 2026-09-30 한 코호트 결과 덱: 연구자가 덱을 넘기며 거듭 잡던 관례를 «덱 파일만 보고» 판정한다.
    #   전에는 그 덱의 빌더 게이트에만 있어 다음 덱이 처음부터 다시 배웠다. 첫 판은 제목 글자크기를
    #   run 에서만 찾아 kit_mono 덱의 제목을 0 개 찾았고 «모든 검사가 빈 채로 PASS» 였다 — 고쳤다.
    dict(genres=["slide", "meeting", "class"], name="deck_conventions",
         path=_tool("audit_deck_conventions.py"), auto=True,
         catches="쪽 번호가 순서와 다름(장을 옮긴 뒤) · 두 줄로 넘친 제목 · 노트 언어 섞임 · 인용 안 한 참고문헌 · "
                 "풀이 없는 약어. 코호트 이름 같은 고유 약어는 --known 으로 넘긴다"),
    dict(genres=["poster"], name="poster_selfcontained",
         path=_tool("audit_poster_selfcontained.py"), auto=True,
         catches="캡션이 길어 안 읽히는 것(렌더 줄 수 실측) + 표 머리행이 도입하는 구성개념이 "
                 "포스터 산문에 없는 것(「저건 어디서 갑자기 튀어나온거야?」). 포스터 전용 -- "
                 "덱은 캡션 관습이 다르고 원고는 Methods 가 본문에 따로 있어 전제가 다르다"),
    dict(genres=["slide", "meeting", "class", "poster"], name="xml_integrity", path=_tool("deck_audit.py"), auto=True,
         catches="구역·줌 링크·creationId 중복·목차↔구분장 제목 불일치·고아 media"),
    dict(genres=["slide", "meeting", "class", "poster"], name="render_overflow", path=_tool("deck_render_audit.py"),
         auto=True, needs="powerpoint",
         catches="**표 행높이 확장으로 인한 잘림.** 좌표·XML 검사가 전부 통과시키는 유형"),
    dict(genres=["slide", "meeting", "class", "poster", "manuscript", "reading", "supplement", "brief", "teaching", "letter", "study"], name="font_sizes",
         path=_tool("audit_font_sizes.py"), auto=True, profiled=True,
         catches="글자 크기 바닥값 미만 [RUN] + 그림 축소배치로 구워진 글씨가 줄어든 것 [SCALE]"),
    # .pptx/.docx만 지원(.hwpx 아직 없음) -- report 장르는 제외.
    dict(genres=["slide", "meeting", "class", "poster", "manuscript", "reading", "supplement", "brief", "teaching", "letter", "study"], name="term_collision",
         path=_tool("audit_term_collision.py"), auto=True,
         catches="각주(†)로 정의한 용어가 다른 곳에서 다른 뜻으로 재사용됨. 완전 자동판정이 "
                 "아니라 [의심] 후보를 추려 사람이 보게 한다 -- 놓치는 것보다 과잉표시가 낫다"),
    # python-docx로 열기 때문에 .pptx를 주면 ValueError로 죽는다 (2026-08-10 확인).
    dict(genres=["manuscript", "reading", "supplement", "brief", "teaching", "study"], name="table_widths", path=_tool("audit_table_widths.py"),
         auto=True, catches="렌더 전에 «숫자가 두 줄로 쪼개질» 표 열을 예측"),
    # surface_leak(사람이 짠 term 목록)·rhetoric(형태)과 축이 다르다: 이쪽은 «빌더가 걷어낸다고
    # 믿었지만 실제로는 남은 것»을 만들어진 파일에서 본다. 2026-09-07 한 원고 실측:
    # 조립기의 주석 제거 규칙이 절 파일 «둘»에만 있어 나머지 넷의 한국어 103줄이 매번 통합본에
    # 실렸고, 참고문헌 추출기는 검사 «순서» 때문에 「인용키」 줄의 앞부분을 남기고 있었다.
    # 소스만 보는 검사는 둘 다 못 본다 -- 소스에는 규칙이 «있었기» 때문이다.
    dict(genres=["outbound", "letter"], name="outbound_marks",
         path=_tool("audit_outbound_marks.py"), auto=True,
         catches="«문안 블록» 안에 남은 작업 표시 — 경고·이모지 기호, «», 마크다운 굵게"
                 "(붙여넣으면 별표가 그대로 보인다), 문장 속 화살표, 내부 파일 경로. "
                 "블록 «밖»(우리끼리 읽는 메모)은 일부러 보지 않는다 — 거기선 표시가 일을 한다"),
    # ⭐ 문서 속성 (2026-09-21 신설). 본문을 보는 검사는 «전부» 통과하는 자리다 — 한 원고 투고
    #   메일을 보내기 직전에 두 원고 docx 의 속성이 creator="python-docx" · description="generated by
    #   python-docx" 인 것이 드러났다. python-pptx 기본값은 lastModifiedBy="Steve Canny" 다(selftest 실측).
    #   심사용(파일명에 blind·심사용) 파일은 사람 이름이 «하나라도» 있으면 결함 — 빌더가 비워도 사람이
    #   열어 저장하면 계정명이 다시 박히므로 «보내기 직전»에 돌아야 한다. 지원 밖 확장자는 건너뛰므로
    #   나가는 장르 전부에 건다("data" 는 남이 준 파일이라 뺀다).
    dict(genres=["slide", "meeting", "class", "poster", "manuscript", "reading", "supplement", "brief",
                 "report", "abstract", "letter", "teaching", "study"], name="doc_properties",
         path=_tool("audit_doc_properties.py"), auto=True,
         catches="문서 속성(작성자·마지막 저장자·설명)에 남은 라이브러리 기본값(python-docx·Steve Canny·"
                 "openpyxl)·AI/도구 이름, 그리고 심사용 파일의 사람 이름. docx·pptx·xlsx·hwpx"),
    dict(genres=["manuscript", "reading", "supplement", "brief", "abstract", "letter", "study"], name="docx_surface",
         path=_tool("audit_docx_surface.py"), auto=True,
         hint="검토용 totale 은 작업 주석을 «일부러» 남긴다 → --reading-copy",
         catches="만들어진 .docx 의 «글자 노드»에 남은 한글·작업 표시 문자(⚠⛔«»)·내부 경로·"
                 "안 채운 대괄호. 한글은 국문 산출물이면 «자동으로» 건너뛴다"),
    dict(genres=["manuscript", "reading", "supplement", "brief", "report", "text", "abstract", "teaching", "letter", "study"], name="text_consistency",
         path=_tool("audit_text_consistency.py"), auto=True,
         catches="괄호 미종결·참고문헌 결번/미인용·단위/대시/p값/콤마 표기 혼용"),
    # text_consistency와 축이 다르다: 저쪽은 «참고문헌» 번호, 이쪽은 «표·그림» 번호다.
    # 2026-08-20 한 원고에서: 보충 표를 더하며 손으로 순서를 확인했더니 원래 있던 Table S1/S2가
    # 첫 인용 순서와 반대였고, 고친 판을 이 검사로 돌리니 Figure 1/2도 같은 상태였다 --
    # 사람이 물어봐 줘야만 드러나던 축이다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "teaching", "letter", "study"], name="display_items",
         path=_tool("audit_display_items.py"), auto=True,
         catches="표·그림 번호가 «본문 첫 인용 순서»와 어긋남 · 인용되지 않은 고아 표시물 · "
                 "표시물이 없는 유령 인용 · 결번. 각주가 그 문서의 형제들보다 튀는 것은 [의심]"),
    # auto=False로 «등록만» 해 둔다 — 목표 저널을 알아야 돌기 때문이다. 그게 요점이다:
    # 이렇게 두면 원고를 감사할 때마다 "목표 저널이 필요해서 안 돌았다"가 «인쇄»된다.
    # 2026-08-20 한 원고에서: 목표 저널이 미정인 채로 원고를 계속 고쳤고, 처음 재봤을 때
    # 세미콜론 145/10k(그 저널 27편의 «최대»가 80)에 1인칭 we는 Q1 미만이었다.
    # 안 정한 것이 조용하면 안 맞춘 것도 조용하다.
    # ⭐ 2026-09-08 — auto=False 의 «원래 이유»(목표 저널을 알아야 돌기 때문)는 그대로 두되,
    #   `--corpus` 캐시가 프로젝트에 있으면 «자동으로» 돈다(corpus_auto). 캐시가 저널명을
    #   품고 있고 네트워크를 쓰지 않기 때문이다(0.13초).
    #   ⛔ 이 구멍이 실제로 열려 있었다: 목표 저널이 정해져 있고 검사도 «돌렸는데», 결과가
    #      세션과 함께 사라져 투고를 결정하는 화면에 없었다. [수동] 표시는 「안 돌렸다」만
    #      말해 주고 「돌렸고 이런 잔여가 있다」는 남기지 못한다.
    #      → 캐시를 만드는 법: journal_fit.py --journal "..." --draft X --save-corpus \
    #                          references/<journal>_style_corpus.json
    dict(genres=["manuscript", "reading", "abstract", "teaching", "letter"], name="journal_fit", path=_tool("journal_fit.py"), auto=False,
         corpus_auto=True,
         hint='--journal "Lancet Healthy Longev"   (프로젝트 CLAUDE.md에 목표 저널을 박아둘 것)',
         catches="**목표 저널 실물 대비 «형태»와 «문체»** — 섹션별 분량 배분, 문장 길이, "
                 "1인칭 비율, 세미콜론·대시 빈도. 분량 제한이 없어도 돌린다(압축 도구가 "
                 "아니라 «그 저널 수준인가»의 측정기다). 용어는 esearch 빈도로 따로 확인"),
    # 암호 걸린 워크북은 --password가 필요해 auto=False. 암호가 없으면 그냥 돌아간다.
    # 인코딩은 «파일 하나»가 아니라 폴더 전체를 봐야 방향(소스=BOM 금지 / 산출물=BOM 필수)이
    # 드러나므로 auto=False. 대상에 파일을 주면 그 파일만, 폴더를 주면 트리 전체를 본다.
    # 「진실」이 무엇인지는 사람이 정해야 한다(--truth) -> auto=False.
    dict(genres=["slide", "meeting", "class", "poster", "manuscript", "reading", "supplement", "brief", "text", "report", "letter"],
         name="stale_numbers", path=_tool("audit_stale_numbers.py"), auto=False,
         hint="--truth <실행로그·산출물>   (대상 파일이 «먼저» 오는 위치인자다 — 사이에 `--` 를 넣으면 argparse 가 거부한다)",
         catches="**박아둔 숫자가 낡았는가.** 자료를 다시 만들면 표는 재계산되지만 «그 표를 "
                 "설명하는 문장»은 손으로 쓴 상수라 조용히 거짓이 된다(`최대 VIF = 1.09 "
                 "(<= 1.08)` 처럼 같은 줄에서 자기모순). 판정이 아니라 «후보»다 — 반올림 "
                 "차이·내장자료가 함께 걸린다"),
    # stale_numbers 의 «문장판». 숫자는 출처 대조를 강제하면서 인용문은 손으로 믿고
    # 있었다 -- 2026-09-10 하루에 두 번 걸렸다(낱말 중간 절단 `효율적일`→`효율적`,
    # 어미만 바꾼 의역 `꼽힌 점은…`→`꼽혔다`). 후자는 부분문자열 검사로도 안 잡힌다.
    # 출처가 무엇인지는 사람만 아므로(--source) auto=False -- 그래서 인용이 있는 문서를
    # 감사할 때마다 「출처를 안 줘서 안 돌았다」가 인쇄된다. 그게 요점이다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "report", "text", "abstract",
                 "teaching", "letter", "slide", "meeting", "class", "poster", "study"],
         name="quote_fidelity", path=_tool("audit_quote_fidelity.py"), auto=False,
         hint='--source <원고·보고서·표 등 따온 곳>',
         catches="**따옴표 안이 출처에 글자 그대로 있는가.** 낱말 중간에서 자른 것[경계]과 "
                 "어미·조사만 바꿔 놓고 따옴표를 씌운 것[의역]을 가른다. 못 찾은 것은 "
                 "[의심]으로만 -- 한국어 따옴표는 강조·용어 도입에도 쓰여 실패로 셀 수 없다"),
    # ⭐⭐ 값이 맞아도 «값들 사이의 관계»는 틀릴 수 있다 (2026-09-11, 한 위원회 발표덱).
    #   한 덱에서 최상급이 «네 번» 거짓이었다 — 「갈린 유일한 지점」(짝 문항 8개 중 여럿),
    #   「온도차가 가장 큰」(셋째), 「가장 높게 지지한」(2위), 「가장 짧다」(네 항목 동률).
    #   값은 전부 원본 표에서 왔고 **이 표의 검사 일곱 개가 전부 통과했다.** 그 축에 눈이 없었다.
    #   참·거짓은 원본을 아는 사람만 판정할 수 있으므로 auto=False — 그래서 이런 문장이
    #   있는 산출물을 감사할 때마다 「안 돌았다」가 인쇄된다. 그게 요점이다.
    dict(genres=["slide", "meeting", "class", "poster", "manuscript", "reading", "supplement",
                 "brief", "report", "text", "abstract", "teaching", "letter", "study"],
         name="absolute_claims", path=_tool("audit_absolute_claims.py"), auto=False,
         hint="<이 파일>   (인자 없음 — 그냥 돌린다)",
         catches="**「가장·유일·전부·1위·없었다」를 세어 보고 썼는가.** 값 대조는 «있는 값»만 "
                 "보고 조판 검사는 글자만 본다 — 순위·동률·유일성은 둘 다 못 본다. 판정이 "
                 "아니라 «목록»이고, 센 결과는 프로젝트 검증 스크립트에 assert 로 박는다"),
    # ⭐ 남의 국문 문서에 «이어 붙인» 부분이 그 문서와 맞는가 (2026-09-11).
    #   기존 산문 검사기(academic_prose·register·text_consistency)는 문서를 «혼자» 본다.
    #   이 축은 관계를 본다 — 원본을 줘야 하므로 auto=False. 그래서 원고를 감사할 때마다
    #   「원본을 안 줘서 안 돌았다」가 인쇄된다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "report", "text", "teaching",
                 "letter", "study"], name="graft_fit", path=_tool("docx_diff.py"), auto=False,
         hint="<원본.docx> <이 파일>   (축 5 가 그것이다. 축 1~4 는 덤)",
         catches="**이어 붙인 문단이 host 의 습관과 맞는가** — 접속어를 그 문서가 쓰지 않는 "
                 "자리에 썼는가, 그 절의 문단 길이 분포를 벗어났는가. 화제가 겉도는지는 "
                 "기계가 못 본다(어휘 겹침을 시험했으나 좋은 판에 켜지고 나쁜 판에 꺼졌다)"),
    dict(genres=["data", "text", "letter"], name="encoding", path=_tool("audit_encoding.py"), auto=False,
         hint="--data-ext .csv,.txt   (폴더를 주면 트리 전체)",
         catches="**BOM 방향 위반 3종.** ① .R/.py 소스에 BOM(Rscript가 즉사) ② 한글 있는 "
                 "산출 csv/txt에 BOM 없음(Excel이 CP949로 읽어 깨짐) ③ 읽기인데 bare "
                 "utf-8(BOM을 못 벗겨 \ufeff가 본문에 섞임). 한쪽을 고치면 다른 쪽이 "
                 "깨지므로 셋을 «함께» 본다"),
    dict(genres=["data"], name="column_profile", path=_tool("profile_columns.py"), auto=True,
         hint="--only-suspect  (암호 걸린 xlsx면 --password ... 추가)",
         catches="**죽은 열·상수 열·희소 열.** 헤더는 멀쩡한데 값이 없거나 전부 같은 값인 "
                 "컬럼 -- 그대로 분석에 실으면 «전원 음성» 같은 틀린 데이터가 된다. "
                 "새 자료원을 파이프라인에 붙이기 «전에» 돌릴 것"),
    # brief에도 등록해 두고 GENRES["brief"].off에서 «이유와 함께» 끈다. 목록에서 빼버리면
    # "브리프엔 왜 이 검사가 없나"에 답이 남지 않는다 — 보이는 제외가 조용한 누락보다 낫다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "teaching", "letter"], name="verify_tables", path=_tool("verify_tables.py"), auto=False,
         hint="--xlsx-dir output/tables",
         catches="렌더한 원고의 표를 «독자가 보는 대로» 검사 + 원본 xlsx의 모든 숫자가 실렸는지"),
    dict(genres=["report"], name="hwpx_verify", path=_skill("hwpx-editing", "verify.py"), auto=True,
         catches="hwpx 구조 무결성 (미주 중첩 등 한글로 열어야만 드러나는 부류 포함)"),
    # needs="hancom": 렌더 기준이라 한글 COM이 필요하다. 표시가 없으면 --fast가 이걸 훅에서
    # 돌리려 들고, 훅이 한글을 띄우면 느려서 결국 가드를 끄게 된다(hwpx_guard가 audit_layout을
    # 제외한 이유와 같다).
    dict(genres=["report"], name="hwpx_layout", path=_skill("hwpx-editing", "audit_layout.py"),
         auto=True, needs="hancom",
         catches="렌더 기준 레이아웃 — 넘침·강제개행·빈 페이지"),
    dict(genres=["report"], name="hwpx_typography", path=_skill("hwpx-editing", "audit_typography.py"),
         auto=False, hint="--expect-face 휴먼명조 --expect-body-pt 10",
         catches="본문 글꼴·크기가 기관 서식과 어긋남 (기대값을 알아야 검사 가능)"),
    dict(genres=["report"], name="hwpx_crossref", path=_skill("hwpx-editing", "crossref_check.py"),
         auto=True, catches="미주·재인용 번호 대응"),
    # 2026-08-20: report(.hwpx) 추가. 그 전까지 hwpx는 «표면 규율» 축에 검사가 0이었다 —
    # report_content_discipline이 통째로 그 얘기인데도. surface_leak_scan이 hwpx를 못 읽던 것이
    # 이유였고, audit_text_consistency.read_hwpx를 재사용해 열었다(추출기 사본을 만들지 않음).
    # ⭐ SIGN/GRADE 판정표를 담은 문서(EBM 수업 사례·근거요약·지침 초안)의 «산수»를 본다.
    # 초기 확실성 − Σ하향 + Σ상향 == 최종 확실성 인지 열마다 확인한다. 「1단계 하향」이라 써 놓고
    # 최종 등급은 안 내린 표가 실제로 있었다(2026-08-25 결핵 v5). 한 열에 결과가 둘 이상 묶여
    # 있고 하향이 그중 «하나»에만 걸릴 때 이 형태로 샌다 — 해법은 등급을 맞추는 것이 아니라
    # 열을 결과별로 쪼개는 것이다. 판정표가 없는 문서에서는 아무것도 출력하지 않으므로 auto 로 둔다.
    # 드리프트 검사(이유는 바뀌었는데 등급은 그대로)는 «--no-drift 로 꺼서» 부른다. 그 검사는
    # 이유를 고칠 때마다 뜨는 것이 정상이라 auto 로 두면 매번 실패가 되고 그러면 아무도 안 본다.
    # 쓸 때는 `--accept` 로 기준선을 저장한 뒤 판정을 고친 다음 같은 스크립트를 그냥 부른다.
    # ⭐ 근거평가 «도구»를 규격대로 썼는가 (2026-09-07 신설). grade_verdicts 가 GRADE 표의
    # 산수를 보는 자리의 «앞 단계»다 — SIGN 체크리스트·RoB 2·ROBINS-I·AMSTAR-2 의 항목번호와
    # «그 항목이 제공하는 체크박스»를 규격에 대고 본다. 2026-09-07 인사의4 실측: 자유기술
    # 칸(체크리스트2 의 1.8, 3 의 1.5)에 Yes/Does not apply 를 넣은 것을 수치검사·문체검사·
    # GRADE 산수가 «전부» 통과시켰다. 도구의 규격을 알아야만 보인다.
    # ⚠ text(.md/.txt/.tex/.pdf)에는 등록하지 «않는다» — 이 검사는 Word 표만 읽는다.
    #   등록했다가 README.md 에서 traceback + exit 1 이 나 [FAIL]로 잡혔다(2026-09-07).
    dict(genres=["teaching", "manuscript", "reading", "supplement", "brief"], name="appraisal",
         path=_tool("audit_appraisal.py"), auto=True,
         catches="근거평가 도구를 규격대로 썼는가 — 판정 칸이 아닌 항목에 판정을 넣었는가, "
                 "그 항목에 «없는» 선택지를 썼는가, 종합이 도메인/2.1 과 맞는가. "
                 "머리말이 규정한 예외는 [의심]으로 내려 사람이 판정한다"),
    # ⭐ 2026-09-09. 결론·환자설명에서 «크기 없는 변화»를 주장하는 버릇 — 사용자 지적:
    #   "학습자료의 문제가 아니라 너의 모든 output에서 종종 생기는 버릇이야."
    #   근거 없는 새 주장과 「절대차로 말한다」 규율을 한 번에 막는다. .docx 만 읽는다.
    dict(genres=["teaching", "manuscript", "reading", "supplement", "brief", "report", "abstract", "study"],
         name="new_in_conclusion", path=_tool("audit_new_in_conclusion.py"), auto=True,
         catches="결론·권고문·환자설명이 «크기 없는 변화»를 주장하는가 — "
                 "「조금 늘 수 있습니다」처럼 앞에 근거가 없고 크기도 안 대는 문장"),
    dict(genres=["teaching", "manuscript", "reading", "supplement", "brief", "text", "abstract"], name="grade_verdicts",
         path=_tool("audit_grade_verdicts.py"), auto=True, args=["--no-drift"],
         catches="GRADE 판정표의 산수 — 초기−Σ하향+Σ상향 ≠ 최종"),
    dict(genres=["text", "manuscript", "reading", "supplement", "brief", "report", "abstract", "teaching", "letter", "study"], name="surface_leak",
         path=_tool("surface_leak_scan.py"),
         auto=False, hint='--terms "이 메모는" "사용자가" "이전 버전"',
         catches="register별 표면 누출. **기본 목록이 없는 것이 설계** — 문서마다 안전한 표현이 다르다"),
    # ⭐ AI 티는 «모든 산출물»에 해당한다 — 덱만의 문제가 아니다. 덱·포스터는 surface_text(auto)로
    # 이미 덮여 있었는데 원고·보고서 쪽은 surface_leak이 auto=False라 사실상 안 돌았다. 그 auto=False의
    # 이유는 "호출자가 terms를 줘야 한다"였고, --rhetoric은 아무것도 안 받으므로 그 이유가 안 걸린다.
    # 그래서 같은 스크립트를 «인자 고정 + auto»인 별도 항목으로 한 번 더 등록한다.
    # 사정거리 근거는 surface_leak_scan.rhetoric_rules() 실측(원고 429줄, 오탐 0) 참조.
    # ⭐ 국문 산출물에 «내가 지어낸 말»이 새는 것 (2026-09-11 한 위원회 문서).
    #   하루에 넷이 샜고 그중 「한 걸음 두다」는 한국어에 없는 결합이었다. register 문제가
    #   아니라 «없는 말을 썼다»는 문제라 장르를 가리지 않는다. rhetoric 과 같은 이유로
    #   auto=True — 인자를 안 받는다. 목록은 좁다(실측 오탐 0: 원고·정리본·연구계획서·
    #   동의서·자유응답 422건). 늘릴 때는 오탐을 재고 늘릴 것 — surface_leak_scan.coined_rules().
    dict(genres=["manuscript", "reading", "supplement", "brief", "text", "report", "abstract",
                 "teaching", "letter", "slide", "meeting", "class", "poster", "study"], name="coined",
         path=_tool("surface_leak_scan.py"), auto=True, args=["--coined"],
         catches="**내가 지어낸 국문 표현** — 대화에서 만든 말이 산출물로 샌 것. "
                 "받는 사람에겐 맥락이 없다. 「소리 내어 읽어 걸리면 내가 만든 말」의 "
                 "기계 대응물 (세부 규율 = korean_prose_grafting.md)"),
    dict(genres=["manuscript", "reading", "supplement", "brief", "text", "report", "abstract", "teaching", "letter", "study"], name="rhetoric",
         path=_tool("surface_leak_scan.py"), auto=True, args=["--rhetoric"],
         catches="**AI가 쓴 티** — 카드 제목이 «수사적 기능»(The hook·Key point), 화살표·≠를 말 대신, "
                 "문장 속 대문자 강조. 내용이 아니라 «형태»라서 사람이 읽어선 잘 안 걸린다"),
    # 고백형 어투. 원본 .md/.tex를 본다(조판된 .docx가 아니라) — 고칠 곳이 원고이기 때문.
    # auto=False 인 이유는 대상이 빌드 산출물이 아니라 소스 원고라서, 경로를 호출자가 준다.
    # ⭐ 국문 학술 산문 문체. rhetoric/register가 «영문 AI 티»를 보는 자리에 대응하는 국문판이
    # 비어 있었다 — brief 장르는 검사가 둘뿐이었다. 규칙은 지도교수가 «실제로 지운» 문장에서만
    # 뽑았고(한 원고 2026-08-22), 넓히기 전에 `--selftest`로 known-good을 안 잡는지 본다.
    # ⚠ 감수자 수정본을 known-good으로 쓰지 않는다 — 「심하게 느껴지는 것만」 고치신 최소 수정본이다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "report", "text", "abstract", "teaching", "letter", "study"], name="academic_prose",
         path=_tool("audit_academic_prose.py"), auto=True,
         catches="**국문 학술 산문 문체** — ①예고·메타 문장(뒤에 나올 내용을 미리 알림) "
                 "②방법 절의 «이유» 서술(근거는 남기되 「~때문이다」를 사실 진술로) "
                 "③저널리즘 표현(사실이 아니라 분위기) ④Key findings 수치의 본문 중복 "
                 "⑤**방법 절의 «부재» 서술**(「안 한 것」은 방법이 아니라 해명이다 — 어휘가 "
                 "아니라 «구조»로 잡으므로 부정형·수동형도 걸린다). "
                 "[FAIL]은 근거가 확실한 것만, 나머지는 [의심]으로 사람에게 넘긴다"),
    # academic_prose 와 축이 다르다: 저쪽은 «문장», 이쪽은 낱말의 «소속»이다. 그리고 국문
    # 전용이다 — journal_fit·journal_term_census 는 PMC 영문 코퍼스라 국문 KCI 원고에는
    # 매번 「해당 없음」이 찍혔고, 그 자리가 2026-09-16 까지 비어 있었다.
    # 그날 실측(한 원고): 「명세」를 model specification 뜻으로 썼는데 교과서 16건·JAMCH
    # 3건이 전부 «진료비청구명세서»였다. 사전에는 있는 말이라 맞춤법·문체 검사로는 안 잡힌다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "report", "text", "abstract",
                 "teaching", "letter", "study"], name="ko_terms",
         path=_tool("audit_ko_terms.py"), auto=True, ko_corpus_auto=True,
         catches="**국문 용어가 그 분야 교과서·목표 저널의 말인가** — ①미등재(신조어 후보) "
                 "②다른 뜻(그 «글자»는 코퍼스에 있는데 «낱말»로는 안 쓰인다) ③자료원·제도의 "
                 "공식 명칭 위반(이것만 [FAIL], 나머지는 [의심]). 코퍼스는 ko_term_corpus.py 로 "
                 "만들고, 판정한 말은 ko_terms_ok.json 에 «이유와 함께» 적으면 조용해진다. "
                 "⚠ 못 보는 것: 같은 «낱말»로 쓰되 뜻만 다른 경우(불가분성 = 경제학 vs 「불가분의 관계」)"),
    # ko_terms 의 자매 — 저쪽은 낱말의 «소속», 이쪽은 원고의 «모양»과 «서술어». 2026-09-16 연구자 질문
    # (「표·그림 과잉? 어색한 한국어·AI 말투? 이 저널의 흐름? 참신성?」)에 기계가 하나도 답하고 있지 않았다.
    # 그날 실측: 내가 새로 쓴 「평균에 녹아 사라진다」「…까지 벌어졌다」가 저널 37편·교과서 4권에 0건인데
    # kprose·academic_prose·rhetoric·register 전부 통과 — «틀린 문장»이 아니라 «그 분야가 안 쓰는 말»이었다.
    dict(genres=["manuscript", "reading", "supplement", "brief", "abstract"], name="ko_venue",
         path=_tool("audit_ko_venue.py"), auto=True, ko_corpus_auto=True,
         catches="**국문 원고의 «모양»이 목표 저널과 맞는가** — ①표·그림 개수 vs 저널 분포(상위 10%% 초과) "
                 "②골격(서론→연구방법→연구결과→고찰→요약 · 「따라서 본 연구는」 · 제한점 · 「결론적으로」 · 요약 한 문단) "
                 "③**낯선 서술어·연어**(저널·교과서 코퍼스에 0건인 용언 — AI 말투의 기계 대응물) "
                 "④참신성 문장 나열 + lit-scan 기록이 30일 안에 있는가. 전부 [의심] 후보, 연구 고유 표현은 ko_terms_ok.json"),
    # 2026-09-16 신설 — «배치»를 본다. 문장 검사(academic_prose·kprose·ko_venue)가 전부 통과시킨 결함 넷
    # (공백 문장이 개념 설명보다 먼저 · 정책 함의가 해석보다 먼저 · 「4절 참조」 · 결과가 방법 수식어를 되풀이)이
    # 전부 «순서»였다(한 원고, 연구자 「논리적 흐름과 학술적 글쓰기 순서도 감시해라」).
    dict(genres=["manuscript", "reading", "supplement", "brief"], name="flow",
         path=_tool("audit_flow.py"), auto=True,
         catches="**논리적 흐름 — 문장·문단의 «기능 순서»** — 서론(배경→정의→쓰임·한계→선행→공백→목적: 목적이 마지막인가, "
                 "공백 문단과 목적 문단이 인접한가, 정의 뒤에 쓰임·한계가 있는가) · 방법 소절 순서(자료원→대상→변수→분석→윤리) · "
                 "결과의 해석어 · 고찰(첫 문단=요약, 함의는 해석 뒤·제한점 앞, 결론이 마지막) · 「N절 참조」 상호참조 · "
                 "방법↔결과 8어절 중복 · 이웃 문단 첫머리 중복. 순서열을 찍어 주므로 `--exemplar DIR` 로 게재본과 나란히 본다. 전부 [의심]"),
    # 2026-09-16 신설 (한 원고) — 편집 라운드마다 표 행·결과 문장을 빼면서 고찰이 «결과에 없는 결과»를 말하게 될 수 있다.
    # 값 게이트(claim_*)는 값이 맞는지만 보고 «결과에 보고됐는지»는 안 본다. 연구자 「이런 부분도 audit 도구에 추가해봐」.
    # 소음 실측: JAMCH 게재본 37편(PDF 추출) 중 27편 0건, 나머지 1–9건.
    dict(genres=["manuscript", "reading", "supplement", "brief"], name="grounding",
         path=_tool("audit_grounding.py"), auto=True,
         catches="**고찰이 결과에 없는 결과를 말하는가** — 고찰·결론의 수치가 결과·표·그림·방법 어디에도 없다(인용 문장·연도·반올림 일치는 제외) · "
                 "「X를 제외한 분석에서」처럼 특정 분석을 가리키는데 그 핵심어가 결과 «본문»에 없다(표에만 있을 수 있음). 전부 [의심]"),
    # 2026-09-21 신설 (한 원고) — 연구자 인쇄본 피드백 넷이 전부 «표의 말»이었고(전체 주석에 별표·기준 시점 없음·
    # 각주 분량·영국식 철자) 구조·숫자·문장 검사는 그 판을 전부 통과시켰다. 고치기 전 판에서 16건, 고친 판에서 참고 2건.
    dict(genres=["manuscript", "reading", "supplement", "brief"], name="table_notes",
         path=_tool("audit_table_notes.py"), auto=True,
         catches="**표 제목·각주의 말** — 표 전체 주석에 기호(*†)를 달았다(기호가 표 안 어느 칸에도 없다·제목 끝 기호) · "
                 "제목·각주에 연도가 없어 «언제 자료인지» 모른다 · 각주에 안 풀린 약어 · 각주 없음/900자 초과 · "
                 "영문 전체의 영국식 철자(목표 저널이 영국식이면 도구에 --uk). 전부 [의심]/[참고]"),
    dict(genres=["manuscript", "reading", "supplement", "brief", "text", "abstract", "letter"], name="register",
         path=_tool("manuscript_register.py"),
         # ⚠ 여기에 SOURCE.md 를 적어 두었었다 -- 이 도구는 위치인자가 «파일 하나»뿐이라
         # 안내대로 치면 `unrecognized arguments: SOURCE.md` 로 죽는다(2026-08-27 실측).
         # 안내 문구는 «그대로 붙여 넣어 돌아가는 것»이어야 한다.
         # 🔴 2026-09-05: auto=False -> True. 근거였던 「--start/--end 가 필요하다」가 거짓이
         #   된 지 오래였다 -- 인자 없이 전문을 훑어도 정상 동작한다(실측). 그 사이 이 검사는
         #   [수동] 목록에 박혀 «한 번도 안 돌았고», 바로 이 검사가 담당하는 결함이 원고와
         #   초록에 실려 세 판을 통과했다. 안 도는 검사는 없는 검사다.
         #   구간을 좁히고 싶으면 사람이 --start/--end 로 다시 부른다.
         auto=True,
         catches="**참인데 방어적으로 읽히는 문장** — 1인칭 과정 서술·자기수정 일화 "
                 "('misled us once', 'the first version said the opposite', 'was not assumed but "
                 "checked'). 숫자 검증기는 전부 통과시킨다. 자기 오류를 많이 잡은 세션 직후일수록 "
                 "심하다. 대조구문·em-dash는 **개수만** 보고하고 실패로 치지 않는다 — 한계·사전등록·"
                 "정밀 대조는 학술 어투이므로 지우면 안 된다"),
]


def detect_genre(path):
    """(장르, 근거) 반환. 확실하지 않으면 그 사실을 근거 문자열로 알린다.

    확장자만으로 갈리지 않는 경우가 핵심이다:
      .pptx  캔버스 모양으로 슬라이드/포스터를 가른다 (포스터는 1장 + 큰/세로 캔버스)
      .docx  원고와 브리프는 파일만 봐선 확정할 수 없다 -> 경로 힌트로 추정하고 «추정»이라 밝힌다
    """
    ext = os.path.splitext(path)[1].lower()
    low = path.replace("\\", "/").lower()

    if ext == ".pptx":
        try:
            from pptx import Presentation
            prs = Presentation(path)
            w = prs.slide_width / 360000.0
            h = prs.slide_height / 360000.0
            n = len(prs.slides._sldIdLst)
            if n == 1 and (h > w or w > 60 or h > 60):
                return "poster", "1장 · 캔버스 %.0fx%.0fcm" % (w, h)
            # 학생에게 배포되는 수업 덱인가. ⛔ 「조교용·강사용」은 제외한다 — 같은 폴더에
            #   있어도 청중이 다르고, 거기서는 운영 문구가 오히려 있어야 한다.
            if (re.search(r"실습자료|수업|강의|교재|\d+주차", low)
                    and not re.search(r"조교용|강사용|내부", low)):
                return "class", "%d장 · 경로가 수업 자료(학생 배포)" % n
            return "slide", "%d장 · 캔버스 %.0fx%.0fcm" % (n, w, h)
        except Exception:
            return "slide", "캔버스를 읽지 못해 슬라이드로 가정"

    # 커버레터를 «가장 먼저» 본다 — 경로에 submission/manuscript 가 함께 있으면
    # 원고로 잡혀 버린다(2026-09-08 실측: Cover_letter.docx 가 manuscript 로 탐지됐다).
    if ext in (".md", ".txt", ".docx") and re.search(r"cover.?letter|커버레터|송부", low):
        return "letter", "파일·경로명의 'cover letter'"

    # 초록은 확장자가 아니라 «이름»으로 갈린다 — .md 로도 .docx 로도 온다. .docx 분기보다
    # 먼저 본다(경로에 submission/manuscript 가 함께 있으면 원고로 잡혀 버린다).
    if ext in (".md", ".txt", ".tex", ".docx") and re.search(r"abstract|초록", low):
        return "abstract", "파일·경로명의 'abstract/초록'"

    if ext == ".docx":
        ## ⭐ 학습 읽기가이드를 «수업자료보다도 먼저» 본다 (2026-09-30). 경로에 「수업」「강의」가
        ##    같이 올 수 있는데, 그러면 teaching(문항·모범답안 규율)으로 잡힌다 — 청중이 다르다.
        if re.search(r"읽기\s?가이드|reading.?guide|study.?guide|자습.?노트", low):
            return "study", "파일·경로명의 '읽기가이드' — 학습 정리본"
        ## ⭐ 수업자료를 «원고보다 먼저» 본다 (2026-09-07). 그 전에는 .docx 기본값인
        ##    manuscript 로 떨어져 목표 저널·xlsx 표 검사가 매번 [수동]으로 남았고,
        ##    audit.py 가 «항상» 「감사 미완료」를 찍었다 — 늘 뜨는 경고는 읽히지 않는다.
        ##    ⚠ 이 목록은 «넓히라고» 있다. 새 수업 프로젝트의 이름 규칙이 안 걸리면 여기 더한다.
        ##      확실히 하려면 호출부에서 --genre teaching 을 주면 된다.
        for kw in ("답안", "문제ver", "조교용", "문항", "모범답안", "실습",
                   "수업", "강의", "교과", "조교"):
            if kw in low:
                return "teaching", "파일·경로명의 '%s' 로 **추정** (틀리면 --genre)" % kw
        ## ⭐ totale 을 «원고보다 먼저» 본다 — 파일명이 `manuscript_totale_260909.docx`
        ##    라서 뒤에 두면 'manuscript' 에 걸려 제출본 규율로 검사되고, 설계상 있어야
        ##    하는 것(출처 경로·캡션 없는 표) 때문에 «항상» 실패 2가 난다.
        if "totale" in low or "reading" in low:
            return "reading", "파일·경로명의 'totale' — 읽기용 사본 (제출본은 --genre manuscript)"
        ## ⚠️ 보충을 «원고보다 먼저» 본다 -- 파일명이 대개 "..._supplement_..." 라
        ##    manuscript 키워드와 함께 오고, 뒤에 두면 원고로 잡혀 display_items 가 헛돈다.
        ## ⭐ 2026-09-08 — BMC/Springer 는 보충자료를 **«Additional file»** 이라 부른다.
        ##    그 어휘가 없어서 Additional_file_1.docx 가 경로의 'submission' 에 먼저 걸려
        ##    **원고로** 잡혔고, 그래서 supplement 장르가 이유를 적고 꺼 둔 display_items 가
        ##    돌아 보충표 전부를 «고아»로 신고했다(2026-09-08 실측, 실제 투고 패키지에서).
        ##    ⚠ 장르가 틀리면 «검사 목록 전체»가 틀린다 — 한 검사의 오탐보다 크다.
        for kw, g in (("supplement", "supplement"), ("보충", "supplement"),
                      ("supplementary", "supplement"),
                      ("additional file", "supplement"), ("additional_file", "supplement"),
                      ("additionalfile", "supplement"),
                      ("연구요약", "brief"), ("brief", "brief"), ("요약", "brief"),
                      ("manuscript", "manuscript"), ("원고", "manuscript"),
                      ("submission", "manuscript")):
            if kw in low:
                return g, "파일·경로명의 '%s' 로 **추정** (틀리면 --genre)" % kw
        return "manuscript", "**추정** — .docx 기본값. 브리프면 --genre brief"

    if ext == ".hwpx":
        return "report", "확장자"

    ## ⭐ .md/.txt 도 «이름»으로 가른다 (2026-09-20).
    ##    그 전에는 위 이름 라우팅이 `.docx` 안에만 있어서, .md 는 아래 확장자 루프에서
    ##    **outbound**(ext=[".md",".txt"])에 먼저 걸렸다. 실측: `manuscript/manuscript_draft_
    ##    260920.md` 가 「대외 문안」으로 탐지돼 검사 **하나만** 돌고 「통과 1 · 실패 0」이 찍혔다.
    ##    ⚠ 장르가 틀리면 한 검사의 오탐이 아니라 «검사 목록 전체»가 틀리고, 그 결과가
    ##      «초록불»이라 아무도 의심하지 않는다 — .docx 쪽에 이미 같은 사고가 두 번 적혀 있다.
    ##    ⚠ outbound 를 «가장 먼저» 본다. 대외 문안은 `docs/outbound/` 에 사는 것이 규약이고,
    ##      그 파일에도 「원고」 같은 낱말이 흔히 들어 있다.
    if ext in (".md", ".txt"):
        if re.search(r"(^|[\\/])outbound([\\/]|$)|발송|공문|신청서", low):
            return "outbound", "경로의 'outbound' — 대외 문안"
        for kw, g in (("supplement", "supplement"), ("보충", "supplement"),
                      ("supplementary", "supplement"),
                      ("연구요약", "brief"), ("brief", "brief"),
                      ("manuscript", "manuscript"), ("원고", "manuscript"),
                      ("submission", "manuscript"),
                      ("poster", "poster"), ("포스터", "poster")):
            if kw in low:
                return g, "파일·경로명의 '%s' 로 **추정** (틀리면 --genre)" % kw
        ## 표지가 없는 .md/.txt 는 «메모·원고 소스»다. 아래 확장자 루프에 맡기면 등록 순서 때문에
        ## outbound 가 먼저 걸리는데(둘 다 .md 를 선언한다), 그건 「대외로 나가는 문안」 규율이라
        ## `notes/RESULTS.md` 같은 내부 메모에 엉뚱한 검사를 돌리고 옳은 검사를 건너뛴다.
        return "text", "확장자 — 표지 없는 .md/.txt 는 메모·원고 소스로 본다"

    for g, meta in GENRES.items():
        if ext in meta["ext"]:
            return g, "확장자"
    return None, "등록된 장르 없음"


def applicable(genre):
    return [e for e in REGISTRY if genre in e["genres"]]


def build_args(entry, target, profile, tmpdir=None, extra=None):
    args = [sys.executable, entry["path"], target]
    if entry.get("profiled") and profile:
        args += ["--profile", profile]
    # deck_render_audit.py는 기본적으로 cwd에 _render/를 만든다 — 남의 저장소를 어지럽히지 않게.
    if entry["name"] == "render_overflow" and tmpdir:
        args += ["--png-dir", tmpdir]
    args += entry.get("args") or []     # 그 검사를 켜는 데 필요한 «고정» 인자 (호출자 입력이 아니다)
    args += extra or []
    return args


def self_check():
    missing = [(e["name"], e["path"]) for e in REGISTRY if not os.path.isfile(e["path"])]
    for name, p in missing:
        print("MISSING  %-16s %s" % (name, p))
    bad_genre = sorted({g for e in REGISTRY for g in e["genres"]} - set(GENRES))
    for g in bad_genre:
        print("UNKNOWN GENRE  %s" % g)
    stale = [(g, n) for g, meta in GENRES.items() for n in meta["off"]
             if not any(n == e["name"] and g in e["genres"] for e in REGISTRY)]
    for g, n_ in stale:
        print("STALE off[]    %s: '%s' 는 이 장르에 등록된 검사가 아니다" % (g, n_))
    empty = [g for g in GENRES if not applicable(g)]
    for g in empty:
        print("EMPTY GENRE    %s — 등록된 검사가 0개다 (REGISTRY의 genres에 빠졌다)" % g)
    stale += [(g, "<empty>") for g in empty]
    if missing or bad_genre or stale:
        return 1
    print("OK  도구 %d개 · 장르 %d개 · 등록 검사 %d개 전부 정합."
          % (len({e["path"] for e in REGISTRY}), len(GENRES), len(REGISTRY)))
    return 0


def show_list():
    for g, meta in GENRES.items():
        print("\n%-11s %s  (%s · profile=%s)"
              % (g, meta["label"], "/".join(meta["ext"]), meta["profile"]))
        for e in applicable(g):
            if e["name"] in meta["off"]:
                print("    OFF %-16s %s" % (e["name"], meta["off"][e["name"]]))
            else:
                mark = " " if e.get("auto") else "*"
                where = "skill" if SKILLS in e["path"] else "tools"
                print("  %s %-16s [%s] %s" % (mark, e["name"], where, e["catches"]))
    print("\n* = 인자가 필요해 자동 실행하지 않음 (명령줄을 안내한다)")
    print("OFF = 이 장르에서 의도적으로 끈 검사 (이유 명시)")
    return 0


def run_one(entry, target, profile, tmpdir=None, extra=None):
    args = build_args(entry, target, profile, tmpdir, extra)
    try:
        r = subprocess.run(args, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
    except Exception as exc:
        return "ERROR", str(exc)
    out = ((r.stdout or "") + (r.stderr or "")).strip()
    if r.returncode != 0:
        ## ⭐ 「그 형식을 못 읽는다」와 「결함을 찾았다」를 가른다 (2026-09-21).
        ##    docx 전용 검사(term_collision·table_widths·table_notes)에 .md 원고를 주면
        ##    python-docx 가 BadZipFile / "Package not found" 로 죽어 exit 1 이 난다. 그게
        ##    [FAIL] 로 찍히면 **모든 마크다운 원고에서 유령 실패가 상시로 뜨고**, 그러면
        ##    사람이 「실패:」 줄을 안 읽게 된다 — 늘 뜨는 경고는 경고가 아니다.
        ##    (실측 2026-09-21 한 초고: 매 실행 3건이 그것이었다.)
        ## ⚠ 좁게 잡는다 — 「파일을 못 열었다」의 서명만 보고 진짜 실패는 그대로 FAIL 로 둔다.
        low = out.lower()
        if any(k in low for k in ("badzipfile", "file is not a zip file",
                                  "package not found at", "지원하지 않는 형식")):
            return "SKIP-FORMAT", out
        return "FAIL", out
    # 🔴 PASS/FAIL 이분법이면 «[의심] 후보»가 집계에서 사라진다.
    # 여러 검사기가 문서에 「완전 자동판정이 아니라 [의심] 후보를 추려 사람이 보게 한다」고
    # 적어 두고 exit 0으로 끝난다(term_collision · display_items · academic_prose).
    # 그러면 여기서 [PASS]가 찍히고 본문이 안 보여, 「통과 7 · 실패 0」을 읽은 사람은
    # 볼 것이 남았다는 사실을 모른다. [[core]] §2⑦ — 통과한 검사는 그 검사가 «본 것»만 보증한다.
    # -> 도구의 exit code를 바꾸지 않고(다른 호출자를 깨지 않게) 출력에서 [의심]을 센다.
    # ⚠ 줄 «머리»의 마커만 센다. 본문 아무 데나 세면 도구의 안내 문구
    # (「[의심]은 후보다 — 원문을 열어…」)까지 후보로 세어 개수가 부풀었다(실측).
    n_susp = len(re.findall(r"(?m)^\s*\[의심\]", out))
    return ("주의" if n_susp else "PASS"), out


def audit(target, genre=None, profile=None, verbose=False, fast=False):
    if not os.path.isfile(target):
        print("없는 파일: %s" % target)
        return 1

    if genre:
        why = "--genre 로 지정"
    else:
        genre, why = detect_genre(target)
    if genre is None:
        print("\n%s\n  등록된 장르가 없다. GENRES에 추가할 것." % target)
        return 0

    meta = GENRES[genre]
    entries = applicable(genre)
    profile = profile or meta["profile"]

    ## ⛔ 2026-09-08 — 장르를 «새로 만들고» REGISTRY 에 등록하지 않으면 검사가 0개인 채
    ##    「통과 0 · 실패 0」이 찍힌다. 조용한 전면 통과는 검사가 없는 것보다 나쁘다 —
    ##    실측: letter 장르를 신설한 직후 커버레터가 그 상태로 「이상 없음」을 받았다.
    if not entries:
        print("\n" + "=" * 74)
        print(target)
        print("⛔ 장르 '%s' 에 등록된 검사가 «하나도» 없다 — REGISTRY 항목의 genres 에"
              % genre)
        print("   이 장르를 추가할 것. ⚠ 이 출력은 «통과»가 아니다.")
        return 1

    print("\n" + "=" * 74)
    print(target)
    print("장르: %s (%s)  ·  %s  ·  폰트 프로파일: %s"
          % (genre, meta["label"], why, profile))
    print("=" * 74)

    have_ppt = _has_powerpoint()
    tmpdir = tempfile.mkdtemp(prefix="audit_render_")
    failed, skipped, manual, off, attention = [], [], [], [], []

    for e in entries:
        if e["name"] in meta["off"]:
            off.append((e["name"], meta["off"][e["name"]]))
            continue
        if not os.path.isfile(e["path"]):
            skipped.append((e["name"], "도구 없음: %s" % e["path"]))
            continue
        if e.get("needs") == "powerpoint" and not have_ppt:
            skipped.append((e["name"], "PowerPoint COM 없음 (Windows + pywin32 필요)"))
            continue
        # --fast: COM/렌더가 필요한 검사를 뺀다. 훅에서 부를 때 쓴다 -- 저장할 때마다
        # 한글·PowerPoint를 띄우면 훅이 느려서 «끄게» 되고, 꺼진 가드는 없는 가드다.
        if fast and e.get("needs"):
            skipped.append((e["name"], "--fast: %s 필요 (수동으로 audit.py 재실행)" % e["needs"]))
            continue
        ## ⭐ 저장된 문체 코퍼스가 있으면 [수동] 을 «자동»으로 승격한다.
        ##    없으면 아래 기존 경로로 떨어져 「목표 저널이 필요해서 안 돌았다」가 인쇄된다.
        corpus_args = None
        if e.get("corpus_auto") and not e.get("auto"):
            corpus_args = _find_style_corpus(target)
        ## 국문 용어 검사 — 목표 저널 코퍼스가 프로젝트에 있으면 함께 넘긴다. 없어도 돈다
        ## (교과서 코퍼스만으로도 보이고, 코퍼스가 아예 없으면 도구가 「검사 못 했다」를 찍는다).
        if e.get("ko_corpus_auto"):
            corpus_args = _find_ko_corpus(target)
        if corpus_args is None and not e.get("auto"):
            cmd = "python %s %s %s" % (os.path.relpath(e["path"], CLAUDE), target, e.get("hint", ""))
            manual.append((e["name"], cmd.strip()))
            continue
        status, out = run_one(e, target, profile, tmpdir,
                              corpus_args or meta.get("extra", {}).get(e["name"]))
        if status == "SKIP-FORMAT":
            ## 이 검사는 «안 돌았다». 통과로도 실패로도 세지 않고 건너뜀에 이유와 함께 남긴다.
            skipped.append((e["name"], "이 파일 형식을 못 읽는다(.docx 전용) — "
                                       "원고를 빌드한 뒤 그 .docx 로 다시 돌릴 것"))
            continue
        if status == "주의":
            susp_lines = [ln for ln in out.splitlines()
                          if re.match(r"\s*\[의심\]", ln)]
            print("  [주의] %s   — 사람이 볼 후보 %d건" % (e["name"], len(susp_lines)))
            attention.append(e["name"])
            # 후보만 보여준다(전문을 쏟으면 읽지 않게 된다). 전문은 --verbose.
            for ln in (out.splitlines() if verbose else susp_lines):
                print("        " + ln)
            continue
        print("  [%s] %s" % (status, e["name"]))
        if status != "PASS":
            failed.append(e["name"])
        if out and (status != "PASS" or verbose):
            for ln in out.splitlines():
                print("        " + ln)

    for name, why_ in off:
        print("  [OFF]  %-15s %s" % (name, why_))
    for name, why_ in skipped:
        print("  [SKIP] %-15s %s" % (name, why_))
    for name, cmd in manual:
        print("  [수동] %-15s %s" % (name, cmd))

    ran = (len(entries) - len(failed) - len(attention)
           - len(skipped) - len(manual) - len(off))
    print("-" * 74)
    print("통과 %d · 주의 %d · 실패 %d · 장르상 제외 %d · 건너뜀 %d · 수동 %d"
          % (ran, len(attention), len(failed), len(off), len(skipped), len(manual)))
    if failed:
        print("실패: " + ", ".join(failed))
    if attention:
        # 실패가 아니므로 exit code는 0이다. 그러나 «남은 것이 있다»는 사실은 보여야 한다.
        print("주의(사람이 판정): " + ", ".join(attention))
    # 🔴 [수동]이 [PASS]처럼 읽힌다 — [주의]를 셋째 상태로 만든 것과 «같은 병»의 넷째 상태다.
    #   집계줄이 「통과 8 … 실패 0」으로 시작해 「수동 5」로 끝나므로 눈이 앞만 읽는다.
    #   2026-09-05 한 원고 실측: 원고 감사 결과를 SESSION_LOG 에 «"감사 통과 8·실패 0"»
    #   이라 적었고 «5개가 안 돌았다»는 안 적었다. 그 다섯 중 하나(register)가 바로 그날
    #   놓친 결함(「참인데 방어적으로 읽히는 문장」)을 담당하는 검사였다.
    #   -> exit code 는 건드리지 않는다(다른 호출자가 깨진다). 대신 «마지막 줄»을 뺏는다.
    #      집계 뒤에 이 문장이 오면 「감사 통과」로 요약하기 어렵다. [[core]] §2⑦.
    if manual or skipped:
        left = [n for n, _ in manual] + [n for n, _ in skipped]
        print("⚠ 감사 미완료 — 미실행 %d개: %s" % (len(left), ", ".join(left)))
        print("  이 상태를 「감사 통과」라고 보고하지 않는다. 위 [수동] 명령을 돌리거나,")
        print("  왜 안 돌렸는지를 결과와 «함께» 적는다.")
    shutil.rmtree(tmpdir, ignore_errors=True)
    return 1 if failed else 0


def main():
    ap = argparse.ArgumentParser(description="산출물 감사 단일 진입점 (장르 기준)")
    ap.add_argument("files", nargs="*")
    ap.add_argument("--genre", default=None, choices=sorted(GENRES),
                    help="장르를 직접 지정 (추정이 틀렸을 때)")
    ap.add_argument("--profile", default=None, help="폰트 프로파일 강제 (기본은 장르가 정함)")
    ap.add_argument("--list", action="store_true", help="장르 x 검사 지도")
    ap.add_argument("--self-check", action="store_true", help="등록 정합성 확인")
    ap.add_argument("-v", "--verbose", action="store_true", help="통과한 검사의 출력도 표시")
    ap.add_argument("--fast", action="store_true",
                    help="COM/렌더가 필요한 검사를 뺀다 (훅에서 부를 때). 뺀 것은 [SKIP]으로 이름이 남는다")
    a = ap.parse_args()

    if a.self_check:
        return self_check()
    if a.list:
        return show_list()
    if not a.files:
        ap.print_help()
        return 0

    rc = 0
    for f in a.files:
        rc |= audit(f, a.genre, a.profile, a.verbose, a.fast)
    return rc


if __name__ == "__main__":
    # Windows 콘솔은 cp949 라 한글·긴줄표(—)·기호 출력에서 UnicodeEncodeError 로 죽는다.
    # 결과를 다 만들어 놓고 «찍는 순간» 죽으므로, 부르는 쪽에는 도구가 고장난 것처럼 보인다.
    # ⚠ 모듈 최상단이 아니라 여기 두는 이유: 이 파일이 import 되기도 하면 최상단
    #    reconfigure 가 «호출자»의 인코딩을 바꾼다. 스크립트로 실행할 때만 돌게 한다.
    try:
        import sys as _s
        _s.stdout.reconfigure(encoding='utf-8', errors='replace')
        _s.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass
    sys.exit(main())
