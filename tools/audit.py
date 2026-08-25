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
    "poster":     dict(ext=[".pptx"], profile="poster", label="학술대회 포스터",
                       off={"render_overflow":
                            "하단 푸터 띠가 설계상 존재 — 슬라이드용 «아래 가장자리» 휴리스틱이 항상 FAIL"}),
    "manuscript": dict(ext=[".docx"], profile="doc", label="저널 투고 원고", off={}),
    "brief":      dict(ext=[".docx"], profile="doc", label="국문 브리프·연구요약",
                       off={"verify_tables":
                            "저널 표 검사다. 브리프는 격자·zebra가 의도된 디자인이라 규칙이 다르다"}),
    "report":     dict(ext=[".hwpx"], profile="doc", label="정부·기관 보고서", off={}),
    "text":       dict(ext=[".md", ".txt", ".tex", ".pdf"], profile="doc", label="메모·원고 소스", off={}),
    # 산출물이 아니라 «들어오는 자료원». 다른 장르가 "내가 만든 걸 감사"한다면 이건
    # "남이 준 걸 쓰기 전에 감사"다 -- 2026-08-18 사고로 신설: 열 이름은 정확히 맞는
    # 과거력 체크리스트 31열을 «제외기준의 구조화된 소스»로 반겼는데, 전 6,260행이
    # 미입력(전부 '0')인 죽은 템플릿 열이었다. 그대로 실었으면 "전원 음성"이라는 틀린
    # 데이터가 됐다. 헤더 존재 ≠ 값 존재.
    "data":       dict(ext=[".xlsx", ".csv"], profile="doc", label="입력 자료원(엑셀·CSV)", off={}),
}

# ---------------------------------------------------------------------------
# REGISTRY — 장르별로 «무엇을 돌리는가». 이 표가 정본이다.
# ---------------------------------------------------------------------------
REGISTRY = [
    dict(genres=["slide", "meeting", "poster"], name="text_fit",
         path=_skill("pptx-editing", "audit_text_fit.py"), auto=True, needs="powerpoint",
         catches="텍스트가 슬라이드 밖으로 나가거나 다른 도형 위에 얹힘 (PowerPoint 실측). "
                 "표 셀은 못 본다 -> render_overflow가 담당"),
    dict(genres=["slide", "meeting", "poster"], name="surface_text",
         path=_skill("pptx-editing", "audit_surface_text.py"), auto=True,
         catches="표면에 남은 편집 해명·내비게이션·재진술·내부 파일명, Figure/Table 번호 drift. "
                 "회의 덱에선 provenance가 오탐일 수 있다(출처 표기가 바람직한 register)"),
    dict(genres=["slide", "meeting", "poster"], name="xml_integrity", path=_tool("deck_audit.py"), auto=True,
         catches="구역·줌 링크·creationId 중복·목차↔구분장 제목 불일치·고아 media"),
    dict(genres=["slide", "meeting", "poster"], name="render_overflow", path=_tool("deck_render_audit.py"),
         auto=True, needs="powerpoint",
         catches="**표 행높이 확장으로 인한 잘림.** 좌표·XML 검사가 전부 통과시키는 유형"),
    dict(genres=["slide", "meeting", "poster", "manuscript", "brief"], name="font_sizes",
         path=_tool("audit_font_sizes.py"), auto=True, profiled=True,
         catches="글자 크기 바닥값 미만 [RUN] + 그림 축소배치로 구워진 글씨가 줄어든 것 [SCALE]"),
    # .pptx/.docx만 지원(.hwpx 아직 없음) -- report 장르는 제외.
    dict(genres=["slide", "meeting", "poster", "manuscript", "brief"], name="term_collision",
         path=_tool("audit_term_collision.py"), auto=True,
         catches="각주(†)로 정의한 용어가 다른 곳에서 다른 뜻으로 재사용됨. 완전 자동판정이 "
                 "아니라 [의심] 후보를 추려 사람이 보게 한다 -- 놓치는 것보다 과잉표시가 낫다"),
    # python-docx로 열기 때문에 .pptx를 주면 ValueError로 죽는다 (2026-08-10 확인).
    dict(genres=["manuscript", "brief"], name="table_widths", path=_tool("audit_table_widths.py"),
         auto=True, catches="렌더 전에 «숫자가 두 줄로 쪼개질» 표 열을 예측"),
    dict(genres=["manuscript", "brief", "report", "text"], name="text_consistency",
         path=_tool("audit_text_consistency.py"), auto=True,
         catches="괄호 미종결·참고문헌 결번/미인용·단위/대시/p값/콤마 표기 혼용"),
    # text_consistency와 축이 다르다: 저쪽은 «참고문헌» 번호, 이쪽은 «표·그림» 번호다.
    # 2026-08-20 한 원고에서: 보충 표를 더하며 손으로 순서를 확인했더니 원래 있던 Table S1/S2가
    # 첫 인용 순서와 반대였고, 고친 판을 이 검사로 돌리니 Figure 1/2도 같은 상태였다 --
    # 사람이 물어봐 줘야만 드러나던 축이다.
    dict(genres=["manuscript", "brief"], name="display_items",
         path=_tool("audit_display_items.py"), auto=True,
         catches="표·그림 번호가 «본문 첫 인용 순서»와 어긋남 · 인용되지 않은 고아 표시물 · "
                 "표시물이 없는 유령 인용 · 결번. 각주가 그 문서의 형제들보다 튀는 것은 [의심]"),
    # auto=False로 «등록만» 해 둔다 — 목표 저널을 알아야 돌기 때문이다. 그게 요점이다:
    # 이렇게 두면 원고를 감사할 때마다 "목표 저널이 필요해서 안 돌았다"가 «인쇄»된다.
    # 2026-08-20 한 원고에서: 목표 저널이 미정인 채로 원고를 계속 고쳤고, 처음 재봤을 때
    # 세미콜론 145/10k(그 저널 27편의 «최대»가 80)에 1인칭 we는 Q1 미만이었다.
    # 안 정한 것이 조용하면 안 맞춘 것도 조용하다.
    dict(genres=["manuscript"], name="journal_fit", path=_tool("journal_fit.py"), auto=False,
         hint='--journal "Lancet Healthy Longev"   (프로젝트 CLAUDE.md에 목표 저널을 박아둘 것)',
         catches="**목표 저널 실물 대비 «형태»와 «문체»** — 섹션별 분량 배분, 문장 길이, "
                 "1인칭 비율, 세미콜론·대시 빈도. 분량 제한이 없어도 돌린다(압축 도구가 "
                 "아니라 «그 저널 수준인가»의 측정기다). 용어는 esearch 빈도로 따로 확인"),
    # 암호 걸린 워크북은 --password가 필요해 auto=False. 암호가 없으면 그냥 돌아간다.
    dict(genres=["data"], name="column_profile", path=_tool("profile_columns.py"), auto=True,
         hint="--only-suspect  (암호 걸린 xlsx면 --password ... 추가)",
         catches="**죽은 열·상수 열·희소 열.** 헤더는 멀쩡한데 값이 없거나 전부 같은 값인 "
                 "컬럼 -- 그대로 분석에 실으면 «전원 음성» 같은 틀린 데이터가 된다. "
                 "새 자료원을 파이프라인에 붙이기 «전에» 돌릴 것"),
    # brief에도 등록해 두고 GENRES["brief"].off에서 «이유와 함께» 끈다. 목록에서 빼버리면
    # "브리프엔 왜 이 검사가 없나"에 답이 남지 않는다 — 보이는 제외가 조용한 누락보다 낫다.
    dict(genres=["manuscript", "brief"], name="verify_tables", path=_tool("verify_tables.py"), auto=False,
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
    dict(genres=["text", "manuscript", "brief", "report"], name="surface_leak",
         path=_tool("surface_leak_scan.py"),
         auto=False, hint='--terms "이 메모는" "사용자가" "이전 버전"',
         catches="register별 표면 누출. **기본 목록이 없는 것이 설계** — 문서마다 안전한 표현이 다르다"),
    # ⭐ AI 티는 «모든 산출물»에 해당한다 — 덱만의 문제가 아니다. 덱·포스터는 surface_text(auto)로
    # 이미 덮여 있었는데 원고·보고서 쪽은 surface_leak이 auto=False라 사실상 안 돌았다. 그 auto=False의
    # 이유는 "호출자가 terms를 줘야 한다"였고, --rhetoric은 아무것도 안 받으므로 그 이유가 안 걸린다.
    # 그래서 같은 스크립트를 «인자 고정 + auto»인 별도 항목으로 한 번 더 등록한다.
    # 사정거리 근거는 surface_leak_scan.rhetoric_rules() 실측(원고 429줄, 오탐 0) 참조.
    dict(genres=["manuscript", "brief", "text", "report"], name="rhetoric",
         path=_tool("surface_leak_scan.py"), auto=True, args=["--rhetoric"],
         catches="**AI가 쓴 티** — 카드 제목이 «수사적 기능»(The hook·Key point), 화살표·≠를 말 대신, "
                 "문장 속 대문자 강조. 내용이 아니라 «형태»라서 사람이 읽어선 잘 안 걸린다"),
    # 고백형 어투. 원본 .md/.tex를 본다(조판된 .docx가 아니라) — 고칠 곳이 원고이기 때문.
    # auto=False 인 이유는 대상이 빌드 산출물이 아니라 소스 원고라서, 경로를 호출자가 준다.
    # ⭐ 국문 학술 산문 문체. rhetoric/register가 «영문 AI 티»를 보는 자리에 대응하는 국문판이
    # 비어 있었다 — brief 장르는 검사가 둘뿐이었다. 규칙은 지도교수가 «실제로 지운» 문장에서만
    # 뽑았고(teacher-ohs 2026-08-22), 넓히기 전에 `--selftest`로 known-good을 안 잡는지 본다.
    # ⚠ 감수자 수정본을 known-good으로 쓰지 않는다 — 「심하게 느껴지는 것만」 고치신 최소 수정본이다.
    dict(genres=["manuscript", "brief", "report", "text"], name="academic_prose",
         path=_tool("audit_academic_prose.py"), auto=True,
         catches="**국문 학술 산문 문체** — ①예고·메타 문장(뒤에 나올 내용을 미리 알림) "
                 "②방법 절의 «이유» 서술(근거는 남기되 「~때문이다」를 사실 진술로) "
                 "③저널리즘 표현(사실이 아니라 분위기) ④Key findings 수치의 본문 중복. "
                 "[FAIL]은 근거가 확실한 것만, 나머지는 [의심]으로 사람에게 넘긴다"),
    dict(genres=["manuscript", "brief", "text"], name="register", path=_tool("manuscript_register.py"),
         auto=False, hint='SOURCE.md --start "## Introduction" --end "## Display items"',
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
            return "slide", "%d장 · 캔버스 %.0fx%.0fcm" % (n, w, h)
        except Exception:
            return "slide", "캔버스를 읽지 못해 슬라이드로 가정"

    if ext == ".docx":
        for kw, g in (("연구요약", "brief"), ("brief", "brief"), ("요약", "brief"),
                      ("manuscript", "manuscript"), ("원고", "manuscript"),
                      ("submission", "manuscript")):
            if kw in low:
                return g, "파일·경로명의 '%s' 로 **추정** (틀리면 --genre)" % kw
        return "manuscript", "**추정** — .docx 기본값. 브리프면 --genre brief"

    if ext == ".hwpx":
        return "report", "확장자"

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
        if not e.get("auto"):
            cmd = "python %s %s %s" % (os.path.relpath(e["path"], CLAUDE), target, e.get("hint", ""))
            manual.append((e["name"], cmd.strip()))
            continue
        status, out = run_one(e, target, profile, tmpdir,
                              meta.get("extra", {}).get(e["name"]))
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
    sys.exit(main())
