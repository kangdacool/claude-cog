#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""precommit_scan 자기시험. 절반은 «울리지 않는 것»을 증명한다.

    python agent/tools/precommit_scan.py --selftest   (또는 이 파일을 직접)

왜 있나
    2026-08-25: 환자 성명이 프로젝트 메모리에 6일 동안 있었다. 게이트는 «돌고 있었다» --
    core.hooksPath 로 매 커밋마다 돌았고, 그냥 통과시켰다. 검출기가 「이름은 문서 파일명
    자리에 온다」를 전제했는데 실제로 샌 것은 따옴표 속 «검증 샘플 이름»이었다.
    통과는 증거가 아니다. 그날의 «문장 «형태»»를 회귀 사례로 박아 둔다.

    ⚠️ 이름은 반드시 «가짜»를 쓴다(가나다). 실명을 시험 자료로 쓰면 지우려던 그 이름을
       코드에 다시 심고, 그 파일이 공개 저장소로 나간다 -- 2026-08-25 에 실제로 그랬다.
       회귀의 값어치는 «형태»에 있지 특정 이름에 있지 않다.
"""
import re
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
import precommit_scan as PS                                    # noqa: E402


def names_in(line):
    out = set()
    for rx in PS.NAME_CONTEXT:
        for m in rx.finditer(line):
            g = m.group(1)
            if g not in PS.NOT_NAME:
                out.add(g)
    return out


# (문장, 잡아야 하나)
CASES = [
    # ── 울려야 한다 ──────────────────────────────────────────────────────
    ('샘플("가나다")이 하필 Template A(파일명 "이름 생년월일 기록지_보안.pdf"', True),
    ("2022 초진기록지-가나다.pdf 를 파싱", True),
    ("가나다 561120 기록지_보안.pdf", True),
    ('환자 "김유신" 의 CIST 점수가', True),
    # 코드의 열 이름 목록 -- 2026-08-25 실측 오탐(psy 저장소에서 9건)
    ('["범주", "환자ID", "파일", "문서"]', False),
    ('"환자판정": 15, "문서": 16', False),
    ('if k in ("조건", "환자", "노출확보"):', False),
    # ⚠️ 괄호형은 «일부러» 안 잡는다 -- 한국어 산문의 `개념(설명)` 과 구별할 수 없어
    #    「환자에서 기저 GdepS(우울)」의 «우울»까지 이름으로 잡았다(실측 오탐).
    #    이름 하나 더 잡으려다 스캐너를 못 쓰게 만드는 쪽이 손해다.
    ("대상자(이순신)는 제외", False),
    ("경도인지장애 환자에서 기저 GdepS(우울)와 인지기능", False),
    ('"강감찬"님이 재방문', True),
    # ── 울리면 «안 된다» ─────────────────────────────────────────────────
    ("샘플(환자 1명)이 하필 Template A 였다", False),          # 고친 뒤의 그 줄
    ('파일명 "이름 생년월일 기록지_보안.pdf" 형식', False),     # 서식 이름
    ("초진기록지-홍길동.pdf 는 예시 파일명이다", False),        # 예시 이름
    ('환자 "표본" 추출 방법', False),                          # NOT_NAME
    ('"정본이" 갈라지지 않게 한다', False),                    # 낱말
    ("샘플 크기가 「충분한지」 확인", False),                   # 낫표 인용
    ('결과를 "재현성" 관점에서 본다', False),                  # 따옴표 속 낱말
    ("피험자 수가 602명이다", False),                          # 이름 없음
    ('def f(x): return "결과값"', False),                      # 코드
    # 2026-09-12 실측 오탐 -- 맥락 낱말과 따옴표 «사이에 낱말이 끼면» 이름표가 아니라 문장이다.
    # 이 한 줄(치과 기사 제목)이 stock-ticker 저장소의 일일 커밋을 사흘 동안 막았다.
    ("환자 중심 ‘투명교정’ 증례 종합적 고찰 - 치과신문", False),
    ("대상자 모집 '중간보고' 자료", False),                     # 같은 형태, 다른 맥락 낱말
    # 붙어 있는 형태는 계속 울려야 한다(위 좁히기가 진짜 유출까지 죽이지 않았는지 확인)
    ('대상자 = "김유신"', True),
    ("성명: '이순신'", True),
]

fails = []
for line, should in CASES:
    got = names_in(line)
    fired = bool(got)
    ok = fired == should
    mark = "OK  " if ok else "실패"
    what = f" -> {sorted(got)}" if got else ""
    print(f"  [{mark}] {'울림' if fired else '조용':4s} (기대 {'울림' if should else '조용'})"
          f"  {line[:52]}{what}")
    if not ok:
        fails.append(line[:40])

print()
if fails:
    print(f"★ {len(fails)}건 실패")
    for f in fails:
        print("   " + f)
    sys.exit(1)
print(f"{len(CASES)}/{len(CASES)} 통과  («울리면 안 되는» 경우 "
      f"{sum(1 for _, s in CASES if not s)}건 포함)")
