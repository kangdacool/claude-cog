#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""precommit_scan 자기시험. 절반은 «울리지 않는 것»을 증명한다.

    python agent/tools/precommit_scan.py --selftest   (또는 이 파일을 직접)

왜 있나
    2026-08-25: 환자 성명이 프로젝트 메모리에 6일 동안 있었다. 게이트는 «돌고 있었다» --
    core.hooksPath 로 매 커밋마다 돌았고, 그냥 통과시켰다. 검출기가 「이름은 문서 파일명
    자리에 온다」를 전제했는데 실제로 샌 것은 따옴표 속 «검증 샘플 이름»이었다.
    통과는 증거가 아니다. 그날의 «진짜 문장»을 회귀 사례로 박아 둔다.
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
    ('샘플("환자 1명")이 하필 Template A(파일명 "이름 생년월일 기록지_보안.pdf"', True),
    ("2022 초진기록지-환자 1명.pdf 를 파싱", True),
    ("환자 1명 561120 기록지_보안.pdf", True),
    ('환자 "김유신" 의 CIST 점수가', True),
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
