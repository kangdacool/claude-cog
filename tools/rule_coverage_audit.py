# -*- coding: utf-8 -*-
"""파생 판정 규칙이 «원문 신호»를 놓치는지 검사한다 — 규칙의 불완전성 탐지.

왜 필요한가 (2026-08-19 실사고)
--------------------------------------------------
같은 부류의 결함이 하루에 두 번 나왔다.
  lives_alone              값 있는 1,631건이 전부 'No' — 「독거」라고 적힌 537건을 놓침
  assistive_device_detail  전부 빈칸 — 기구 이름으로 Yes를 판정해 놓고 그 이름을 버림
둘 다 «원문에는 정보가 있는데 우리 규칙이 못 잡은» 경우다. 파서는 조용히 통과하고, 값도
«정상적인 음성»처럼 보이며, 스키마 검증도 전부 지나간다.

`profile_columns.py`의 CONSTANT/SKEWED는 결과가 «한쪽으로 몰렸다»를 잡는다. 이 도구는 한 걸음
더 가서 **왜 몰렸는지** — 원문에 반대 신호가 있는데도 몰린 것인지 — 를 직접 센다. 규칙이
«틀린» 게 아니라 «불완전»한 것을 잡는 유일한 방법이다.

무엇을 하는가
-------------
규칙표(CSV)의 각 줄마다: 원문 열에 신호 정규식이 걸리는데 판정 열이 기대값이 아닌 레코드를 센다.
0이 아니면 규칙이 그만큼 놓치고 있다는 뜻이다.

규칙표 형식 (UTF-8 CSV, 헤더 필수)
  field,raw_field,pattern,expected,note
  lives_alone,living_arrangement_raw,독거|혼자\\s*거주,Yes,동거 칸에 직접 적는다
  suicide_attempt_hx,misc_history_raw_text,자살,Yes,

  field      판정 열
  raw_field  원문 열 (신호를 찾을 곳)
  pattern    파이썬 정규식
  expected   신호가 있으면 field가 가져야 할 값
  note       사람용 메모(검사에 안 쓰임)

사용
----
  python rule_coverage_audit.py RECORDS.json RULES.csv
  python rule_coverage_audit.py RECORDS.json RULES.csv --show 5

RECORDS.json은 «레코드 객체의 배열»이면 무엇이든 된다(파이프라인 중간 산출물).

종료코드: 놓친 것이 하나라도 있으면 1.

⚠️ 새 판정 규칙을 만들 때마다 규칙표에 한 줄 추가하라. 한 줄이 곧 회귀 검사다.
"""
import argparse
import csv
import json
import re
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def load_rules(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    need = {"field", "raw_field", "pattern", "expected"}
    for i, r in enumerate(rows, 2):
        missing = need - {k for k, v in r.items() if v not in (None, "")}
        if missing:
            raise SystemExit(f"[ABORT] {path}:{i} 필수 열 누락: {', '.join(sorted(missing))}")
    return rows


def audit(records, rules, show=3):
    total_missed = 0
    print(f"{'판정 열':<26}{'원문 신호':>9}{'판정 양성':>10}{'놓친 것':>9}")
    print("-" * 56)
    for rule in rules:
        field, raw_field = rule["field"], rule["raw_field"]
        expected = rule["expected"]
        try:
            pat = re.compile(rule["pattern"])
        except re.error as e:
            print(f"★ {field:<24} 정규식 오류: {e}")
            total_missed += 1
            continue

        n_signal = n_pos = 0
        missed = []
        for r in records:
            raw = r.get(raw_field)
            has_signal = bool(raw) and bool(pat.search(str(raw)))
            n_signal += has_signal
            n_pos += str(r.get(field)) == expected
            # 판정이 아예 없는(None) 레코드는 «놓쳤다»고 하지 않는다 — 미평가와 오판은 다르다.
            if has_signal and r.get(field) is not None and str(r.get(field)) != expected:
                missed.append(r)

        mark = "  " if not missed else "★ "
        print(f"{mark}{field:<24}{n_signal:>9,}{n_pos:>10,}{len(missed):>9,}")
        if missed:
            total_missed += len(missed)
            for r in missed[:show]:
                ident = r.get("source_file") or r.get("id") or r.get("canonical_id") or "?"
                print(f"      {str(ident)[:40]:42s} {field}={r.get(field)!r}  "
                      f"{raw_field}={str(r.get(raw_field))[:40]!r}")
    return total_missed


def main():
    ap = argparse.ArgumentParser(description="파생 판정 규칙의 원문 신호 누락 검사")
    ap.add_argument("records", help="레코드 배열 JSON")
    ap.add_argument("rules", help="규칙표 CSV")
    ap.add_argument("--show", type=int, default=3, help="놓친 사례를 몇 건까지 보일지")
    # ⚠️ 허용치는 «봐줄 만큼»이 아니라 «이미 사람이 열어 보고 판정한 잔여분»이다.
    # 자연어 판정에는 사람도 갈리는 사례가 남는다(실측 2026-08-20 흡연 규칙:
    # 「음주시에만 흡연」·「군대 가서 배웠다가 약 2년」). 그것 때문에 검사가 영구히
    # 빨간불이면 아무도 안 본다 -- 통과 불가능한 검사는 신호가 0이다(precommit_scan 교훈).
    # 호출부가 «세어 보고 사유를 적은 수»를 넣고, 그보다 늘면 실패한다. 즉 «신규»만 잡는다.
    ap.add_argument("--max-missed", type=int, default=0,
                    help="이미 확인한 잔여 누락 건수. 이보다 많아지면 실패한다")
    a = ap.parse_args()

    for p in (a.records, a.rules):
        if not Path(p).exists():
            print(f"[ABORT] 파일 없음: {p}")
            return 2

    records = json.loads(Path(a.records).read_text(encoding="utf-8"))
    if not isinstance(records, list):
        print("[ABORT] RECORDS.json은 레코드 객체의 배열이어야 합니다")
        return 2
    rules = load_rules(a.rules)
    print(f"레코드 {len(records):,}건 · 규칙 {len(rules)}개\n")

    missed = audit(records, rules, a.show)
    print()
    if missed > a.max_missed:
        print(f"⚠️ 규칙이 놓친 레코드 {missed:,}건 — 원문에 신호가 있는데 판정이 다릅니다."
              + (f" (허용 {a.max_missed}건)" if a.max_missed else ""))
        return 1
    if missed:
        print(f"놓친 레코드 {missed:,}건 — 호출부가 확인한 잔여분({a.max_missed}건) 이내입니다.")
        print("  ⚠️ 이 숫자가 늘면 실패한다. 늘었다면 «새로 생긴 누락»이므로 원문을 볼 것.")
        return 0
    print("모든 규칙이 원문 신호를 빠짐없이 반영합니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
