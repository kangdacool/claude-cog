"""파생물이 입력보다 낡았는지 검사한다 — 「화석화된 중간산출물」 잡기.

## 왜 필요한가 (2026-08-19 실사고)

한 파이프라인의 연구요약에 이렇게 적혀 있었다:

    "여름은 구름 때문에 위성 AOD 확보율이 낮아 표본에 들어오지 못했다."

자료의 성질처럼 읽히는 한계 서술이다. 사실이 아니었다. 격자 AOD는 **1,092일치가 디스크에
있었고**, 측정소 추출 단계(`aod_station_*.rds`)를 **22일치만 돌린 상태로 멈춰 있었을 뿐**이다.
그 추출 스크립트에는 날짜 필터조차 없었다 — 그냥 «먼저 돌린 날»의 입력량에 얼어붙은 것이다.
다시 돌리자 18일이 547일이 됐고, 그 위에 쌓아 올렸던 결론의 근거가 30배가 됐다.

**파생 단계는 만들어진 시점의 입력량에 화석화된다.** 그리고 그 화석은 「자료가 없다」처럼
보인다. mtime 한 번만 비교했으면 즉시 잡혔다.

## 무엇을 검사하나

1. **신선도** — 파생물보다 «나중에 생긴 입력»이 몇 개인가. 하나라도 있으면 파생물은 낡았다.
2. **커버리지** — 양쪽 파일명에서 날짜(YYYYMMDD/YYYY-MM-DD)를 뽑을 수 있으면, 입력에는
   있는데 파생물이 만들어진 뒤 추가된 날이 몇 개인지 센다.

파일명·mtime만 본다. 파생물 «내용»의 날짜는 형식이 제각각이라(rds/parquet/RData) 열지 않는다 —
그래서 이 도구는 «반드시 낡았다»는 말은 해도 «신선하다»는 보증은 못 한다. 통과가 곧 최신은
아니라는 뜻이고, 그건 도구의 한계로 명시한다.

## 사용

    # 한 쌍 직접
    python audit_stale_derived.py --derived "data/processed/aod_station_*.rds" \
                                  --input   "data/processed/aod/**/*.tif"

    # 프로젝트가 선언한 여러 쌍 (TSV: name<TAB>derived_glob<TAB>input_glob, # 은 주석)
    python audit_stale_derived.py --config .stale_pairs.tsv

    # 파이프라인 게이트로: 낡은 것이 있으면 exit 1
    python audit_stale_derived.py --config .stale_pairs.tsv --fail-on-stale
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from glob import glob
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 파일명 안의 날짜. 긴 것부터 시도해야 2024-03-17 을 2024 로 잘못 읽지 않는다.
DATE_PATTERNS = [
    (re.compile(r"(20\d{2})-(\d{2})-(\d{2})"), "%Y-%m-%d"),
    (re.compile(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)"), "%Y%m%d"),
]


def file_date(path: Path):
    """파일명에서 날짜를 뽑는다. 스탬프(YYMMDD)는 «내용 날짜»가 아니므로 보지 않는다."""
    name = path.name
    for rx, fmt in DATE_PATTERNS:
        m = rx.search(name)
        if m:
            try:
                return datetime.strptime("".join(m.groups()), fmt.replace("-", "")).date()
            except ValueError:
                continue
    return None


def expand(pattern: str, root: Path):
    hits = [Path(p) for p in glob(str(root / pattern), recursive=True)]
    return sorted(p for p in hits if p.is_file())


def check_pair(name: str, derived_glob: str, input_glob: str, root: Path):
    derived = expand(derived_glob, root)
    inputs = expand(input_glob, root)

    r = {"name": name, "derived_glob": derived_glob, "input_glob": input_glob,
         "n_derived": len(derived), "n_input": len(inputs), "verdict": "OK", "notes": []}

    if not derived:
        r["verdict"] = "NO_DERIVED"
        r["notes"].append("파생물이 없다 — 아직 안 돌렸거나 경로가 틀렸다")
        return r
    if not inputs:
        r["verdict"] = "NO_INPUT"
        r["notes"].append("입력이 없다 — 경로가 틀렸을 가능성이 높다")
        return r

    # 파생물이 여럿이면 가장 최근 것이 정본이라고 본다(날짜 스탬프 관행)
    d_latest = max(derived, key=lambda p: p.stat().st_mtime)
    d_mtime = d_latest.stat().st_mtime
    r["derived_file"] = d_latest.name
    r["derived_mtime"] = datetime.fromtimestamp(d_mtime).strftime("%Y-%m-%d %H:%M")

    newer = [p for p in inputs if p.stat().st_mtime > d_mtime]
    r["n_input_newer"] = len(newer)
    if newer:
        r["verdict"] = "STALE"
        newest = max(newer, key=lambda p: p.stat().st_mtime)
        r["notes"].append(
            f"입력 {len(newer)}/{len(inputs)}개가 파생물보다 나중에 생겼다 "
            f"(가장 최근 입력 {newest.name} "
            f"{datetime.fromtimestamp(newest.stat().st_mtime):%Y-%m-%d %H:%M})")

    # 커버리지 — 양쪽에서 날짜를 뽑을 수 있을 때만
    in_dates = {d for d in (file_date(p) for p in inputs) if d}
    if in_dates:
        r["n_input_dates"] = len(in_dates)
        missed = {d for d in in_dates
                  if datetime.combine(d, datetime.min.time()).timestamp() > d_mtime}
        # 파생물 생성 시점 이후 «내용 날짜»가 있는 입력은 애초에 담길 수 없었다
        added = [p for p in inputs if p.stat().st_mtime > d_mtime and file_date(p)]
        if added:
            r["n_dates_added_since"] = len({file_date(p) for p in added})
            r["notes"].append(
                f"파생물 생성 뒤 추가된 입력 날짜 {r['n_dates_added_since']}개 — "
                f"그 날들은 파생물에 «있을 수 없다»")
        del missed

    return r


def render(results, fail_on_stale: bool):
    bad = [r for r in results if r["verdict"] in ("STALE", "NO_DERIVED", "NO_INPUT")]

    print("=" * 78)
    print("파생물 신선도 검사")
    print("=" * 78)
    for r in results:
        mark = {"OK": "  ok  ", "STALE": " STALE", "NO_DERIVED": " 없음 ",
                "NO_INPUT": " 입력? "}[r["verdict"]]
        print(f"\n[{mark}] {r['name']}")
        print(f"         파생물 {r['n_derived']}개  입력 {r['n_input']}개")
        if "derived_file" in r:
            print(f"         최신 파생물: {r['derived_file']}  ({r['derived_mtime']})")
        for n in r["notes"]:
            print(f"         → {n}")

    print()
    print("-" * 78)
    if bad:
        print(f"낡았거나 확인이 필요한 쌍: {len(bad)} / {len(results)}")
        print()
        print("낡은 파생물은 «자료가 없다»처럼 보인다. 한계 문단에 부재를 쓰기 전에")
        print("그 부재가 자료의 성질인지 이 화석 때문인지 확인할 것.")
    else:
        print(f"낡은 쌍 없음 ({len(results)} 쌍 검사)")
    print()
    print("※ 파일명과 mtime만 본다. 파생물 내용은 열지 않으므로 «최신»을 보증하지 못한다 —")
    print("  이 검사가 통과해도 파생 스크립트가 입력 일부만 읽었을 가능성은 남는다.")

    return 1 if (bad and fail_on_stale) else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--derived", help="파생물 glob (루트 기준 상대경로)")
    ap.add_argument("--input", help="입력 glob (루트 기준 상대경로)")
    ap.add_argument("--config", help="TSV: name<TAB>derived_glob<TAB>input_glob")
    ap.add_argument("--root", default=".", help="기준 디렉토리 (기본: 현재)")
    ap.add_argument("--fail-on-stale", action="store_true",
                    help="낡은 쌍이 있으면 exit 1 — 파이프라인 게이트용")
    a = ap.parse_args()

    root = Path(a.root).resolve()
    pairs = []
    if a.config:
        for line in Path(a.config).read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [c.strip() for c in line.split("\t")]
            if len(parts) < 3:
                sys.exit(f"config 형식 오류 (name<TAB>derived<TAB>input): {line}")
            # 헤더 행은 건너뛴다 — 없으면 「파생물 0개」 쌍으로 잡혀 오탐이 된다
            if [c.lower() for c in parts[:3]] == ["name", "derived", "input"]:
                continue
            pairs.append((parts[0], parts[1], parts[2]))
    elif a.derived and a.input:
        pairs.append(("(직접 지정)", a.derived, a.input))
    else:
        ap.error("--config 또는 --derived/--input 을 줄 것")

    results = [check_pair(n, d, i, root) for n, d, i in pairs]
    sys.exit(render(results, a.fail_on_stale))


if __name__ == "__main__":
    main()
