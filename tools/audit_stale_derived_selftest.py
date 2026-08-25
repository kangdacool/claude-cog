"""audit_stale_derived.py 셀프테스트.

기준은 «이 도구가 있었으면 2026-08-19의 사고를 잡았는가»다. 그 사고를 합성 자료로 재현한다:
  - 입력: 날짜별 파일 1,092개
  - 파생물: 그중 22개만 있던 시점에 만들어져 그대로 멈춤
기대: STALE 판정 + 「나중에 생긴 입력 1,070개」.

    python audit_stale_derived_selftest.py
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TOOL = Path(__file__).with_name("audit_stale_derived.py")

T_EARLY = 1_700_000_000.0          # 파생물이 만들어진 시점
T_LATE = T_EARLY + 7 * 86400.0     # 입력이 계속 늘어난 시점


def touch(p: Path, mtime: float):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(b"x")
    os.utime(p, (mtime, mtime))


def run(*args, cwd):
    return subprocess.run([sys.executable, str(TOOL), *args],
                          cwd=cwd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def case_fossilised(tmp: Path):
    """실사고 재현 — 파생물이 22일치 시점에 얼어붙고 입력은 1,092일로 늘었다."""
    d0 = date(2022, 1, 1)
    for i in range(1092):
        day = d0 + timedelta(days=i)
        touch(tmp / "in" / f"{day:%Y}" / f"{day:%Y%m%d}.tif",
              T_EARLY if i < 22 else T_LATE)
    touch(tmp / "out" / "station_260812.rds", T_EARLY + 60)

    r = run("--derived", "out/station_*.rds", "--input", "in/**/*.tif",
            "--fail-on-stale", cwd=tmp)
    ok = ("STALE" in r.stdout and "1070" in r.stdout.replace(",", "") and r.returncode == 1)
    return ok, r.stdout


def case_fresh(tmp: Path):
    """파생물이 모든 입력보다 나중 — 낡지 않았다."""
    d0 = date(2022, 1, 1)
    for i in range(40):
        day = d0 + timedelta(days=i)
        touch(tmp / "in" / f"{day:%Y%m%d}.tif", T_EARLY)
    touch(tmp / "out" / "station_260819.rds", T_LATE)
    r = run("--derived", "out/station_*.rds", "--input", "in/*.tif",
            "--fail-on-stale", cwd=tmp)
    return ("낡은 쌍 없음" in r.stdout and r.returncode == 0), r.stdout


def case_missing(tmp: Path):
    """파생물이 아예 없을 때 조용히 통과하면 안 된다."""
    touch(tmp / "in" / "20220101.tif", T_EARLY)
    r = run("--derived", "out/nope_*.rds", "--input", "in/*.tif",
            "--fail-on-stale", cwd=tmp)
    return ("없음" in r.stdout and r.returncode == 1), r.stdout


def case_config(tmp: Path):
    """TSV 선언 모드 — 여러 쌍을 한 번에."""
    touch(tmp / "in" / "20220101.tif", T_LATE)
    touch(tmp / "out" / "a_260812.rds", T_EARLY)
    touch(tmp / "in2" / "20220101.tif", T_EARLY)
    touch(tmp / "out" / "b_260819.rds", T_LATE)
    cfg = tmp / "pairs.tsv"
    cfg.write_text("# name\tderived\tinput\n"
                   "낡은쌍\tout/a_*.rds\tin/*.tif\n"
                   "멀쩡한쌍\tout/b_*.rds\tin2/*.tif\n", encoding="utf-8")
    r = run("--config", str(cfg), cwd=tmp)
    return ("낡은쌍" in r.stdout and "STALE" in r.stdout
            and "1 / 2" in r.stdout), r.stdout


CASES = [
    ("화석화된 파생물 (실사고 재현)", case_fossilised),
    ("신선한 파생물", case_fresh),
    ("파생물 부재", case_missing),
    ("config 다중 쌍", case_config),
]


def main():
    passed = 0
    for name, fn in CASES:
        with tempfile.TemporaryDirectory() as td:
            ok, out = fn(Path(td))
        print(f"[{'PASS' if ok else 'FAIL'}] {name}")
        if not ok:
            print("--- 출력 ---")
            print(out)
        passed += ok
    print(f"\n{passed}/{len(CASES)} 통과")
    sys.exit(0 if passed == len(CASES) else 1)


if __name__ == "__main__":
    main()
