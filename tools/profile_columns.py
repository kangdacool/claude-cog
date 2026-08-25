# -*- coding: utf-8 -*-
"""새 자료원의 열을 «쓰기 전에» 값 분포를 센다 — 죽은 열·상수 열·거의 빈 열 적발.

왜 필요한가 (2026-08-18 실사고)
--------------------------------------------------
간호사 정리본 엑셀에서 31개짜리 과거력 체크리스트(1DM, 2HiBP, … 21MentalRet,
27Epilepsy, 28DrugAbu, 29AlcoAbu)를 발견하고 «연구의 제외기준을 그대로 담은 구조화된
소스»라고 기뻐했다. 헤더도 정확히 맞았고 파서도 값을 잘 읽었다. 그런데 전 6,260행을
세어 보니 **30개 열이 전부 '0'**이고 예외는 단 두 건이었다 — 아무도 채운 적 없는 빈
템플릿 열이었던 것이다. 그대로 실었다면 «전체 환자가 정신질환·뇌전증·물질사용장애 전부
음성»이라는 **틀린 데이터**가 되어 제외기준 판정을 통째로 망쳤을 것이다.

핵심은 «열이 있다»와 «값이 있다»가 다른 명제라는 것. 헤더 존재는 아무것도 보증하지
않는다. 그래서 새 소스를 붙이기 전에 이걸 먼저 돌린다(수 초면 끝난다).

무엇을 잡는가
-------------
  [DEAD]     전 행이 비어 있음
  [CONSTANT] 값이 단 하나뿐 (예: 전부 '0') — 위 사고가 정확히 이 유형
  [SPARSE]   채움률이 임계값 미만 (기본 5%)
  [BINARYish] 값이 2종뿐 — 코딩(0/1, Y/N)인지 확인 필요

사용
----
  python profile_columns.py FILE.xlsx                 # 모든 시트
  python profile_columns.py FILE.xlsx --sheet Sheet2
  python profile_columns.py FILE.csv
  python profile_columns.py FILE.xlsx --password 1234 # 암호화 xlsx (msoffcrypto-tool)
  python profile_columns.py FILE.xlsx --only-suspect  # 문제 있는 열만
  python profile_columns.py FILE.xlsx --max-rows 5000 # 큰 파일 표본만

종료코드: 의심 열이 하나라도 있으면 1 (CI/게이트에서 쓰라고).
"""
import argparse
import csv
import io
import os
import sys
from collections import Counter

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SPARSE_DEFAULT = 0.05
# 비결측의 이 비율 이상이 한 값이면 «거의 상수»로 본다(값 종류가 3개 이하일 때).
NEAR_CONSTANT_DEFAULT = 0.98


def _open_workbook(path, password=None):
    import openpyxl
    if password:
        import msoffcrypto
        buf = io.BytesIO()
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=password)
            off.decrypt(buf)
        buf.seek(0)
        return openpyxl.load_workbook(buf, read_only=True, data_only=True)
    return openpyxl.load_workbook(path, read_only=True, data_only=True)


def _iter_xlsx(path, sheet=None, password=None, max_rows=None):
    wb = _open_workbook(path, password)
    names = [sheet] if sheet else wb.sheetnames
    for name in names:
        ws = wb[name]
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            continue
        yield name, [("" if h is None else str(h).strip()) for h in header], rows, max_rows


def _iter_csv(path, max_rows=None):
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            header = next(reader)
        except StopIteration:
            return
        yield os.path.basename(path), [h.strip() for h in header], reader, max_rows


def profile(header, rows, max_rows=None):
    """열마다 (채운 개수, 전체, 서로 다른 값 Counter)."""
    counters = [Counter() for _ in header]
    filled = [0] * len(header)
    n = 0
    for row in rows:
        if max_rows is not None and n >= max_rows:
            break
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        n += 1
        for i in range(len(header)):
            v = row[i] if i < len(row) else None
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            filled[i] += 1
            if len(counters[i]) <= 25:      # 상한 -- 고유값이 많은 열은 세부까지 볼 필요 없음
                counters[i][str(v)[:40]] += 1
    return filled, counters, n


def classify(filled, distinct_n, total, sparse_threshold, top_share=0.0,
             near_constant=NEAR_CONSTANT_DEFAULT):
    """top_share = 비결측 중 «가장 흔한 값»의 비율.

    ⚠️ CONSTANT(distinct=1)만 보면 «거의» 한쪽으로 몰린 열을 놓친다. 판정 규칙이
    95%만 한쪽으로 떨어지는 결함이 실제로 있었다(2026-08-19: 독거 여부가
    값 있는 1,631건 전부 'No'였고, 다른 사례는 98.6%였다). 그래서 CONSTANT를
    비율로 일반화한다 -- distinct=1은 top_share=1.0인 특수한 경우일 뿐이다."""
    if total == 0:
        return "EMPTYFILE"
    if filled == 0:
        return "DEAD"
    if distinct_n == 1:
        return "CONSTANT"
    if filled / total < sparse_threshold:
        return "SPARSE"
    if distinct_n <= 3 and top_share >= near_constant:
        return "SKEWED"
    if distinct_n == 2:
        return "BINARYish"
    return "OK"


SUSPECT = {"DEAD", "CONSTANT", "SKEWED", "SPARSE", "EMPTYFILE"}


def run(path, sheet=None, password=None, max_rows=None, sparse=SPARSE_DEFAULT,
        only_suspect=False, near_constant=NEAR_CONSTANT_DEFAULT):
    ext = os.path.splitext(path)[1].lower()
    source = _iter_csv(path, max_rows) if ext == ".csv" else _iter_xlsx(path, sheet, password, max_rows)

    n_suspect = 0
    for sheet_name, header, rows, cap in source:
        filled, counters, total = profile(header, rows, cap)
        print(f"\n=== {os.path.basename(path)} :: {sheet_name}  ({total} rows x {len(header)} cols) ===")
        for i, h in enumerate(header):
            if not h:
                continue
            distinct = counters[i]
            top_share = (distinct.most_common(1)[0][1] / filled[i]) if filled[i] else 0.0
            verdict = classify(filled[i], len(distinct), total, sparse,
                               top_share, near_constant)
            if verdict in SUSPECT:
                n_suspect += 1
            elif only_suspect:
                continue
            pct = (100.0 * filled[i] / total) if total else 0.0
            top = ", ".join(f"{v!r}x{c}" for v, c in distinct.most_common(3))
            more = "+" if len(distinct) > 25 else ""
            print(f"  [{verdict:9s}] {h[:44]:44s} fill={pct:5.1f}%  distinct={len(distinct)}{more}  {top}")
    print(f"\n의심 열 {n_suspect}개"
          + ("" if n_suspect else " -- 값 분포상 바로 쓸 수 있어 보임"))
    return 1 if n_suspect else 0


def main():
    ap = argparse.ArgumentParser(description="열 채움률·상수 여부 프로파일러")
    ap.add_argument("path")
    ap.add_argument("--sheet")
    ap.add_argument("--password")
    ap.add_argument("--max-rows", type=int)
    ap.add_argument("--sparse", type=float, default=SPARSE_DEFAULT,
                    help=f"이 비율 미만 채움이면 SPARSE (기본 {SPARSE_DEFAULT})")
    ap.add_argument("--near-constant", type=float, default=NEAR_CONSTANT_DEFAULT,
                    help="비결측의 이 비율 이상이 한 값이면 SKEWED (기본 0.98, 0이면 끔)")
    ap.add_argument("--only-suspect", action="store_true")
    a = ap.parse_args()
    if not os.path.exists(a.path):
        print(f"[ABORT] 파일 없음: {a.path}")
        return 2
    return run(a.path, a.sheet, a.password, a.max_rows, a.sparse, a.only_suspect,
               a.near_constant)


if __name__ == "__main__":
    sys.exit(main())
