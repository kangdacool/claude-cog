# -*- coding: utf-8 -*-
"""탐색 국면의 「후보 기각」을 기록하고, 연속 기각이 쌓이면 「기준을 감사하라」고 막는다.

WHY THIS EXISTS
CORE §2⑨ 「탐색 국면에서 기각 기준을 성공 사례도 통과 못 할 높이로 올린다」는
**다섯 번 재발했다**(2026-08-19 두 번 · 08-20 참신성 판정 · 08-21 보고 단계 ·
08-23 한 프로젝트에서 여섯 연속 기각). CORE 자신이 적어 둔 규칙 —
「기각할 때 3줄을 「쓴다」. 안 쓰면 기각이 아니다」 — 은 산문이라 매 세션 리셋된다.
CORE §5: 「규칙이 반복해서 깨지면 더 크게 쓰지 말고 코드로 옮길 수 있는지부터 물을 것.」
그래서 옮긴다.

이 도구가 강제하는 것 두 가지:
  1. 기각을 기록하려면 **세 줄이 다 있어야 한다** — 없으면 거부한다(기각이 성립 안 함).
     ① 쓴 기준 ② known-good에 대면 결과 ③ 죽인 사실이 「검증」인가 「추측」인가
  2. **연속 3회**가 되면 status·add가 큰 경고를 낸다 — 의심 대상은 후보가 아니라 「기준」이다.
     (CORE: 「세 번째부터는 후보를 더 보지 말고 기준을 감사한다」)

⚠ 이 도구는 판정하지 않는다. 기각을 「막지」도 않는다(--force 없이도 3회 이후 기록은 된다).
   하는 일은 「세 줄을 쓰게 만들고, 연속 횟수를 세어 보여주는 것」뿐이다.
   known-good 대면에서 통과 못 하면 그 기각은 무효라는 판단은 사람이 한다.

USAGE
    python reject_ledger.py add --project myproject \
        --candidate "청소년기 외로움 -> 성인기 자살생각" \
        --criterion "선행 논문이 같은 주장을 이미 냈으면 기각" \
        --known-good "Paper B(재현+신규로 accept)에 대보면 통과 못 함 -> 기준이 과함" \
        --evidence verified --note "Jin 2026 Arch Suicide Res, Add Health"

    python reject_ledger.py survive --project myproject --candidate "some idea"
    python reject_ledger.py status  --project myproject
    python reject_ledger.py list    --project myproject

기록 위치: agent/state/reject_ledger.jsonl (프로젝트별로 한 파일에 섞어 저장, project 필드로 구분)

EXIT STATUS: add/survive/list = 0. status = 연속 3회 이상이면 1(그 외 0).
"""
import argparse
import json
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
LEDGER = os.path.join(os.path.dirname(HERE), "state", "reject_ledger.jsonl")
STREAK_ALARM = 3          # CORE §2⑨: 「세 번째부터는 기준을 감사한다」

BANNER = "=" * 72


def _load():
    if not os.path.exists(LEDGER):
        return []
    out = []
    with open(LEDGER, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def _append(rec):
    os.makedirs(os.path.dirname(LEDGER), exist_ok=True)
    with open(LEDGER, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def _streak(rows, project):
    """마지막 「생존」 이후의 연속 기각 수. 생존이 없으면 전체 기각 수."""
    n = 0
    for r in reversed([r for r in rows if r.get("project") == project]):
        if r.get("kind") == "survive":
            break
        if r.get("kind") == "reject":
            n += 1
    return n


def _alarm(n, project):
    if n < STREAK_ALARM:
        return False
    print(BANNER)
    print("  [!] %s: 연속 기각 %d회 -- CORE §2⑨ 임계(%d) 도달" % (project, n, STREAK_ALARM))
    print(BANNER)
    print("  의심 대상은 「후보」가 아니라 「기준」이다.")
    print("  후보를 더 보지 말고, 지금까지 쓴 기준을 감사할 것:")
    print("   1. 그 기준을 known-good(이미 성공한 선례)에 대면 통과하는가?")
    print("      통과 못 하면 그 기각들은 무효다.")
    print("   2. 참신성 판정이면: 진짜 사망은 「자료·주기·노출정의·결과·분석이 전부」")
    print("      겹칠 때뿐이다. 6각도(새 주기·새 결과·측정 개선·효과변경·방법 개선·")
    print("      독립 재현) 중 하나만 서면 산다 -- 죽이기 전에 6개를 「써 본다」.")
    print("   3. '사망'과 '검정력 부족'을 구분했는가? 후자는 고칠 수 있다.")
    print(BANNER)
    return True


def cmd_add(a):
    missing = [n for n, v in (("--criterion", a.criterion),
                              ("--known-good", a.known_good)) if not (v or "").strip()]
    if missing:
        raise SystemExit(
            "기각이 성립하지 않는다 -- 다음이 비어 있다: %s\n"
            "  CORE §2⑨: 「기각할 때 3줄을 「쓴다」. 안 쓰면 기각이 아니다.」\n"
            "  ① 쓴 기준(--criterion)  ② known-good 대면(--known-good)\n"
            "  ③ 죽인 사실이 검증인가 추측인가(--evidence verified|guess)"
            % ", ".join(missing))
    rec = {"kind": "reject", "date": a.date or date.today().isoformat(),
           "project": a.project, "candidate": a.candidate,
           "criterion": a.criterion.strip(), "known_good": a.known_good.strip(),
           "evidence": a.evidence, "note": (a.note or "").strip()}
    _append(rec)
    rows = _load()
    n = _streak(rows, a.project)
    print("기록: [기각] %s / %s (연속 %d회)" % (a.project, a.candidate, n))
    if a.evidence == "guess":
        print("  [!] 죽인 근거가 「추측」이다 -- 검증 전에는 기각을 확정하지 말 것.")
    _alarm(n, a.project)
    return 0


def cmd_survive(a):
    _append({"kind": "survive", "date": a.date or date.today().isoformat(),
             "project": a.project, "candidate": a.candidate,
             "note": (a.note or "").strip()})
    print("기록: [생존] %s / %s -- 연속 기각 계수를 0으로 되돌린다." % (a.project, a.candidate))
    return 0


def cmd_status(a):
    rows = _load()
    proj = [r for r in rows if r.get("project") == a.project]
    if not proj:
        print("%s: 기록 없음" % a.project)
        return 0
    n = _streak(rows, a.project)
    nrej = sum(1 for r in proj if r.get("kind") == "reject")
    nsur = sum(1 for r in proj if r.get("kind") == "survive")
    guesses = [r for r in proj if r.get("kind") == "reject" and r.get("evidence") == "guess"]
    print("%s: 기각 %d · 생존 %d · **연속 기각 %d**" % (a.project, nrej, nsur, n))
    if guesses:
        print("  [!] 근거가 「추측」인 기각 %d건 -- 재검토 후보:" % len(guesses))
        for r in guesses:
            print("      - %s (%s)" % (r.get("candidate"), r.get("date")))
    return 1 if _alarm(n, a.project) else 0


def cmd_list(a):
    rows = [r for r in _load() if r.get("project") == a.project]
    if not rows:
        print("%s: 기록 없음" % a.project)
        return 0
    for r in rows:
        if r.get("kind") == "survive":
            print("  [생존] %s  %s" % (r.get("date"), r.get("candidate")))
            continue
        print("  [기각] %s  %s   (근거: %s)" % (r.get("date"), r.get("candidate"),
                                                r.get("evidence")))
        print("         기준     : %s" % r.get("criterion"))
        print("         known-good: %s" % r.get("known_good"))
        if r.get("note"):
            print("         비고     : %s" % r.get("note"))
    return 0


def main(argv=None):
    p = argparse.ArgumentParser(description="후보 기각 원장 (CORE §2⑨ 기계화)")
    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("add", help="기각을 기록한다 (세 줄이 없으면 거부)")
    a.add_argument("--project", required=True)
    a.add_argument("--candidate", required=True)
    a.add_argument("--criterion", default="", help="① 쓴 기준(한 문장)")
    a.add_argument("--known-good", dest="known_good", default="",
                   help="② 이 기준을 이미 성공한 선례에 대면 어떻게 되는가")
    a.add_argument("--evidence", choices=["verified", "guess"], required=True,
                   help="③ 죽인 사실이 검증인가 추측인가")
    a.add_argument("--note", default="")
    a.add_argument("--date", default="")
    a.set_defaults(fn=cmd_add)

    s = sub.add_parser("survive", help="후보가 살아남았다 -- 연속 계수 초기화")
    s.add_argument("--project", required=True)
    s.add_argument("--candidate", required=True)
    s.add_argument("--note", default="")
    s.add_argument("--date", default="")
    s.set_defaults(fn=cmd_survive)

    t = sub.add_parser("status", help="연속 기각 수와 경고")
    t.add_argument("--project", required=True)
    t.set_defaults(fn=cmd_status)

    l = sub.add_parser("list", help="기록 전체")
    l.add_argument("--project", required=True)
    l.set_defaults(fn=cmd_list)

    args = p.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
