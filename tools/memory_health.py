"""
##################################################################
#####  MEMORY HEALTH — 기억 시스템 자체의 드리프트 검사       #####
##################################################################

`feedback/`는 읽혀야 작동하므로 줄이 늘면 매 세션 비용이 오른다. 그런데 세션마다
산문을 덧붙이는 게 기본값이라, 규칙("코드로 옮겨라")만으로는 안 지켜졌다 —
그래서 검사로 옮긴다.

기계로 확실한 것만 본다:
  1. 줄 수 / 예산 초과 — **파일당 상한 + 묶음 상한**(총합은 참고치)
  2. MEMORY.md·CORE.md에 적힌 규모 서술("2,300줄")이 실제와 어긋남
  3. 끊어진 [[wikilink]] — 대응하는 name: 을 가진 파일이 없음
  4. 고아 파일 — MEMORY.md 인덱스에서 참조되지 않는 feedback/reference 파일
  5. 인덱스가 가리키는데 존재하지 않는 경로
  6. 오래 묵은 pending/TODO (프로젝트 메모)

Usage:  python tools/memory_health.py [--root <repo>/agent]
Exit 1 이면 고칠 것이 있다.
"""

import argparse
import re
import sys
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# CORE_BUDGET 은 «적정 길이»가 아니라 **드리프트 탐지선**이다.
# 2026-07-28 에 나는 그때의 CORE 길이(180)를 그대로 상한으로 굳혔고, 근거가 없었기 때문에
# 만든 그 세션에 바로 190 으로 어겼다. 길이는 원래 지킬 수 있는 규칙이 아니다.
# 진짜 규율은 CORE.md 머리말에 있다 — «§2·§3·§4 만, 서사는 topical, 검사 가능하면 tool,
# 새 항목은 기존 항목을 밀어내야 들어온다». 아래 숫자는 그 규율이 무너졌을 때 울리는 종이다.
# 2026-08-19: 200 → 250 (연구자 승인 "조금 늘려도 돼"). 실측 245.
#   근거: §2가 9개로 늘었고 «가르는 기준»이 붙은 항목(⑧·⑨)은 한 줄로 못 줄인다 —
#   금지만 적으면 «필요한 것»까지 지우게 되므로 판단 기준이 본문에 있어야 한다.
#   실제로 이번에 ⑧을 압축해 262 → 245로 «밀어낸 뒤» 올린 것이다(먼저 밀고 나서 올렸다).
#   ⚠️ 다음에 또 넘으면 올리지 말고 밀어낼 것. 상향은 두 번째부터 «미룸»이 된다.
CORE_BUDGET = 250
# 2026-08-03: 3900 → 4000. **정당한 새 상한이 아니라 «미룸»의 기록이다** — 당시엔 34개
#   파일을 실제로 읽지 않고, pipeline_workflow(541줄)·manuscript_rules(429줄) 두 파일이
#   범인일 거라고 짐작만 했다.
# 2026-08-16: 5822 → 5600 → **5820** (실측 5816 + 여유 4). 포크가 34개 파일을 Read로 전수조사해 6개·~280줄의
#   드레인 후보를 냈다(figure_table_to_word.md 등) — 근거는 있었으나 **포크 보고를 그대로
#   실행하지 않고 직접 열어 재검증**했더니 과대추정이었다: numbers_generated_not_typed.md는
#   "삭제 후보"였지만 실제로 흡수처(cross-verify 스킬)에 그 내용이 아직 없어 지금 지우면
#   순삭제다(보류) · revision_workflow.md는 "한 프로젝트 전용 구현"이라 했지만 대부분 재사용 가능한
#   구현 패턴이고 그 프로젝트는 reference impl일 뿐(보류) · 실제로 안전하게 뺀 건 2건뿐, ~19줄
#   (figure_table_to_word.md의 2026-07-20 전환 완료 changelog, generated_reasoning_vs_checking.md의
#   scope_calls_are_the_users.md와 중복된 사례 1개를 링크로 축약). **교훈: 포크의 "중복/이관
#   후보" 판정은 파일 존재·주제 유사성까지만 보고 실제 흡수 여부·재사용성은 못 본다 — 실행
#   전 최소 1개는 직접 열어 검증할 것.** 이후 `numbers_generated_not_typed.md`는 흡수처
#   (`/cross-verify` 스킬)에 실제로 내용을 옮긴 뒤 포인터로 축약(순삭제 아님) — 추가 17줄 절감,
#   실측 5799. 다음 4개 후보(manuscript_rules·ppt_rules 각각 재확인)도 열어봤더니 **전부
#   중복 아님**으로 판정(다른 툴체인·스킬에 없는 내용) — 더 뺄 게 없다는 뜻이지 게을러서가 아니다.
#   "비중복"이 "더 안 줄여도 된다"가 아니었다 — 사용자 지시로 실제 정독 드레인을 했다:
#   `revision_workflow.md`(202→187, §5를 r-technical-gotchas로 이관·§9/§10 서사 압축·Word
#   재저장 빈페이지 교훈을 §2로 재배치) + `poster_rules.md` 메타발언 1줄 제거. 추가 39줄 절감,
#   실측 5783. **차이는 2026-08-14에 이미 드레인된 파일(포크가 헷갈렸던 4개)과 안 된 파일
#   (revision_workflow)이었다** — 드레인 이력 없는 파일이 진짜 후보다, 다음엔 거기부터 볼 것.
#   상한을 실측치(5783) 위 약간의 여유로 둠.
# 2026-08-20: 예산을 «올리지 않았다». 초과(6,812 대 5,800)의 원인을 재보니 오래된
#   군더더기가 아니라 «누적 속도»였다 — 7일간 +1,368줄이 들어왔고 되돌린 건 3분의 1.
#   세션이 서로를 못 보니 각자 쓰기만 한다(같은 사고가 두 파일에 통째로 들어간 사례 2건).
#   그래서 숫자를 올리는 대신 둘을 했다:
#     (a) 이 도구에 --drain 을 붙여 «무엇을» 뺄지 같이 내게 했다. "초과다"만 말하면
#         다음 한 걸음이 없어 아무도 안 뺀다.
#     (b) CLAUDE.md 쓰기 규칙에 「feedback에 교훈을 추가한 세션은 이 도구를 돌리고,
#         초과면 먼저 빼고 넣는다」를 박았다 — 드레인을 «청소»가 아니라 «쓰기의 일부»로.
#   상한을 올리는 것은 미룸이다. 위 CORE_BUDGET 이력이 그 증거다.
# 2026-08-20 (같은 날, 후속): **총합 예산을 폐기하고 «파일당 + 묶음»으로 바꿨다.**
#   드레인을 하다 깨달은 것: SESSION START PROTOCOL은 feedback/ 전체를 읽지 않는다.
#   CORE + MEMORY + 프로젝트 MEMORY + 작업별 2~4개만 읽는다. 원고 작업이면
#   manuscript_rules + report_content_discipline + data_integrity를 읽고, 그 비용은
#   **총합이 6,300이든 5,800이든 똑같다.** 즉 총합 예산은 실제 읽기 비용과 연결돼 있지
#   않았고, 그래서 «지키면 이득이 없고 어기면 죄책감만 남는» 숫자였다(두 번 상향한 이유).
#   실측(2026-08-20): 묶음은 763 ~ 1,696줄(CORE+MEMORY 512 포함). 총합 6,333.
#   → 진짜 위험 신호는 «한 파일이 비대해지는 것»과 «한 묶음이 무거워지는 것»이다.
#   ⚠️ 아래 두 숫자도 «적정 길이»가 아니라 드리프트 탐지선이다. 넘으면 올리지 말고 민다.
# 2026-08-20 (세 번째): **줄 수만 재면 MEMORY.md를 놓친다.** 실측하니 MEMORY.md가 267줄인데
#   48,469자(줄당 182자 — CORE의 54자, feedback 평균 50자의 3.4배)로 **시스템에서 가장 비싼
#   파일**이었다(~28–39k 토큰, 게다가 «작업 종류와 무관하게» 매 세션). 원인은 인덱스 항목이
#   훅이 아니라 «요약»이라서 — 400자 넘는 한 줄이 41개, 최대 1,691자였다. 줄 기반 게이트는
#   이 축에 구조적으로 눈이 없다(「통과한 검사는 그 검사가 본 것만 보증한다」).
#   → 인덱스 항목에 «글자» 상한을 건다. 넘으면 그건 인덱스가 아니라 문서다 — 가리키는 파일에
#   넣거나(대개 이미 있다), 도구라서 문서가 없으면 그 도구의 문서를 만들 신호다.
# ⭐ 백테스트(2026-08-20) — 이 숫자들을 «한 시점의 실측 + 여유»로 정한 게 걸려서, 25일 이력의
#   feedback 커밋 43건 전부에 되돌려 돌려봤다(각 시점의 CLAUDE.md 매핑표로 묶음을 계산).
#     묶음 1,800 : 43건 중 **1건**만 울림(08-14, 1,935) → 잘 맞는다. 진짜 트립와이어다.
#     파일  450  : 31건에서 울렸으나 **매번 «한 파일»을 지목**했다 — 08-03~19는 계속
#                  pipeline_workflow, 08-19부터 data_integrity. "전부 초과"가 아니라
#                  "이 파일 하나"라 실행 가능하다. 실제로 그 둘이 드레인이 필요했던 파일이다.
#     항목  400  : 43/43 전부 울림. **숫자가 낮은 게 아니라 그 관행이 아예 없었다**(첫 커밋부터
#                  최대 617자, 최대 1,691자까지). 2026-08-20에 250자로 내렸으니 판정은 이후에.
#   ⭐ **세 축이 «동시에» 울린 커밋은 43건 중 1건뿐** — 숫자가 전반적으로 낮았다면 셋이 늘 같이
#   울렸을 것이다. 즉 축이 서로 다른 것을 보고 있다는 증거이기도 하다.
#   ⭐ 백테스트가 드러낸 것: pipeline_workflow가 **08-03~08-19 17일 연속** 450줄을 넘긴 채
#   아무도 모르고 지나갔다. 게이트가 없으면 그 상태는 보이지 않는다 — 그게 이 숫자들의 값이다.
#   (재현: 이력에서 `git show <rev>:agent/feedback/*`로 각 시점을 재는 스크립트. 필요하면 다시 짠다.)
FILE_BUDGET = 450          # feedback 파일 하나. 넘으면 그 파일을 쪼개거나 도구/스킬로 이관
BUNDLE_BUDGET = 1800       # CORE + MEMORY + 작업유형 묶음. 세션이 실제로 읽는 양

# ⚠️ 예외는 «전역 상한»을 올리지 않고 여기에 둔다. FILE_BUDGET 을 600 으로 올리면
#    35개 파일의 가드가 «전부» 느슨해진다 — 위 CORE_BUDGET 이력이 경고하는 바로 그
#    미룸이 된다. 특정 작업의 교훈 유입이 많다면 그 파일만 넓힌다.
#
# 2026-08-28 연구자 결정: 「manuscript_rules는 약간 더 예산을 늘려도 될 것 같다.
#   너가 «가장 잘해야 하는 작업» 중 하나거든. 그리고 앞으로도 많은 피드백이 있을 것이고.」
#   근거: 원고 피드백은 «한 번에 파일 단위»로 온다(psy 260826: 13항목·세부 40여 개).
#   그 교훈이 들어갈 곳은 이 파일 하나인데 2026-08-27에 448/450 으로 차서 새 교훈을
#   막고 있었다(실제로 「보고 방식」 교훈 하나가 갈 곳이 없어 미기록으로 남았다).
FILE_BUDGET_OVERRIDE = {"manuscript_rules.md": 600}

# ⚠️ 파일 상한만 올리면 «묶음»에서 막힌다. 2026-08-28 실측: 원고 묶음이 이미
#    1,648/1,800 이고 같은 묶음의 data_integrity 도 449/450 이라 여유가 152줄뿐이었다.
#    -- 그래서 둘을 «같이» 올린다. 라벨은 CLAUDE.md 매핑표 첫 칸과 정확히 같아야 한다.
BUNDLE_BUDGET_OVERRIDE = {"원고 · 보고서 · 논문요약서": 2000}


def file_budget(name):
    return FILE_BUDGET_OVERRIDE.get(name, FILE_BUDGET)


def bundle_budget(label):
    return BUNDLE_BUDGET_OVERRIDE.get(label, BUNDLE_BUDGET)
ENTRY_BUDGET = 400         # MEMORY.md 인덱스 «한 항목». 훅이지 요약이 아니다
# 실행 결과(2026-08-20): 48,469 → 27,779자(-43%), 400자 초과 41개 → 0개, 평균 180자.
#   세션 시작 총비용 ~80–115k → ~67–97k 토큰. feedback을 깎는 것과 다른 점은 **작업 종류와
#   무관하게 매 세션** 빠진다는 것이다.
#   ⚠️ 줄이기 «전에» 확인할 것: 인덱스 항목의 코드토큰·숫자가 대상 문서에 실제로 있는가.
#   없으면 축약이 아니라 순삭제다. 실제로 2건이 인덱스에만 있었다(hwpx의 `save_paginated()`
#   블록, `utf-8-sig` BOM 함정) — 문서로 «먼저» 옮기고 나서 줄였다.
#   ⚠️ 도구 항목은 그 줄이 «곧 문서»라 같은 잣대를 대면 통과 불가가 된다(= 늘 빨간 게이트).
#   그래서 .md를 가리키면 FIX, 도구를 가리키면 NOTE로 가른다. 실제로 열어 재니 6개 중 5개가
#   이미 자기 머리말에 1,150~1,525자를 갖고 있었다 — 인덱스가 그걸 한 번 더 쓰고 있었을 뿐.
#   (⚠️ `ast.get_docstring`으로 재면 «자체문서 0자»라는 오판이 난다 — `# -*- coding -*-` 줄과
#    이 랩의 `#` 섹션헤더 관습을 못 본다. 도구가 아니라 검출기가 틀린 경우였다.)
# 묶음 정의는 이 파일에 박지 않는다 — 정본은 CLAUDE.md의 SESSION START PROTOCOL 매핑표다.
# (「감사 도구는 산문이 아니라 레지스트리로」 — 여기 복사하면 표가 바뀔 때 조용히 낡는다.)
# 이 파일은 <repo>/agent/tools/ 에 있다 -> parents[2] 가 저장소 루트.
REPO = Path(__file__).resolve().parents[2]
CLAUDE_MD = REPO / "claude-config" / "CLAUDE.md"
STALE_MONTHS = 6


def wc(p):
    return len(p.read_text(encoding="utf-8").splitlines())


def prose(p):
    """코드 펜스와 인라인 코드를 제거한 본문. 링크를 *설명하는* 산문(`[[wikilink]]`)이나
    코드 예시를 실제 링크로 오인하지 않기 위함."""
    txt = p.read_text(encoding="utf-8")
    txt = re.sub(r"```.*?```", "", txt, flags=re.S)
    return re.sub(r"`[^`\n]*`", "", txt)


def front_name(p):
    m = re.search(r"^name:\s*(\S+)\s*$", p.read_text(encoding="utf-8"), re.M)
    return m.group(1) if m else None


def triggered(root, fb_files):
    """SESSION START PROTOCOL이 «언제 읽어라»라고 말해주는 파일 집합.
    MEMORY.md 인덱스에 있는 것만으로는 부족하다 — 인덱스는 «존재»를 알려줄 뿐 «지금 읽어라»를
    말하지 않는다. 2026-08-20에 손으로 세어보니 6개 파일 335줄이 그 상태였다(내용은 전부
    살아 있는 규칙인데 아무 작업도 그것을 부르지 않았다). 「러너에 없으면 없는 도구다」의
    메모리판이라 검사로 옮긴다."""
    seed = set()
    if CLAUDE_MD.exists():                      # 작업 유형 매핑표
        # 숫자가 든 파일명(table1_...)을 놓치던 버그 -- [a-z_]+ 로는 안 잡힌다
        seed |= set(re.findall(r"`([a-z0-9_]+)\.md`", CLAUDE_MD.read_text(encoding="utf-8")))
    core = root / "feedback" / "CORE.md"
    if core.exists():                           # CORE 본문의 링크(§4 인덱스 포함)
        seed |= {s.replace("-", "-") for s in
                 re.findall(r"\[\[([a-z][a-z0-9-]*)\]\]", core.read_text(encoding="utf-8"))}
    names = {p: (front_name(p) or p.stem) for p in fb_files}
    hit = {p for p in fb_files
           if names[p] in seed or p.stem in seed or names[p].replace("-", "_") in seed}
    # 트리거된 파일이 링크하는 파일도 «그 작업 중에» 따라갈 수 있다(한 다리까지만 인정)
    for p in list(hit):
        for lk in re.findall(r"\[\[([a-z][a-z0-9-]*)\]\]", p.read_text(encoding="utf-8")):
            hit |= {q for q in fb_files if names[q] == lk or q.stem == lk.replace("-", "_")}
    return hit


def bundles(root):
    """CLAUDE.md 매핑표를 읽어 (작업유형, [파일], 줄수, [없는 파일]) 목록을 낸다.
    표가 정본이므로 여기서 파싱한다 — 복사해두면 표가 바뀔 때 조용히 낡는다."""
    if not CLAUDE_MD.exists():
        return None
    out = []
    for ln in CLAUDE_MD.read_text(encoding="utf-8").split("\n"):
        if not ln.startswith("|") or ln.count("|") < 3:
            continue
        cells = [c.strip() for c in ln.strip("|").split("|")]
        # ⚠ 부정 문맥의 파일명은 «읽지 않는» 파일이다 — 예산에 세면 안 된다.
        #   포스터 행의 "(슬라이드 덱과 다른 장르 — `ppt_rules.md` 아님)" 이 그 예다.
        #   2026-08-23: 이걸 세는 바람에 포스터 묶음이 317 줄 부풀어 헛 초과가 났다.
        NEG = ("아님", "아니", "제외", "말 것", "말고", "не")
        # ⚠ 매핑표 산문에는 «다른 저장소»의 경로도 등장한다 — 2026-09-07 실측: 다른 저장소의 행이
        #   「rubric 이 `prompts/rubric.md` 로 «코드 밖»에 있다」고 설명하는데, 그걸 «읽어야 할
        #   묶음 파일»로 파싱해 「없는 파일을 가리킨다」는 FIX 가 매 세션 떴다. 부정문맥 필터로는
        #   안 걸린다(부정이 아니라 «남의 경로»다). 이 예산이 세는 것은 agent/feedback·reference
        #   의 .md 뿐이므로 그 밖의 디렉터리를 가진 경로는 아예 세지 않는다.
        #   ⛔ 매핑표 쪽 표기를 고쳐서 우회하지 말 것 — 남의 경로가 산문에 나오는 일은 또 있다.
        OURS = ("feedback/", "reference/")
        files = []
        for m in re.finditer(r"`([a-z_/]+\.md)`", cells[-1]):
            tail = cells[-1][m.end():m.end() + 14]
            if any(tail.lstrip(" )·,").startswith(w) for w in NEG):
                continue
            f = m.group(1)
            if "/" in f and not f.startswith(OURS):
                continue
            files.append(f)
        if not files:
            continue
        tot, miss = 0, []
        for f in files:
            for cand in (root / "feedback" / Path(f).name, root / f):
                if cand.exists():
                    tot += wc(cand)
                    break
            else:
                miss.append(f)
        out.append((re.sub(r"\*\*|`", "", cells[0])[:38], files, tot, miss))
    return out


##################################################
#####  드레인 후보 — «초과다»만 말하면 아무도 안 뺀다  #####
##################################################
# 2026-08-20에 확인한 것: 예산 초과의 원인은 «오래된 군더더기»가 아니라 «누적 속도»였다.
# 08-16에 5,783까지 드레인해 놓고 나흘 만에 +1,508줄이 들어왔다 — 병렬 세션 12개가
# 각자 교훈을 적었고 되돌린 건 ~450줄뿐. 세션은 서로를 못 보므로 «같은 사고»가 여러
# 파일에 통째로 들어간다(실제로 그날 2건 발견: lmtp/JASA 사고, 그림 크기 검증 코드).
#
# 그런데 도구는 "draining이 늦었다"고만 말했다. 다음 한 걸음이 없으면 아무도 안 뺀다.
# 그래서 «무엇을» 뺄지 여기서 같이 낸다.
#
# ⚠️ 이 출력은 «후보»지 판정이 아니다. 2026-08-16에 포크의 중복 판정 6건을 그대로
# 실행하려다 재검증했더니 안전한 것은 2건뿐이었다. 주제 유사성은 흡수 여부·재사용성을
# 증명하지 않는다 — 실행 전 반드시 열어 볼 것.

_STOP = set("이 그 저 것 수 등 및 때 더 안 못 잘 다 를 을 는 은 가 이다 있다 없다 한다 "
            "된다 한 두 세 전 후 위 아래 같은 다른 모든 어떤".split())


def _sections(md_dir):
    """(파일, 제목, 줄수, 본문) — ## / ### 단위."""
    out = []
    for p in sorted(md_dir.glob("*.md")):
        lines = p.read_text(encoding="utf-8").split("\n")
        cur, buf, start = None, [], 0
        for i, ln in enumerate(lines):
            if re.match(r"^#{2,3} ", ln):
                if cur:
                    out.append((p.name, cur, i - start, "\n".join(buf)))
                cur, buf, start = ln.strip("# ").strip(), [], i
            elif cur:
                buf.append(ln)
        if cur:
            out.append((p.name, cur, len(lines) - start, "\n".join(buf)))
    return out


def _toks(t):
    return {w for w in re.findall(r"[가-힣A-Za-z]{2,}", t.lower()) if w not in _STOP}


def _index_entries(md):
    """MEMORY.md 를 «항목» 단위로 자른다. -> [(글자수, 이름, 가리키는 경로)]

    ⚠⚠ **한 항목은 여러 줄이다.** 2026-09-07 까지 이 검사는 물리적 «한 줄»을 재서
      642자짜리 항목(7줄 x 38~89자)을 그냥 통과시켰다 — 어느 «줄»도 400을 안 넘으니까.
      이 파일 머리말이 「줄 수로는 안 보이는 축」이라 적어 두고 정작 줄로 재고 있었다.
      실측: 그렇게 놓치고 있던 항목이 6개였다.
    """
    out, cur = [], None
    for l in md.split("\n"):
        if l.startswith("- "):
            if cur:
                out.append(cur)
            cur = [l]
        elif cur is not None:
            if not l.strip() or l[0] not in " \t":   # 빈 줄·새 블록에서 끊는다
                out.append(cur)
                cur = None
            else:
                cur.append(l.strip())               # 접힌 이어짐
    if cur:
        out.append(cur)
    res = []
    for e in out:
        body = " ".join(e)
        tgt = re.search(r"\]\(([^)]+)\)", body)
        nm = re.search(r"\[([^\]]{0,40})\]", body)
        res.append((len(body), nm.group(1) if nm else "?", tgt.group(1) if tgt else ""))
    return res


def drain_report(fb, index=None):
    import itertools
    secs = _sections(fb)
    print("=" * 68)
    print("드레인 후보 — ⚠️ 판정이 아니라 후보다. 실행 전 열어서 확인할 것")
    print("=" * 68)

    print("\n[1] 파일을 넘나드는 근접 중복 (병렬 세션이 서로를 못 봐서 생긴다)")
    n = 0
    for a, b in itertools.combinations(secs, 2):
        if a[0] == b[0] or a[2] < 6 or b[2] < 6:
            continue
        ta, tb = _toks(a[3]), _toks(b[3])
        if not ta or not tb:
            continue
        j = len(ta & tb) / len(ta | tb)
        if j >= 0.30:
            n += 1
            print(f"  자카드 {j:.2f}  (합쳐 {a[2] + b[2]}줄)")
            print(f"      {a[0]:38s} {a[1][:50]}")
            print(f"      {b[0]:38s} {b[1][:50]}")
    if not n:
        print("  없음 — 절 단위 중복은 정리돼 있다")

    print("\n[2] 파일당 상한을 넘은 파일의 «가장 긴 절» — 여기부터 민다")
    over = {p.name for p in sorted(fb.glob("*.md"))
            if len(p.read_text(encoding="utf-8").splitlines()) > file_budget(p.name)
            and p.name != "CORE.md"}
    if over:
        print(f"  상한 초과 파일: {', '.join(sorted(over))}")
        for f, t, ln, _ in sorted([s for s in secs if s[0] in over], key=lambda x: -x[2])[:8]:
            print(f"  {ln:>4}줄  {f:36s} {t[:44]}")
    else:
        print(f"  없음 — 모든 파일이 {FILE_BUDGET}줄 이하다")
    print("\n[2-b] 전체에서 가장 긴 절 (참고 — 상한 안이어도 이관 후보일 수 있다)")
    for f, t, ln, _ in sorted(secs, key=lambda x: -x[2])[:5]:
        print(f"  {ln:>4}줄  {f:36s} {t[:44]}")

    print("\n[4] 최근 7일 증가분 — 오래된 것보다 «방금 들어온 것»이 대개 압축 여지가 크다")
    import subprocess
    repo = str(fb.parent.parent)

    def git(*a):
        return subprocess.run(["git", "-C", repo, *a], capture_output=True,
                              text=True, encoding="utf-8", timeout=30).stdout or ""
    try:
        # ⚠️ HEAD~N으로 기간을 흉내내면 안 된다 -- 커밋 밀도가 날마다 달라 «7일»과 무관해진다.
        base = git("rev-list", "-1", "--before=7.days.ago", "HEAD").strip()
        if not base:
            print("  (7일 전 커밋이 없다 -- 저장소가 그보다 젊다)")
        else:
            rows = []
            for ln in git("diff", "--numstat", base, "HEAD", "--",
                          "agent/feedback/").splitlines():
                f = ln.split("\t")
                if len(f) == 3 and f[0].isdigit() and f[1].isdigit():
                    rows.append((int(f[0]) - int(f[1]), f[2].split("/")[-1]))
            tot = sum(d for d, _ in rows)
            print(f"  기준 {git('log', '-1', '--format=%h %ad', '--date=short', base).strip()}"
                  f"  → 합계 {tot:+d}줄")
            for d, f in sorted(rows, reverse=True)[:6]:
                print(f"  {d:+5d}줄  {f}")
            if not rows:
                print("  (증가 없음)")
    except Exception as e:
        print(f"  (git 조회 실패: {e})")

    # ⭐ 2026-09-07 신설. 그 전까지 이 보고서의 후보 넷이 «전부 feedback/» 대상이었고,
    #   정작 매 세션 읽히는 «가장 큰 파일»(MEMORY.md)은 드레인 시야 밖에 있었다.
    #   feedback 을 깎는 것과 다른 점: 인덱스는 **작업 종류와 무관하게** 매 세션 빠진다.
    print(f"\n[5] MEMORY.md 인덱스에서 {ENTRY_BUDGET}자를 넘는 «항목» (줄이 아니라 항목 단위)")
    if index is None or not Path(index).exists():
        print("  (MEMORY.md 를 못 찾았다)")
    else:
        fat = [(n, nm, t) for n, nm, t in
               _index_entries(Path(index).read_text(encoding="utf-8")) if n > ENTRY_BUDGET]
        if not fat:
            print("  없음 — 모든 항목이 훅 크기다")
        else:
            print(f"  {len(fat)}개 · 초과분 합계 {sum(n - ENTRY_BUDGET for n, _, _ in fat):,}자")
            for n, nm, t in sorted(fat, reverse=True):
                kind = "→ .md 로 옮긴다" if t.endswith(".md") else "⚠ 도구 항목 — 줄이면 순삭제"
                print(f"  {n:5d}자  {nm[:44]:<44} {kind}")

    print("\n뺄 때: 흡수처에 내용이 «실제로» 있는지 먼저 확인한다. 없으면 순삭제다.")
    print("포인터로 축약할 땐 고유 내용(그 파일에만 있는 팁)은 남긴다.")


def main(root):
    root = Path(root)
    fb, ref = root / "feedback", root / "reference"
    index = root / "MEMORY.md"
    problems, notes = [], []

    fb_files = sorted(fb.glob("*.md"))
    ref_files = sorted(ref.glob("*.md"))
    fb_lines = sum(wc(p) for p in fb_files)
    ref_lines = sum(wc(p) for p in ref_files)
    core = fb / "CORE.md"
    core_lines = wc(core) if core.exists() else 0

    idx_lines = wc(index) if index.exists() else 0
    print(f"feedback   {len(fb_files):3d} files  {fb_lines:5d} lines   (총합은 참고치 — 게이트는 파일당/묶음)")
    print(f"reference  {len(ref_files):3d} files  {ref_lines:5d} lines")
    print(f"CORE.md                  {core_lines:5d} lines   (budget {CORE_BUDGET})")
    if index.exists():
        ichars = len(index.read_text(encoding="utf-8"))
        print(f"MEMORY.md                {idx_lines:5d} lines  {ichars:7,}자   "
              f"(항목 상한 {ENTRY_BUDGET}자 — 줄 수로는 안 보이는 축)")
    else:
        print("MEMORY.md MISSING")

    # ---- 1a. CORE (항상 읽힘 → 하드 예산)
    if core_lines > CORE_BUDGET:
        problems.append(f"CORE.md {core_lines} > {CORE_BUDGET} lines — 항상 읽히는 파일이다. "
                        "§4 인덱스로 밀어내거나 코드/스킬로 옮길 것")
    if core_lines > 0.9 * CORE_BUDGET:
        notes.append(f"CORE.md {core_lines}/{CORE_BUDGET} — 예산의 90%를 넘었다. 다음 교훈은 "
                     "CORE에 붙이지 말고 topical 파일 + §4 인덱스 한 줄로")

    # ---- 1b. 파일당 상한 — 한 파일이 비대해지면 그 주제를 다루는 모든 세션이 값을 치른다
    for p in fb_files:
        n = wc(p)
        fb_cap = file_budget(p.name)
        if n > fb_cap and p.name != "CORE.md":
            problems.append(f"{p.name} {n} > {fb_cap} lines — 파일 하나가 너무 무겁다. "
                            f"절을 도구/스킬로 이관하거나 쪼갤 것 (초과 {n - fb_cap}줄)")

    # ---- 1b-1. 인덱스 «항목» 글자 상한 — 줄 수로는 안 보이는 축
    # 두 종류를 갈라 잡는다. 같은 잣대를 대면 도구 항목이 «통과 불가»가 되어
    # 늘 빨간 게이트가 된다(= 신호 0). 도구 항목은 그 줄이 «곧 문서»이기 때문이다.
    if index.exists():
        # ⚠ 예전에는 여기서 «물리적 한 줄»을 쟀다. 항목은 접혀서 여러 줄이므로
        #   642자짜리(7줄 x 38~89자)가 통과했다 — _index_entries 주석 참조.
        doc_fat, tool_fat = [], []
        for n, name, tgt in _index_entries(index.read_text(encoding="utf-8")):
            if n <= ENTRY_BUDGET:
                continue
            (doc_fat if tgt.endswith(".md") else tool_fat).append((n, name))
        if doc_fat:
            problems.append(
                f"MEMORY.md: {ENTRY_BUDGET}자 넘는 인덱스 항목 {len(doc_fat)}개 "
                f"(합 {sum(n for n, _ in doc_fat):,}자, 최대 {max(doc_fat)[0]:,}자 "
                f"「{max(doc_fat)[1]}」) — **가리키는 .md가 이미 있다.** 인덱스는 «무엇이고 "
                "언제 여는가»까지이므로 내용은 그 파일로 옮긴다")
        if tool_fat:
            notes.append(
                f"MEMORY.md: 도구를 가리키는 긴 항목 {len(tool_fat)}개 "
                f"(합 {sum(n for n, _ in tool_fat):,}자, 최대 {max(tool_fat)[0]:,}자 "
                f"「{max(tool_fat)[1]}」) — 이 줄이 «곧 그 도구의 문서»라 지금 줄이면 순삭제다. "
                "줄이려면 먼저 `reference/`에 그 도구 문서를 만들 것(그러면 사용할 때만 읽힌다)")

    # ---- 1b-2. 트리거 없는 파일 — 인덱스에 있어도 «지금 읽어라»가 없으면 안 읽힌다
    trig = triggered(root, fb_files)
    for p in fb_files:
        if p.name != "CORE.md" and p not in trig:
            problems.append(f"{p.name}: SESSION START PROTOCOL에 트리거가 없다 "
                            "— CLAUDE.md 매핑표나 CORE §4에 한 줄을 넣거나, 그 교훈이 «쓰이는 순간에 "
                            "읽히는 파일»로 흡수할 것 (MEMORY.md 인덱스만으로는 안 읽힌다)")

    # ---- 1c. 묶음 상한 — 세션이 «실제로» 읽는 양(CORE + MEMORY + 작업유형 파일들)
    bs = bundles(root)
    if bs is None:
        notes.append(f"CLAUDE.md를 못 찾아 묶음 검사를 건너뛰었다 ({CLAUDE_MD})")
    else:
        base = core_lines + idx_lines
        worst = max(bs, key=lambda b: b[2]) if bs else None
        for label, files, tot, miss in bs:
            if miss:
                problems.append(f"CLAUDE.md 매핑표 「{label}」이 없는 파일을 가리킨다: "
                                f"{', '.join(miss)}")
            bcap = bundle_budget(label)
            if tot + base > bcap:
                problems.append(f"묶음 「{label}」 {tot + base} > {bcap} lines "
                                f"(CORE+MEMORY {base} + {' + '.join(Path(f).stem for f in files)}) "
                                "— 이 작업을 하는 모든 세션이 이만큼 읽는다")
        if worst:
            ## ⚠️ 「최대」는 예외를 받은 묶음이 아니라 «상한에 가장 가까운» 묶음이어야
            ##    한다 -- 예외 묶음이 늘 1등으로 찍히면 나머지가 안 보인다.
            tight = max(bs, key=lambda b: (b[2] + base) / bundle_budget(b[0]))
            print(f"최대 묶음                 {tight[2] + base:5d} lines   "
                  f"(budget {bundle_budget(tight[0])})  「{tight[0]}」")
            for lb in BUNDLE_BUDGET_OVERRIDE:
                hit = next((b for b in bs if b[0] == lb), None)
                if hit:
                    print(f"  예외 묶음                {hit[2] + base:5d} lines   "
                          f"(budget {bundle_budget(lb)})  「{lb}」")

    # ---- 2. 규모 서술이 실제와 어긋나는가
    # 어떤 숫자가 무엇을 가리키는지 문맥으로 귀속시키려 했더니("feedback"·"CORE"가 한 문장에
    # 같이 나온다) 오탐만 났다. 귀속을 포기하고 **추적 중인 실제값 어느 것과도 맞지 않을 때만**
    # 플래그한다 — 드리프트는 그대로 잡히고 오탐은 0이 된다.
    actuals = {"feedback": fb_lines, "reference": ref_lines,
               "CORE": core_lines, "MEMORY": wc(index) if index.exists() else 0}
    for p in (index, core):
        if not p.exists():
            continue
        for m in re.finditer(r"([0-9][0-9,]{2,})\s*줄", p.read_text(encoding="utf-8")):
            claimed = int(m.group(1).replace(",", ""))
            if not any(abs(claimed - v) <= max(3, 0.15 * v) for v in actuals.values()):
                problems.append(f"{p.name}: '{claimed}줄'이 추적 중인 어떤 실제값과도 다르다 "
                                f"(실제: {', '.join(f'{k} {v}' for k, v in actuals.items())})")

    # ---- 3. 끊어진 wikilink
    # kebab-case 슬러그만 링크로 본다. 코드·산문 속 대괄호([[2]], [["p"]], [[BLANK]])를
    # 링크로 오인하면 목록이 신뢰를 잃고 아무도 안 본다.
    LINK = re.compile(r"\[\[([a-z][a-z0-9-]*)\]\]")

    # 이름 색인은 넓게(프로젝트 메모까지 — 공유 파일이 이들을 가리킬 수 있다),
    # 링크 검사 대상은 공유 계층만. 프로젝트 메모끼리의 링크는 관습이 느슨해
    # (`[[some-project-status]]`) 기계로 판정하면 오탐만 쌓인다.
    names = {}
    for p in fb_files + ref_files + sorted(root.glob("*.md")):
        n = front_name(p)
        if n:
            names[n] = p
        elif p.parent.name in ("feedback", "reference"):
            problems.append(f"{p.relative_to(root)}: frontmatter에 name: 없음 "
                            "— wikilink 대상이 못 되어 이 파일을 가리키는 링크가 전부 끊긴다")
        names.setdefault(p.stem.replace("_", "-"), p)
    if (root / "projects").exists():
        for p in sorted((root / "projects").rglob("*.md")):
            n = front_name(p)
            if n:
                names.setdefault(n, p)
    # ⭐ 스킬도 wikilink 대상이다. 장르 절차서를 Skill 로 옮기면(본문이 «쓸 때만» 로드된다)
    #   [[ppt-rules]] 가 여기서만 「끊어졌다」고 나온다 — 링크는 멀쩡하고 검사기가 몰랐던 것이다.
    skills = Path.home() / ".claude" / "skills"
    if skills.exists():
        for d in sorted(skills.iterdir()):
            f = d / "SKILL.md"
            if f.is_file():
                names.setdefault(front_name(f) or d.name, f)

    for p in fb_files + ref_files + ([index] if index.exists() else []):
        for link in sorted(set(LINK.findall(prose(p)))):
            if link not in names:
                problems.append(f"{p.relative_to(root)}: 끊어진 링크 [[{link}]]")

    # ---- 4/5. 인덱스 ↔ 파일 대조
    if index.exists():
        itxt = index.read_text(encoding="utf-8")
        linked = set(re.findall(r"\(((?:feedback|reference)/[^)]+\.md)\)", itxt))
        for rel in sorted(linked):
            if not (root / rel).exists():
                problems.append(f"MEMORY.md: 존재하지 않는 경로 링크 {rel}")
        for p in fb_files + ref_files:
            rel = p.relative_to(root).as_posix()
            if rel not in linked and p.name != "CORE.md":
                problems.append(f"고아 파일 (MEMORY.md 인덱스에 없음): {rel}")

    # ---- 6. 묵은 pending
    # 산문에서 날짜를 긁으면 안 된다 — 본문의 '2023-07-01'(군위군 행정구역 이관 같은 사실)을
    # 작업 날짜로 오인해, 이틀 전에 갱신된 활발한 메모를 "36개월 묵음"으로 신고했다.
    # 실제 신호는 **파일 mtime**이다.
    today = date.today()
    for p in sorted((root / "projects").rglob("*.md")) if (root / "projects").exists() else []:
        if not re.search(r"pending|TODO|미완|대기", p.read_text(encoding="utf-8"), re.I):
            continue
        mtime = date.fromtimestamp(p.stat().st_mtime)
        age = (today.year - mtime.year) * 12 + (today.month - mtime.month)
        if age >= STALE_MONTHS:
            notes.append(f"{p.relative_to(root)}: pending/TODO인데 {age}개월간 수정 없음 "
                         f"(마지막 {mtime}) — 현재 파일 상태 확인 후 사용")

    # ---- 7. projects/ 폴더 ↔ MEMORY.md 인덱스 덮개
    #
    # 5번이 feedback/·reference/ 의 고아만 세고 projects/ 는 안 봤다. 그 틈에서 사고가 났다 —
    # 2026-09-12, 인덱스를 간결한 목록으로 개편하며 **6개 프로젝트가 조용히 빠졌다**
    # (어느 여섯인지는 그날 git 이력에). 메모리 파일은 그대로 있는데 인덱스에서 사라져서, 그 프로젝트를
    # 여는 세션이 자기 메모리를 못 찾는다. 줄 수 예산은 그때 전부 통과였다 —
    # 「무겁지 않은가」는 봤고 「다 있는가」는 안 봤다.
    pdir = root / "projects"
    if pdir.exists() and index.exists():
        itxt2 = index.read_text(encoding="utf-8")
        plinked = set(re.findall(r"\(projects/([^/]+)/[^)]+\)", itxt2))
        # 「- 이름 — 그 폴더의 CLAUDE.md」 = 정본을 프로젝트 폴더로 내린 것. 누락이 아니다
        pbare = {m.strip() for m in re.findall(r"^- ([^\[\n]+?) — 그 폴더의", itxt2, re.M)}
        orphan_dirs, empty_dirs = [], []
        for d in sorted(p.name for p in pdir.iterdir() if p.is_dir()):
            if d in plinked or d in pbare:
                continue
            (empty_dirs if not any((pdir / d).iterdir()) else orphan_dirs).append(d)
        if orphan_dirs:
            problems.append(
                "projects/ 에 있는데 MEMORY.md 인덱스에 없는 프로젝트: "
                + ", ".join(orphan_dirs)
                + " — 그 프로젝트를 여는 세션이 자기 메모리를 못 찾는다"
            )
        if empty_dirs:
            notes.append(f"빈 프로젝트 폴더(파일 0개): {', '.join(empty_dirs)} — 지울지 판단할 것")

    print()
    for n in notes:
        print(f"  NOTE  {n}")
    for pr in problems:
        print(f"  FIX   {pr}")
    print(f"\n{len(problems)} to fix, {len(notes)} notes")
    return 1 if problems else 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=str(REPO / "agent"))
    ap.add_argument("--drain", action="store_true",
                    help="예산 초과일 때 «무엇을» 뺄지 후보를 낸다")
    a = ap.parse_args()
    rc = main(a.root)
    if a.drain:
        print()
        drain_report(Path(a.root) / "feedback", Path(a.root) / "MEMORY.md")
    sys.exit(rc)
