##################################
#####  precommit_scan  #####
##################################
# 커밋 «전에» 소스에 개인정보·비밀번호가 섞였는지 훑는다. (랩 공용 -- 어느 저장소에서든)
#
# .gitignore가 data/·output/·*.pdf를 막고 있어도 그것으로 끝이 아니다. 위험은 두 군데
# 더 있다.
#   ① 설정 파일에 하드코딩된 «문서 비밀번호»
#   ② 코드 주석에 인용한 «환자 이름·원문 문장» -- 디버깅하며 실제 사례를 적어 두기 쉽다
# 한 번 커밋되면 이력에서 지우기가 매우 어려우므로, 들어가기 전에 잡는 편이 훨씬 싸다.
#
#   python tools/precommit_scan.py
#
# ⚠️ 왜 검사가 «좁은가» — 처음엔 「암호」가 든 줄을 전부 잡아 산문(`print("...암호 입력 불가")`)
# 까지 걸려 **항상 종료코드 1**이었다. 통과가 불가능한 게이트는 신호가 0이다. 그래서 자격증명
# «이름을 가진 변수에 문자열을 대입하는 형태»만 보도록 좁히고, 일부러 심은 양성·음성으로
# 자기검사해 정밀도를 확인했다. 넓히고 싶어지면 이 사고부터 떠올릴 것.
# 근본 해결은 검사가 아니다 — 암호를 소스가 아니라 gitignore된 파일에서 읽게 바꾸는 것이다.

import sys, re, subprocess, importlib.util
from pathlib import Path

# ⚠️ 이 스캐너는 «찾았을 때» 한국어와 ⚠ 기호를 stdout에 쓴다. Windows cp949 콘솔에서 그대로
# 쓰면 UnicodeEncodeError로 **적발한 바로 그 순간에 죽는다** — 2026-08-20 실측으로 확인했다.
# 죽은 게이트는 «조용히 통과»와 구별되지 않으므로, 개인정보를 지키라고 만든 도구가 정확히
# 지켜야 할 때 없는 것이 된다. hwpx_guard·deliverable_guard와 같은 가드를 맨 앞에 둔다.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

import os
# 검사 대상은 «현재 작업 디렉토리의 git 저장소»다. 프로젝트에 종속되지 않는다.
ROOT = Path(os.environ.get("PRECOMMIT_ROOT", os.getcwd())).resolve()

TEXT_EXT = {".py", ".csv", ".md", ".txt", ".json", ".yml", ".yaml", ".toml", ".cfg"}

# 환자 이름은 «문서 파일명 자리»에 나타난다. 한글 낱말 전체를 이름 후보로 보면
# 오탐이 수백 건이 되어 스캐너를 못 쓰게 되므로, 이름이 실제로 놓이는 자리만 본다.
#   2022 초진기록지-홍길동.pdf   /   (정신)2024 경과기록지- 홍길동.pdf
#   홍길동 561120 기록지_보안.pdf
NAME_CONTEXT = [
    re.compile(r"(?:초진|경과)\s*기록지\s*[-–]\s*([가-힣]{2,4})"),
    re.compile(r"([가-힣]{2,4})\s+\d{6}\s*기록지"),
    re.compile(r"기록지[^\s\"']{0,12}[-–]\s*([가-힣]{2,4})"),
]
# 서식 자체를 가리키는 말은 이름이 아니다.
NOT_NAME = {"이름", "홍길동", "수정", "실습", "선생님", "보안", "원스탑", "정신", "칸없음"}


def scan_file(path: Path):
    hits = []
    try:
        text = path.read_text(encoding="utf-8")
    except Exception:
        return hits
    is_self = path.name == Path(__file__).name
    for n, line in enumerate(text.splitlines(), 1):
        s = line.strip()
        if not s:
            continue
        # 스캐너 자신의 패턴 정의·예시는 검사 대상이 아니다(영구 오탐).
        if is_self:
            continue
        # ① 비밀번호가 «값으로» 박혀 있는가.
        # ⚠️ "암호"라는 낱말이 든 줄을 전부 잡으면 산문(print 문구·주석)까지 걸려 이 검사가
        # 늘 실패로 끝난다. 늘 실패하는 검사는 아무도 안 본다. 그래서 «자격증명 이름을 가진
        # 변수에 문자열 리터럴을 대입하는 형태»만 본다.
        # «라벨»은 자격증명이 아니다. PASSWORD_DIALOG_TITLE = "문서 암호" 는 화면에 띄울 글귀지
        # 비밀번호가 아니다(2026-08-20 실측 오탐 1건, 이 저장소). 이름이 화면 문구·필드명을
        # 가리키는 접미사로 끝나면 뺀다 -- 값이 아니라 «무엇을 부르는 말인가»로 가른다.
        m_cred = re.search(r"\b(\w*(?:PASSWORD|PASSWD|PWD|SECRET|TOKEN|API_KEY)\w*)\s*=\s*"
                           r"[\[\(]?\s*[\"'][^\"']{3,}[\"']", line, re.IGNORECASE)
        if m_cred and not re.search(
                r"_(?:TITLE|LABEL|PROMPT|MSG|MESSAGE|CAPTION|FIELD|HINT|PLACEHOLDER"
                r"|DIALOG|HEADER|PATTERN|REGEX|ENV|VAR|FILE|PATH|KEYNAME)$",
                m_cred.group(1), re.IGNORECASE):
            hits.append((n, "비밀번호 하드코딩", s[:110]))
        # ② 환자 이름 (문서 파일명 자리에 나타난다)
        for pat in NAME_CONTEXT:
            for m in pat.finditer(line):
                if m.group(1) in NOT_NAME:
                    continue
                hits.append((n, "환자 이름", s[:110]))
                break
    # NAME_CONTEXT 세 패턴이 «같은 줄»을 함께 잡는 일이 흔하다(기록지-이름 형태는 2~3개가 동시에
    # 맞는다). 같은 줄·같은 종류를 여러 번 세면 건수가 부풀어 심각도를 오독하게 된다 -- 한 번만.
    seen, uniq = set(), []
    for h in hits:
        if h not in seen:
            seen.add(h); uniq.append(h)
    return uniq


def main():
    r = subprocess.run(["git", "ls-files", "-co", "--exclude-standard"],
                       cwd=ROOT, capture_output=True, text=True, encoding="utf-8")
    files = [ROOT / f for f in r.stdout.splitlines() if f.strip()]
    files = [f for f in files if f.suffix.lower() in TEXT_EXT and f.exists()]
    print(f"커밋 대상 텍스트 파일 {len(files)}개 검사\n")

    total = 0
    for f in sorted(files):
        hits = scan_file(f)
        if not hits:
            continue
        rel = f.relative_to(ROOT)
        print(f"[{rel}]  {len(hits)}건")
        for n, kind, s in hits[:12]:
            print(f"   {n:>5}  {kind:14s} {s}")
        if len(hits) > 12:
            print(f"   ... 외 {len(hits)-12}건")
        print()
        total += len(hits)

    print(f"총 {total}건")
    if total:
        print("\n⚠️ 커밋 전에 처리하십시오. 한 번 들어가면 이력에서 제거하기 어렵습니다.")
    return 1 if total else 0


if __name__ == "__main__":
    raise SystemExit(main())
