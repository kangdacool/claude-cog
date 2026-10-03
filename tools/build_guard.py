# -*- coding: utf-8 -*-
"""재생성 스크립트가 **사람이 손으로 고친 산출물**을 덮어쓰지 못하게 막는다.

문서를 스크립트로 만들면 반드시 이런 일이 생긴다: 사람이 Word/한글에서 한 군데를 직접
고치고, 나중에 스크립트를 다시 돌리면 그 수정이 **조용히 사라진다.** 스크립트는 그 수정을
재현할 방법이 없기 때문이다.

"수동편집본 > 자동생성물"은 오래된 규칙이지만 규칙만으로는 안 지켜졌다(2026-07-27 실사고:
사용자가 한글에서 고친 영문 제목을 빌드가 덮어씀). 그래서 코드로 막는다.

쓰는 법 — 빌드 스크립트의 저장 직전·직후에 한 줄씩:

    from build_guard import guard, stamp
    ...
    guard(OUT)                 # 손으로 고쳤으면 여기서 SystemExit
    save_document(OUT)
    stamp(OUT)                 # 이번 산출물의 지문을 남긴다

⭐ `guard()` 는 산출물이 **지금 Office 에 열려 있으면** 그것부터 막는다(`open_in_office` — 2026-10-01
   OneDrive 공동편집이 열린 덱과 새 빌드를 합쳐 도형을 두 벌 넣은 사고). 손으로 확인: `build_guard.py open <산출물>`.

지문은 `<OUT>.build-md5`. 손으로 고친 뒤 그 수정을 스크립트에 반영했다면 그 파일을 지우고
다시 돌리면 된다(또는 `guard(OUT, force=True)`).

⚠️ `guard()` 는 «다음 빌드»에서만 운다. 빌드 «직후»에 파일이 바뀌는 사고는 그때까지
   드러나지 않는다 -- 그 사이에 감사·렌더·전달이 다 지나간다. 그래서 `verify()` 가 있다:

    python build_guard.py verify <산출물>      # 아직 그 빌드의 산출물인가?

**언제 부르나 — pptx/docx 를 COM 으로 렌더한 «직후»에 반드시.** 렌더는 PowerPoint/Word 를
띄우고, 그 앱이 파일을 저장해 버리는 일이 실제로 있다. 2026-08-26 teacher 포스터에서 하루에
두 번 났다: 렌더 뒤 PowerPoint 가 저장하며 표·그림·푸터가 **두 벌씩 복제**됐고(도형 48개,
같은 좌표 중복 11쌍), PDF·PNG 는 깨끗했기에 «렌더를 보는 검사»로는 안 잡혔다. 그 pptx 를
그대로 보냈으면 인쇄물이 겹쳐 나온다.
"""
import hashlib
import os


def _md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def stamp_path(out_path):
    return out_path + '.build-md5'


def stamp(out_path):
    """방금 만든 산출물의 지문을 기록한다. 저장 직후에 부른다."""
    with open(stamp_path(out_path), 'w', encoding='utf-8') as f:
        f.write(_md5(out_path))
    return stamp_path(out_path)


def open_in_office(out_path):
    """산출물이 지금 PowerPoint·Word·Excel 에 열려 있는가. 열려 있다는 근거(문자열) 목록을 돌려준다.

    왜 (2026-10-01 한 북리딩 덱, 하루 세 번): 연구자가 덱을 PowerPoint 로 열어 둔 채 재빌드했더니
    OneDrive 공동편집(AutoSave)이 열린 문서와 새 파일을 «합쳐» 저장해 도형이 두 벌씩 들어갔다.
    빌드도 감사도 렌더도 다 통과했고, 연구자가 「왜 글씨가 2번 적혀있냐」로 발견했다.
    ⚠ 공동편집 중인 문서는 로컬 파일을 잠그지 않을 수 있다 — 그래서 파일 잠금만 보지 않고
       실행 중인 Office 에 «같은 이름의 문서가 열려 있는가»를 COM 으로 묻는다(새 Office 를 띄우지 않는다).
    """
    hits = []
    if not os.path.exists(out_path):
        return hits
    name = os.path.basename(out_path).lower()
    try:
        import win32com.client
        for prog, coll in (("PowerPoint.Application", "Presentations"),
                           ("Word.Application", "Documents"),
                           ("Excel.Application", "Workbooks")):
            try:
                app = win32com.client.GetActiveObject(prog)
            except Exception:
                continue                       # 그 앱이 안 떠 있다
            try:
                for d in getattr(app, coll):
                    if str(d.Name).lower() == name or str(d.Name).lower() + os.path.splitext(name)[1] == name:
                        hits.append("%s 에 열려 있음: %s" % (prog.split(".")[0], d.FullName))
            except Exception:
                pass
    except ImportError:
        pass
    try:                                       # 로컬 잠금(공동편집이 아닌 보통 열기)
        with open(out_path, "r+b"):
            pass
    except PermissionError:
        hits.append("파일이 잠겨 있음(다른 프로그램이 쓰는 중)")
    return hits


def refuse_if_open(out_path):
    """열려 있으면 SystemExit — 빌드를 시작하기 «전»에 부른다. 사람의 Office 를 닫지 않는다."""
    hits = open_in_office(out_path)
    if hits:
        raise SystemExit(
            '\n[중단] 산출물이 열려 있습니다 — 빌드하지 않았습니다.\n        %s\n        %s\n\n'
            '        열린 채 덮어쓰면 OneDrive 공동편집이 두 판을 합쳐 도형이 두 벌씩 들어갑니다.\n'
            '        그 파일을 닫아 달라고 사람에게 말하고, 닫힌 뒤 다시 실행하세요.\n'
            '        (사람의 PowerPoint/Word 를 대신 닫지 말 것)\n'
            % (out_path, "\n        ".join(hits)))
    base = os.path.basename(out_path)
    d = os.path.dirname(os.path.abspath(out_path))
    owners = [f for f in os.listdir(d) if f.startswith("~$") and (f[2:] == base or base.endswith(f[2:]))]
    if owners:                                 # 남은 소유자 파일은 강제 종료의 흔적일 수도 있어 막지 않는다
        print("  ⚠ Office 소유자 파일이 남아 있음(%s) — 열려 있지 않다면 지난 강제 종료의 흔적" % ", ".join(owners))


def guard(out_path, force=False):
    """산출물이 지난 빌드 이후 바뀌었거나 지금 Office 에 열려 있으면 SystemExit로 중단한다.

    지문이 없으면(첫 빌드) 통과시킨다 — 없는 지문을 근거로 막으면 첫 실행이 불가능하다.
    열림 검사는 force 와 무관하게 한다 — 손편집 무시(force)와 열린 문서 덮어쓰기는 다른 일이다.
    """
    refuse_if_open(out_path)
    if force:
        return
    sp = stamp_path(out_path)
    if not os.path.exists(out_path) or not os.path.exists(sp):
        return
    now = _md5(out_path)
    with open(sp, encoding='utf-8') as f:
        was = f.read().strip()
    if now == was:
        return
    raise SystemExit(
        '\n[중단] 산출물이 지난 빌드 이후 수정되었습니다 — 덮어쓰지 않았습니다.\n'
        '        %s\n\n'
        '        누군가 이 파일을 직접 고쳤습니다. 스크립트는 그 수정을 재현할 수 없으므로\n'
        '        지금 빌드하면 사라집니다. 수정 내용을 빌드 스크립트에 반영한 뒤\n'
        '        %s 를 지우고 다시 실행하세요.\n' % (out_path, sp))


def verify(out_path):
    """산출물이 아직 «그 빌드가 만든 것»인지 확인한다. (ok, 지금md5, 기록md5) 를 돌려준다.

    ok 가 False 면 빌드 뒤에 누군가/무언가가 파일을 고쳤다는 뜻이다. 렌더용 COM 앱이
    저장해 버린 경우가 대표적이며, 그때 파일은 «렌더한 것과도 다르다» -- 렌더 PNG 를
    아무리 들여다봐도 안 보인다. 재빌드가 정답이다.

    지문이 없으면 (None, md5, None) -- 판정하지 않는다(첫 빌드일 수 있다).
    """
    if not os.path.exists(out_path):
        raise SystemExit('파일이 없다: ' + out_path)
    now = _md5(out_path)
    sp = stamp_path(out_path)
    if not os.path.exists(sp):
        return None, now, None
    with open(sp, encoding='utf-8') as f:
        was = f.read().strip()
    return (now == was), now, was


if __name__ == '__main__':
    # Windows 콘솔은 cp949 라 한글·긴줄표(—)·기호 출력에서 UnicodeEncodeError 로 죽는다.
    # 결과를 다 만들어 놓고 «찍는 순간» 죽으므로, 부르는 쪽에는 도구가 고장난 것처럼 보인다.
    # ⚠ 모듈 최상단이 아니라 여기 두는 이유: 이 파일이 import 되기도 하면 최상단
    #    reconfigure 가 «호출자»의 인코딩을 바꾼다. 스크립트로 실행할 때만 돌게 한다.
    try:
        import sys as _s
        _s.stdout.reconfigure(encoding='utf-8', errors='replace')
        _s.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass
    import sys
    if len(sys.argv) != 3 or sys.argv[1] not in ('guard', 'stamp', 'verify', 'open'):
        raise SystemExit('사용: python build_guard.py {guard|stamp|verify|open} <산출물경로>')
    if sys.argv[1] == 'open':
        hits = open_in_office(sys.argv[2])
        print('\n'.join(hits) if hits else '열려 있지 않음: ' + sys.argv[2])
        raise SystemExit(1 if hits else 0)
    if sys.argv[1] == 'stamp':
        print('지문 기록:', stamp(sys.argv[2]))
    elif sys.argv[1] == 'verify':
        ok, now, was = verify(sys.argv[2])
        if ok is None:
            print('지문 없음 — 판정하지 않음:', sys.argv[2])
        elif ok:
            print('OK 빌드 산출물 그대로:', sys.argv[2])
        else:
            raise SystemExit(
                '\n[경고] 빌드 «이후» 파일이 바뀌었습니다.\n'
                '        %s\n        지금 %s\n        기록 %s\n\n'
                '        렌더용 PowerPoint/Word 가 저장했을 가능성이 큽니다. 이 파일은\n'
                '        방금 렌더한 PNG/PDF 와도 다르므로 «렌더를 보는 검사»로는 안 잡힙니다.\n'
                '        지문 파일을 지우고 재빌드하세요.\n' % (sys.argv[2], now, was))
    else:
        guard(sys.argv[2])
        print('변경 없음 — 덮어써도 안전:', sys.argv[2])
