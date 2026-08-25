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

지문은 `<OUT>.build-md5`. 손으로 고친 뒤 그 수정을 스크립트에 반영했다면 그 파일을 지우고
다시 돌리면 된다(또는 `guard(OUT, force=True)`).
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


def guard(out_path, force=False):
    """산출물이 지난 빌드 이후 바뀌었으면 SystemExit로 중단한다.

    지문이 없으면(첫 빌드) 통과시킨다 — 없는 지문을 근거로 막으면 첫 실행이 불가능하다.
    """
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


if __name__ == '__main__':
    import sys
    if len(sys.argv) != 3 or sys.argv[1] not in ('guard', 'stamp'):
        raise SystemExit('사용: python build_guard.py {guard|stamp} <산출물경로>')
    if sys.argv[1] == 'stamp':
        print('지문 기록:', stamp(sys.argv[2]))
    else:
        guard(sys.argv[2])
        print('변경 없음 — 덮어써도 안전:', sys.argv[2])
