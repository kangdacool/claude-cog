# -*- coding: utf-8 -*-
"""한글(HWPX) COM 렌더 공용 헬퍼 — 자가치유형 보안모듈 등록.

목적: 한글 COM 자동화 시 "파일 접근 허용" 보안창을 영구히 없앤다.
원리: 유효한 FilePathCheckerModule.dll을 등록하면 RegisterModule이 True가 되어
      창이 뜨지 않는다. 이 헬퍼는 등록이 안 돼 있으면(새 PC·레지스트리 초기화 등)
      DLL을 안정 경로로 복사 + 레지스트리 등록까지 **자동 복구**한 뒤 재시도한다.

사용:
    import os, sys; sys.path.insert(0, "<repo>/agent/tools")
    from hwp_render import make_hwp, render_pdf
    render_pdf("문서.hwpx", "출력.pdf")          # 가장 간단
    # 또는
    h = make_hwp(); h.Open(...); h.SaveAs(...); h.Quit()

⚠️ 한글 COM 렌더는 반드시 이 make_hwp()를 경유할 것(직접 RegisterModule 호출 금지).
"""
import os
import shutil

MODULE_NAME = "FilePathCheckerModule"
REG_KEY = r"Software\HNC\HwpAutomation\Modules"
_LOCAL = os.environ.get("LOCALAPPDATA", os.path.expanduser(r"~\AppData\Local"))
STABLE_DLL = os.path.join(_LOCAL, "HwpSecurityModule", "FilePathCheckerModule.dll")
SYNCED_DLL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "FilePathCheckerModule.dll")   # 저장소에 동봉한 백업본


def _find_dll():
    """우선순위: 안정경로 > 동기화 백업 > pyhwpx 번들. 없으면 None."""
    for p in (STABLE_DLL, SYNCED_DLL):
        if os.path.exists(p):
            return p
    try:
        import pyhwpx
        cand = os.path.join(os.path.dirname(pyhwpx.__file__), "FilePathCheckerModule.dll")
        if os.path.exists(cand):
            return cand
    except Exception:
        pass
    return None


def _install():
    """DLL을 안정 경로로 복사 + 레지스트리 등록. 등록한 DLL 경로 반환(실패 시 None)."""
    src = _find_dll()
    if not src:
        return None
    try:
        os.makedirs(os.path.dirname(STABLE_DLL), exist_ok=True)
        if not os.path.exists(STABLE_DLL):
            shutil.copyfile(src, STABLE_DLL)
    except Exception:
        pass
    dll = STABLE_DLL if os.path.exists(STABLE_DLL) else src
    try:  # 다른 기기용 동기화 백업도 확보
        if not os.path.exists(SYNCED_DLL):
            shutil.copyfile(dll, SYNCED_DLL)
    except Exception:
        pass
    try:
        import winreg
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        winreg.SetValueEx(k, MODULE_NAME, 0, winreg.REG_SZ, dll)
        winreg.CloseKey(k)
    except Exception:
        return None
    return dll


class HwpSecurityModuleError(RuntimeError):
    """보안모듈 등록 실패. 경고가 아니라 «예외»인 이유는 아래 make_hwp() 참고."""


def make_hwp(register=True, strict=True):
    """등록까지 마친 HwpObject 반환. RegisterModule True면 접근허용 창이 안 뜬다.

    ⚠️ **등록 실패는 경고가 아니라 기본적으로 예외다(strict=True).** 2026-08-18 실측:
    등록이 안 된 채로 Open()을 부르면 «파일 접근 허용» 모달이 뜨는데, 자동화 세션에서는
    그 창이 **화면에 보이지도 않는다**(창 클래스 HNC_DIALOG, IsWindowVisible=0). 그래서
    Open()이 **영원히 블록**되고, 호출자는 "왜 멈췄는지" 알 방법이 없다 -- 타임아웃도
    예외도 없이 그냥 안 끝난다. 예전 버전은 여기서 print 경고만 하고 객체를 그대로
    돌려줬는데, 그 경고는 (a) 배치 로그에 묻히고 (b) 바로 다음 줄에서 무한 대기가
    시작되므로 사실상 아무도 못 본다. 멈추는 대신 **즉시 죽는 쪽**이 훨씬 낫다.

    같은 날 실증된 흔한 원인: 호출자가 이 헬퍼를 안 쓰고 손으로
    `RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")`를 부른 경우 --
    한컴 «예제 코드»에 나오는 이름이라 흔히 복사되는데, 실제 레지스트리에 등록된 이름은
    `FilePathCheckerModule`이다. 이름이 다르면 등록이 조용히 False가 되고 위 무한 대기로
    이어진다. **직접 RegisterModule을 부르지 말고 이 함수를 경유할 것.**
    """
    import win32com.client as win32
    h = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    if not register:
        return h
    ok = h.RegisterModule("FilePathCheckDLL", MODULE_NAME)
    if not ok:                       # 자가치유: 등록 복구 후 새 인스턴스로 재시도
        if _install():
            try:
                h.Quit()
            except Exception:
                pass
            h = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            ok = h.RegisterModule("FilePathCheckDLL", MODULE_NAME)
    if not ok:
        msg = ("[hwp_render] 보안모듈(%s) 등록 실패 → 이 상태로 Open()을 부르면 "
               "«파일 접근 허용» 창이 보이지 않는 채로 뜨고 무한 대기한다.\n"
               "  복구: pip install pyhwpx  (그 뒤 재호출 시 자동 등록)" % MODULE_NAME)
        if strict:
            try:
                h.Quit()
            except Exception:
                pass
            raise HwpSecurityModuleError(msg)
        print(msg)
    return h


def render_pdf(hwpx_path, pdf_path, fmt=None):
    """한글 문서 → PDF. 창 없이 렌더.

    `fmt`를 안 주면 확장자로 정한다. ⚠️ 형식을 틀리게 주면 한글이 **조용히 빈 문서를
    열고** 백지 1쪽짜리 PDF가 나온다(예외가 안 난다). 예전에는 "HWPX"가 박혀 있어서
    `.hwp`를 넘기면 그렇게 됐다 — 렌더 결과의 쪽수·글자 수를 항상 확인할 것.
    """
    src = os.path.abspath(hwpx_path)
    if fmt is None:
        fmt = "HWP" if src.lower().endswith(".hwp") else "HWPX"
    h = make_hwp()
    h.Open(src, fmt, "forceopen:true")
    h.SaveAs(os.path.abspath(pdf_path), "PDF", "")
    try:
        h.Quit()
    except Exception:
        pass
    return os.path.abspath(pdf_path)


MESSAGEBOX_AUTO_ANSWER = 0x00011011
"""SetMessageBoxMode에 넣으면 한글이 대화상자에 «자동 응답»한다.

이게 왜 중요한가 (2026-08-19 실증): **암호 걸린 문서를 연 직후 모든 COM 호출이
반환하지 않는 현상**의 해법이다. SaveAs도 GetTextFile도 예외 없이 그냥 안 끝나고,
그때 이름 없는 HNC_DIALOG가 떠 있는데 창을 찾아 눌러도 안 풀린다(버튼이 '설정'
하나뿐이거나 아예 없음). 이 값을 Open() «전에» 걸면 통과한다. 0x20이나 미설정은 실패.
"""


def open_with_password(hwp, path, passwords, fmt=None, timeout=60):
    """암호 걸린 .hwp를 연다. 성공하면 True.

    ⚠️ **반드시 메인 스레드에서 호출**하고, `hwp`는 다른 스레드에서 만든 것이 아니어야
    한다 -- 아래 대화상자 조작(pywinauto uia)이 보조 스레드에서는 창을 찾기까지만 되고
    조작이 조용히 실패한다(2026-08-19 실측).

    한컴은 Open()에 암호를 넘기는 API를 «의도적으로» 제공하지 않는다(개발자포럼 공식
    답변). 그래서 «자동화 불가»라고 결론내기 쉬운데, 사실은 대화상자를 대신 채우면 된다:
      ① 보안모듈이 정상 등록돼야 '문서 암호' 창이 «보인다»(make_hwp가 보장).
      ② 입력칸은 WPF라 네이티브 자식이 없다 -> uia 백엔드로 접근.
      ③ **확인 버튼의 invoke()/click은 안 먹는다. Edit에 {ENTER}를 보내야 제출된다.**
      ④ 연 다음을 위해 SetMessageBoxMode(MESSAGEBOX_AUTO_ANSWER)가 필요하다.
    """
    import threading
    import time
    from pywinauto import Desktop
    import win32gui
    import win32process
    import psutil

    fmt = fmt or ("HWP" if str(path).lower().endswith(".hwp") else "HWPX")
    try:
        hwp.SetMessageBoxMode(MESSAGEBOX_AUTO_ANSWER)
    except Exception:
        pass

    state = {}

    def _open():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            state["ok"] = bool(hwp.Open(str(path), fmt, "forceopen:true"))
        except Exception as e:
            state["error"] = e
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=_open, daemon=True)
    t.start()

    def _dialog():
        found = []

        def cb(h, _):
            try:
                _, pid = win32process.GetWindowThreadProcessId(h)
                if psutil.Process(pid).name().lower() != "hwp.exe":
                    return
            except Exception:
                return
            if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h) == "문서 암호":
                found.append(h)
        win32gui.EnumWindows(cb, None)
        return found[0] if found else None

    deadline, attempt = time.monotonic() + timeout, 0
    while t.is_alive() and time.monotonic() < deadline:
        h = _dialog()
        if h is not None and attempt < len(passwords):
            try:
                dlg = Desktop(backend="uia").window(handle=h)
                edits = dlg.descendants(control_type="Edit")
                if edits:
                    edits[0].set_focus()
                    edits[0].set_edit_text(passwords[attempt])
                    edits[0].type_keys("{ENTER}")
                    attempt += 1
                    time.sleep(1.2)
                    continue
            except Exception:
                pass
        time.sleep(0.3)
    t.join(1)
    return bool(state.get("ok"))


def convert_batch(files, timeout=60, on_progress=None):
    """여러 .hwp/.hwpx를 «같은 자리에» .pdf로 일괄 변환. (경로, 성공여부, 사유) 리스트 반환.

    왜 render_pdf()를 for로 도는 것으로 충분하지 않은가 -- 2026-08-18 실측(5,187개
    배치)에서 얻은 세 가지:

    ① **한 파일이 멈추면 배치 전체가 멈춘다.** 암호 걸린 .hwp가 대표적이다(한컴은 COM
       자동화로 암호 문서를 여는 API를 «의도적으로» 제공하지 않는다 -- 개발자포럼 공식
       답변). 그런 파일은 예외도 False도 아니고 그냥 안 끝난다. 그래서 파일마다 워커
       스레드 + 타임아웃으로 감싸고, 넘기면 실패로 기록하고 다음으로 간다.
    ② **멈춘 뒤에는 그 COM 객체를 재사용할 수 없다.** taskkill로 Hwp.exe를 정리하고
       다음 파일은 새 객체로 연다. COM 객체는 만든 스레드에 묶이므로(apartment) 워커가
       자기 것을 만들어야 한다 -- 메인에서 만든 걸 넘기면 'HWPFrame.HwpObject.Open'
       오류가 난다.
    ③ **OneDrive 동기화 폴더의 원본을 직접 열지 않는다.** 임시폴더로 복사해서 그 사본을
       열고 결과만 원본 옆에 돌려놓는다(원본 무변경). 플레이스홀더 상태 파일에서 나던
       불안정을 피하기 위한 것으로, 비용은 복사 한 번뿐이다.

    Args:
        files: 변환할 경로들
        timeout: 파일당 상한(초). 넘으면 실패로 기록하고 계속.
        on_progress: 있으면 (i, n, path)로 매 파일 호출.
    """
    import shutil
    import subprocess
    import tempfile
    import threading

    def _worker(src, dst, result):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            h = make_hwp()          # 등록 실패면 여기서 예외 -> 무한대기 대신 즉시 실패
            try:
                fmt = "HWP" if str(src).lower().endswith(".hwp") else "HWPX"
                if not h.Open(str(src), fmt, "forceopen:true"):
                    result["error"] = "Open() returned False"
                    return
                h.SaveAs(str(dst), "PDF", "")
                h.Clear(1)
                result["ok"] = True
            finally:
                try:
                    h.Quit()
                except Exception:
                    pass
        except Exception as e:
            result["error"] = repr(e)
        finally:
            pythoncom.CoUninitialize()

    files = [os.path.abspath(str(f)) for f in files]
    stage = tempfile.mkdtemp(prefix="hwp_batch_")
    out = []
    try:
        for i, src in enumerate(files, 1):
            if on_progress:
                on_progress(i, len(files), src)
            dst = os.path.splitext(src)[0] + ".pdf"
            s_src = os.path.join(stage, os.path.basename(src))
            s_dst = os.path.splitext(s_src)[0] + ".pdf"
            shutil.copy2(src, s_src)

            res = {}
            t = threading.Thread(target=_worker, args=(s_src, s_dst, res), daemon=True)
            t.start()
            t.join(timeout)
            if t.is_alive():
                out.append((src, False, "timeout(%ds) -- 암호 문서일 수 있음" % timeout))
                subprocess.run(["taskkill", "/IM", "Hwp.exe", "/F"], capture_output=True)
            elif res.get("ok"):
                shutil.move(s_dst, dst)
                out.append((src, True, ""))
            else:
                out.append((src, False, res.get("error", "unknown")))
            for p in (s_src, s_dst):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
    finally:
        shutil.rmtree(stage, ignore_errors=True)
    return out


def _lineseg_count(path):
    """hwpx의 조판 캐시(linesegarray/lineseg) 개수. .hwp는 셀 수 없어 None."""
    if not path.lower().endswith(".hwpx"):
        return None
    import re
    import zipfile
    n = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.fullmatch(r"Contents/section\d+\.xml", name):
                n += z.read(name).count(b"<hp:lineseg ")
    return n


def save_paginated(src, out=None, also=(), tries=3):
    """한글로 열어 **조판까지 끝낸 뒤** 저장한다. 발송본을 만들 때 쓴다.

    왜 필요한가 — 스크립트로 구조를 편집하면 `linesegarray`(줄 배치 캐시)를 지운다
    (안 지우면 자간·줄간격이 깨진다). 그 상태로 내보내면 받는 쪽에서 파일을 뜯었을 때
    「사람이 한글에서 저장한 게 아니다」가 바로 보인다. 기능상 문제는 없지만(한글이 열 때
    다시 계산한다) 남에게 보내는 문서라면 지우고 갈 이유가 없다.

    실측으로 알아낸 세 가지 (2026-08-03, 가습기살균제 취합본):
      ① 한글은 **있는 캐시는 보존하지만 없으면 만들지 않는다.** 그냥 열고 저장하면 0 → 0.
      ② **조판을 끝까지 시킨 뒤** 저장해야 생긴다 (0 → 1,696). 쪽수를 읽는 것이 가장
         확실하다 — 페이지를 다 세려면 조판이 끝나야 하므로.
      ③ **비동기라 한 번에 안 될 때가 있다.** 같은 코드가 1,696을 냈다가 0을 냈다.
         그래서 결과를 «세어» 확인하고 재시도한다.
      ④ 열려 있는 **그 경로로 덮어쓰면 반영되지 않는다** → 새 경로에 저장한 뒤 바꿔치기.

    Args:
        src:  원본 .hwpx
        out:  저장 경로(기본: src를 제자리에서 교체)
        also: 함께 낼 다른 형식들 — [(경로, "HWP"), …]. 확장자만 주면 형식을 추론한다.
    Returns:
        (out 경로, 조판 캐시 개수)
    """
    import shutil
    import tempfile
    src = os.path.abspath(src)
    out = os.path.abspath(out or src)
    extra = [(os.path.abspath(p), f or ("HWP" if p.lower().endswith(".hwp") else "HWPX"))
             for p, *rest in ((a if isinstance(a, (list, tuple)) else (a,)) for a in also)
             for f in [rest[0] if rest else None]]
    tmp = os.path.join(tempfile.gettempdir(), "_hwp_paginated.hwpx")
    pages = None
    for attempt in range(1, tries + 1):
        os.system("taskkill /F /IM Hwp.exe >nul 2>&1")
        if os.path.exists(tmp):
            os.remove(tmp)
        h = make_hwp()
        h.Open(src, "HWP" if src.lower().endswith(".hwp") else "HWPX", "forceopen:true")
        try:
            h.HAction.Run("MoveDocEnd")
        except Exception:
            pass
        try:                                   # 쪽수를 읽으면 조판이 끝난다
            pages = h.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.PageCount
        except Exception:
            pages = None
        h.SaveAs(tmp, "HWPX", "")
        for path, fmt in extra:
            h.SaveAs(path, fmt, "")
        try:
            h.Quit()
        except Exception:
            pass
        if (_lineseg_count(tmp) or 0) > 0:
            break
    n = _lineseg_count(tmp) or 0
    shutil.move(tmp, out)
    if n == 0:
        print("[hwp_render] 경고: 조판 캐시가 여전히 0이다 (%d회 시도). 한글에서 직접 "
              "열어 저장해야 할 수 있다." % tries)
    return out, n


if __name__ == "__main__":
    import sys
    h = make_hwp()
    print("RegisterModule OK, dll=", _find_dll())
    h.Quit()
