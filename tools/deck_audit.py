#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deck_audit.py — PPTX 구조 정합성 감사기

validate.py가 "파일이 열리는가"를 본다면, 이 스크립트는
"덱이 말이 되는가"를 본다. 슬라이드 XML만 고쳐서는 절대 드러나지 않는,
슬라이드끼리 서로를 가리키는 참조들을 검사한다.

검사 항목
  [SEC]  구역(p14:sectionLst) — 모든 슬라이드가 정확히 한 구역에 속하는가,
         유령 ID를 가리키지 않는가, 구역이 슬라이드 순서상 연속적인가
  [ZOOM] 슬라이드 확대/축소 — 가리키는 sldId가 실재하는가,
         cId가 대상 슬라이드의 실제 p14:creationId와 일치하는가,
         미리보기 이미지 rId가 해석되는가
  [CID]  p14:creationId 중복 — 슬라이드 복제 후 고유화를 빠뜨리면
         줌이 엉뚱한 슬라이드로 점프한다 (조용히 실패하는 유형)
  [TOC]  목차 ↔ 구분 슬라이드 제목 verbatim 일치
  [ORPH] 참조되지 않는 media / 깨진 rId

사용법
  python deck_audit.py deck.pptx
  python deck_audit.py deck.pptx --toc 2 --dividers 3,7,17,25,29,33
      --toc        목차 슬라이드 번호(1-based 화면 순서)
      --dividers   목차와 대조할 구분 슬라이드 번호들
      (--toc만 주고 --dividers를 생략하면 줌이 가리키는 슬라이드를 자동 사용)

종료 코드: 문제 발견 시 1, 이상 없으면 0
"""
import argparse
import re
import sys
import zipfile
from collections import Counter, defaultdict

# 출력에 em-dash·«» 등이 섞이는데 Windows 콘솔 기본 코덱(cp949)이 이를 못 찍어
# UnicodeEncodeError로 스크립트가 통째로 죽는다. 2026-08-10에 audit.py로 처음
# 일괄 실행했을 때 드러났다 — 그전엔 단독 실행(대개 UTF-8 터미널)만 해서 안 보였다.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NS_SLIDE = 'slides/'


# ────────────────────────────────────────────────── 패키지 로딩

class Deck:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.pres = self._read('ppt/presentation.xml')
        self.pres_rels = self._read('ppt/_rels/presentation.xml.rels')

        # rId → slideN.xml
        self.rid2file = dict(
            (r, f) for r, f in
            re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', self.pres_rels)
        )
        # 화면 순서: [(sldId, slideN.xml)]
        self.order = [
            (sid, self.rid2file[rid])
            for sid, rid in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', self.pres)
            if rid in self.rid2file
        ]
        self.id2file = {sid: f for sid, f in self.order}
        self.file2pos = {f: i + 1 for i, (_, f) in enumerate(self.order)}
        self.pos2file = {i + 1: f for i, (_, f) in enumerate(self.order)}
        self.slide_xml = {f: self._read('ppt/slides/' + f) for _, f in self.order}
        self.slide_rels = {}
        for _, f in self.order:
            try:
                self.slide_rels[f] = self._read(f'ppt/slides/_rels/{f}.rels')
            except KeyError:
                self.slide_rels[f] = ''

    def _read(self, name):
        return self.z.read(name).decode('utf-8', 'replace')

    def names(self):
        return set(self.z.namelist())

    # 슬라이드의 p14:creationId (줌이 대상을 식별하는 열쇠)
    def creation_id(self, f):
        m = re.search(r'<p14:creationId[^>]*val="(\d+)"', self.slide_xml[f])
        return m.group(1) if m else None

    # 제목 placeholder 텍스트
    def title_of(self, f):
        x = self.slide_xml[f]
        for m in re.finditer(r'<p:sp>(?:(?!</p:sp>).)*?</p:sp>', x, re.S):
            blk = m.group(0)
            if '<p:ph type="title"' in blk:
                return ''.join(re.findall(r'<a:t>([^<]*)</a:t>', blk)).strip()
        return None

    # 본문 placeholder 문단들 (목차 항목 추출용)
    def body_paras(self, f):
        x = self.slide_xml[f]
        for m in re.finditer(r'<p:sp>(?:(?!</p:sp>).)*?</p:sp>', x, re.S):
            blk = m.group(0)
            if re.search(r'<p:ph (?:type="body" )?idx="1"/>', blk):
                out = []
                for p in re.findall(r'<a:p>.*?</a:p>', blk, re.S):
                    t = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', p)).strip()
                    if t:
                        out.append(t)
                return out
        return []


# ────────────────────────────────────────────────── 개별 검사

def check_sections(d, issues, notes):
    m = re.search(r'<p14:sectionLst.*?</p14:sectionLst>', d.pres, re.S)
    if not m:
        notes.append('[SEC] 구역이 정의되어 있지 않음 (구역 기능 미사용 덱이면 정상)')
        return
    secs = re.findall(
        r'<p14:section name="([^"]*)"[^>]*><p14:sldIdLst>(.*?)</p14:sldIdLst>', m.group(0), re.S)
    seen = Counter()
    flat = []
    for name, body in secs:
        ids = re.findall(r'<p14:sldId id="(\d+)"/>', body)
        pages = []
        for sid in ids:
            seen[sid] += 1
            f = d.id2file.get(sid)
            if f is None:
                issues.append(f'[SEC] 구역 "{name}"이 존재하지 않는 슬라이드 id={sid}를 참조 (유령 참조)')
            else:
                pages.append(d.file2pos[f])
        flat.append((name, pages))
        if pages:
            span = list(range(min(pages), max(pages) + 1))
            if sorted(pages) != span:
                issues.append(f'[SEC] 구역 "{name}"의 슬라이드가 연속적이지 않음: {sorted(pages)}')

    for sid, f in d.order:
        if seen[sid] == 0:
            issues.append(f'[SEC] {d.file2pos[f]}쪽({f})이 어떤 구역에도 속하지 않음')
        elif seen[sid] > 1:
            issues.append(f'[SEC] {d.file2pos[f]}쪽({f})이 {seen[sid]}개 구역에 중복 소속')

    prev_end = 0
    for name, pages in flat:
        if pages and min(pages) < prev_end:
            issues.append(f'[SEC] 구역 "{name}"이 앞 구역과 순서가 뒤엉킴')
        if pages:
            prev_end = max(pages)
    notes.append('[SEC] 구역 구성: ' + ' | '.join(
        f'{n}: {min(p)}–{max(p)}쪽({len(p)}장)' if p else f'{n}: (빈 구역)' for n, p in flat))


def check_creation_ids(d, issues, notes):
    by_cid = defaultdict(list)
    for _, f in d.order:
        cid = d.creation_id(f)
        if cid is None:
            notes.append(f'[CID] {d.file2pos[f]}쪽에 creationId 없음 (줌 대상으로 쓸 수 없음)')
        else:
            by_cid[cid].append(d.file2pos[f])
    for cid, pages in by_cid.items():
        if len(pages) > 1:
            issues.append(
                f'[CID] creationId {cid} 중복: {pages}쪽 — 슬라이드 복제 후 고유화 누락. '
                f'줌이 엉뚱한 슬라이드로 점프할 수 있음')


def check_zooms(d, issues, notes):
    total = 0
    targets = []
    for _, f in d.order:
        x = d.slide_xml[f]
        rels = d.slide_rels[f]
        rid_ok = set(re.findall(r'Id="(rId\d+)"', rels))
        for zm in re.finditer(
                r'<pslz:sldZmObj sldId="(\d+)" cId="(\d+)">(.*?)</pslz:sldZmObj>', x, re.S):
            total += 1
            sid, cid, body = zm.group(1), zm.group(2), zm.group(3)
            page = d.file2pos[f]
            tgt = d.id2file.get(sid)
            if tgt is None:
                issues.append(f'[ZOOM] {page}쪽의 줌이 존재하지 않는 슬라이드 id={sid}를 가리킴 (끊긴 링크)')
                continue
            targets.append((page, d.file2pos[tgt], tgt))
            real = d.creation_id(tgt)
            if real != cid:
                issues.append(
                    f'[ZOOM] {page}쪽 줌의 cId={cid}가 대상 {d.file2pos[tgt]}쪽의 '
                    f'실제 creationId={real}와 불일치 — PowerPoint가 링크를 잃거나 오작동함')
            for rid in re.findall(r'<a:blip r:embed="(rId\d+)"/>', body):
                if rid not in rid_ok:
                    issues.append(f'[ZOOM] {page}쪽 줌 미리보기 {rid}가 rels에 없음')
    # Fallback pic의 rId까지 확인
    for _, f in d.order:
        rels = d.slide_rels[f]
        rid_ok = set(re.findall(r'Id="(rId\d+)"', rels))
        for m in re.finditer(r'<mc:Fallback.*?</mc:Fallback>', d.slide_xml[f], re.S):
            for rid in re.findall(r'<a:blip r:embed="(rId\d+)"/>', m.group(0)):
                if rid not in rid_ok:
                    issues.append(f'[ZOOM] {d.file2pos[f]}쪽 줌 Fallback 이미지 {rid}가 rels에 없음')
    if total:
        notes.append('[ZOOM] 줌 %d개 → %s쪽' % (total, [t[1] for t in targets]))
    return [t[2] for t in targets]


def check_toc(d, toc_page, divider_files, issues, notes):
    f = d.pos2file.get(toc_page)
    if f is None:
        issues.append(f'[TOC] {toc_page}쪽이 존재하지 않음')
        return
    items = d.body_paras(f)
    titles = [d.title_of(t) for t in divider_files]
    if not items:
        notes.append(f'[TOC] {toc_page}쪽 본문에서 항목을 찾지 못함 (수동 확인 필요)')
        return
    if len(items) != len(titles):
        issues.append(f'[TOC] 목차 {len(items)}개 항목 vs 구분 슬라이드 {len(titles)}개 — 개수 불일치')
    for i, (it, ti) in enumerate(zip(items, titles), 1):
        if it != ti:
            issues.append(f'[TOC] {i}번 항목 불일치\n        목차: {it!r}\n        제목: {ti!r}')
    if len(items) == len(titles) and all(a == b for a, b in zip(items, titles)):
        notes.append(f'[TOC] 목차 {len(items)}개 항목이 구분 슬라이드 제목과 정확히 일치')


def check_media(d, issues, notes):
    names = d.names()
    referenced = set()
    for n in names:
        if not n.endswith('.rels'):
            continue
        body = d._read(n)
        base = n.rsplit('/_rels/', 1)[0]
        for tgt in re.findall(r'Target="([^"]+)"', body):
            if tgt.startswith('http') or tgt.startswith('#'):
                continue
            parts = (base + '/' + tgt).split('/')
            stack = []
            for p in parts:
                if p == '..':
                    if stack:
                        stack.pop()
                elif p not in ('.', ''):
                    stack.append(p)
            referenced.add('/'.join(stack))
    for n in sorted(names):
        if n.endswith('/'):  # zip 디렉터리 엔트리는 파일이 아님
            continue
        if n.startswith('ppt/media/') and n not in referenced:
            issues.append(f'[ORPH] 참조되지 않는 미디어: {n}')


# ────────────────────────────────────────────────── 진입점

def main():
    ap = argparse.ArgumentParser(description='PPTX 구조 정합성 감사')
    ap.add_argument('pptx')
    ap.add_argument('--toc', type=int, default=None, help='목차 슬라이드 번호 (화면 순서)')
    ap.add_argument('--dividers', default=None, help='구분 슬라이드 번호들, 쉼표 구분')
    a = ap.parse_args()

    d = Deck(a.pptx)
    issues, notes = [], []

    check_sections(d, issues, notes)
    check_creation_ids(d, issues, notes)
    zoom_targets = check_zooms(d, issues, notes)
    check_media(d, issues, notes)

    if a.toc:
        if a.dividers:
            files = [d.pos2file[int(p)] for p in a.dividers.split(',')]
        else:
            files = zoom_targets
        if files:
            check_toc(d, a.toc, files, issues, notes)
        else:
            notes.append('[TOC] 대조할 구분 슬라이드를 찾지 못함 (--dividers로 지정)')

    print(f'=== {a.pptx} — 슬라이드 {len(d.order)}장 ===')
    for n in notes:
        print('  ' + n)
    if issues:
        print(f'\n문제 {len(issues)}건:')
        for i in issues:
            print('  ✗ ' + i)
        return 1
    print('\n구조 검사 통과 — 구역·줌·목차·미디어 참조 모두 정합')
    return 0


if __name__ == '__main__':
    sys.exit(main())
