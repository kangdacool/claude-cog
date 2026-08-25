---
name: hwpx
description: 한글(HWPX) 파일을 Python/lxml로 읽기·편집·생성하는 전문 에이전트. hwpx 파일의 텍스트 추출, 표 편집, 이미지 삽입, 재압축 등 모든 hwpx 작업에 사용.
---

# HWPX 편집 가이드 (v5)

> 한글(HWPX) 파일을 Python/lxml로 안전하게 편집하기 위한 실전 가이드.
> 모든 항목은 **한글에서 실제로 열어 렌더링까지 확인(라운드트립)된 것**만 "검증됨"으로 표기.
> v4 → v5 변경: ★재압축 작업본 vs 레퍼런스 치명적 이슈, 셀 높이 리셋, 캡션 생성, treatAsChar(표·그림), 세로병합 rowSpan, 헤더 음영, LibreOffice 미지원·diff 검증 추가.

---

## ⚠️ 흔한 실패 7가지 (먼저 읽을 것)

1. **★레퍼런스 원본을 직접 재압축 → 한글에서 안 열림.** 한글이 저장한 원본 hwpx를 Python `zipfile`로 재압축하면 한글이 거부한다(원인: 원본 일부 엔트리의 `flag_bits=4`를 Python이 보존 못함). **반드시 "작업본"(한 번 재압축돼 열림이 확인된 파일) 기반으로 빌드**할 것. → §2
2. **셀 클론 시 `cellSz`의 height를 안 바꿈 → 칸이 비정상적으로 높아짐.** width만 바꾸면 원본의 큰 height가 남는다. 본문=282, 헤더=0. → §5
3. **열 수를 바꾸는 표 클론에서 `cellSz width`·`sz`를 재계산 안 함 → 기하구조 깨져 거부.** → §5
4. **미주/캡션의 `<hp:ctrl>` 래퍼 누락 → 한글이 인식 못함.** → §4, §5
5. **`header.xml`에 charPr/paraPr 추가 후 `itemCnt` 갱신 안 함 → 파일 거부.** → §7
6. **텍스트를 `.text`로만 추출 → `<hp:lineBreak/>` tail·인접 셀 누락/오인.** `itertext()` 사용. → §1
7. **띄어쓰기 일괄 정규식 치환 → 제N절·표N·코드(N17) 등 의도된 표기 파괴.** → §8
8. **★`linesegarray`를 안 지우거나 박아 넣음 → 자간·줄간격이 완전히 망가짐.** 워드 만들듯 위치를 박아 생성·편집할 때 깨지는 주원인. 편집이든 신규 생성이든 절대 넣지 말 것. → §3

---

## §0. 파일 구조

HWPX = **zip + XML(HWPML)**. 주요 엔트리:

```
mimetype                    # 첫 엔트리, STORED, 내용 'application/hwp+zip'
version.xml
Contents/header.xml         # 글꼴·charPr·paraPr·borderFill 정의
Contents/section0.xml ~ N   # 본문(여러 섹션일 수 있음)
Contents/content.hpf        # manifest(이미지 등록 등)
BinData/                    # 임베드 이미지 등
META-INF/manifest.xml       # (보통 비어 있음)
```

핵심 네임스페이스:
- `hp` = `http://www.hancom.co.kr/hwpml/2011/paragraph` (단락·표·런·텍스트)
- `hc` = `http://www.hancom.co.kr/hwpml/2011/core` (인라인 이미지 `<hc:img>`)
- `hh` = header.xml의 charPr/paraPr/fontfaces
- `opf` = content.hpf manifest

---

## §1. 파싱 (읽기)

```python
from lxml import etree
import zipfile
P='{http://www.hancom.co.kr/hwpml/2011/paragraph}'
z=zipfile.ZipFile('file.hwpx')
root=etree.fromstring(z.read('Contents/section0.xml'))
```

- **본문은 section0~N 전부 확인.** section0만 보고 "내용 없음" 오판한 사례 있음.
- **텍스트 추출은 `itertext()`** — `<hp:t>.text`만 보면 `<hp:lineBreak/>` 뒤 `tail`을 놓침.
- 제목 등 수정 시 `t.text`뿐 아니라 `el.tail`도 확인(실제로 제목이 tail에 있던 사례).
- **그림 존재는 `<hp:pic>` 직접 검색**으로 확인(텍스트 추출로 "그림 없음" 단정 금지).
- 표 셀 텍스트를 itertext로 한 번에 뽑으면 인접 셀이 붙어 보임(별개 `<hp:tc>`라 거짓양성). 셀별로 순회.
- `styleIDRef`/`paraPrIDRef`/`charPrIDRef`/`borderFillIDRef`는 **파일마다 다르므로 실제 파일에서 읽어** 쓸 것.

본문 단락 텍스트(미주 제외) 헬퍼:
```python
def own(p):
    return ''.join(''.join(t.itertext()) for t in p.findall(f'.//{P}t')
                   if not any(a.tag==f'{P}endNote' for a in t.iterancestors()))
```

---

## §2. 재압축 (저장) — ★가장 중요

```python
def repack(src_hwpx, changed: dict, out_hwpx):
    # changed = {'Contents/section0.xml': new_bytes, ...}
    zin=zipfile.ZipFile(src_hwpx); zout=zipfile.ZipFile(out_hwpx,'w')
    for it in zin.infolist():
        data=changed.get(it.filename) or zin.read(it.filename)
        zi=zipfile.ZipInfo(it.filename, date_time=it.date_time)
        zi.compress_type=it.compress_type      # 엔트리별 보존
        zi.external_attr=it.external_attr
        zout.writestr(zi, data)
    zin.close(); zout.close()
```

규칙:
- **엔트리 순서·엔트리별 `compress_type` 보존.**
- `mimetype`은 **첫 엔트리·STORED(compress_type=0)**, 내용 `application/hwp+zip`.
- XML 선언 `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` 유지(직렬화 후 앞에 붙임).
- 안전 패턴 = 원본 엔트리 전부 복사 + 바뀐 `Contents/sectionN.xml`만 교체.
- 신규 BinData(이미지 등)는 원본에 없던 엔트리 → **별도 `add`(ZIP_DEFLATED)**.

### ★치명적: 작업본 vs 레퍼런스 (검증됨)

- **한글이 저장한 "레퍼런스" 원본을 Python `zipfile`로 재압축하면 한글에서 안 열린다.**
  원인: 원본 일부 엔트리의 `flag_bits=4`(압축 옵션 비트)를 Python이 보존하지 못하고 재deflate하기 때문.
- **"작업본"**(이미 한 번 Python으로 재압축돼 열림이 확인된 파일 = `flag_bits` 전부 0)은 재압축해도 정상 개봉.
- **결론: 보고서·표 등 모든 빌드는 작업본 기반으로.** 레퍼런스 직접 재압축은 피한다.
  레퍼런스의 구조(표·캡션 등)가 필요하면 **그 요소만 작업본으로 옮겨** 사용.
- **진단법**: "작업본 기반 파일은 열리는데 레퍼런스 기반만 안 열린다" → 이 이슈.

```python
# flag_bits 확인
for i in zipfile.ZipFile('x.hwpx').infolist():
    print(i.filename, i.flag_bits)   # {0}이면 안전, 4가 섞이면 레퍼런스
```

---

## §3. 단락·텍스트 편집

**단락 삽입**: 같은 수준 단락을 `deepcopy` → 불필요 run/t 제거 → 새 `id` → `<hp:linesegarray>` 제거(레이아웃은 한글이 재계산) → `addprevious`/`addnext`. 클론하면 구조 유효성이 보장된다.

```python
import copy
def clone_para(src, text, uid):
    np=copy.deepcopy(src); np.set('id', str(uid()))
    for ls in np.findall(f'{P}linesegarray'): np.remove(ls)
    for r in np.findall(f'{P}run'):                       # 미주 등 불필요 run 제거
        if r.find(f'.//{P}endNote') is not None: np.remove(r)
    placed=False
    for r in np.findall(f'{P}run'):
        if not placed:
            for ch in list(r):
                if ch.tag==f'{P}t': r.remove(ch)
            etree.SubElement(r, f'{P}t').text=text; placed=True
        else:
            for t in r.findall(f'{P}t'): t.text=''
    return np
```

- **새 id 발급**: 문서 내 모든 `id/instId/instid` 최댓값 + 2씩.
- **단락 텍스트 전체교체** 시 **미주(endNote) 포함 단락 주의** — 미주 run을 보존해야 한다(미주 없는 단락만 통째 교체가 안전).
- **단락 삭제**: `root.remove(p)`(또는 `p.getparent().remove(p)`). 여러 개 지울 땐 인덱스가 밀리므로 **요소 참조를 먼저 모으고** 삭제.
- 표는 §5의 정식 생성 방법을 쓰되, 간단할 땐 기존 셀 텍스트에 라벨 접두가 더 안전.

### ★자간·줄간격 깨짐의 원인 (편집·신규 생성 공통)

`<hp:linesegarray>`는 **한글이 미리 계산해 저장해 둔 글자·줄 배치값**(각 세그먼트의 위치·폭)이다. 이게 남아 있으면 한글은 **재조판하지 않고 저장된(어긋난) 배치를 그대로 신뢰**해서 글자가 겹치거나 띄엄띄엄 벌어지는 등 **자간·줄간격이 완전히 망가진다.**

- **워드 만들듯 위치를 박아 HWPX를 새로 생성하거나 변환 도구로 찍어낼 때 깨지는 주원인.**
- **편집이든 무(無)에서 신규 생성이든 `linesegarray`는 절대 넣지 말 것.** 제거하면(또는 애초에 안 넣으면) 한글이 열 때 스스로 정상 재조판한다.
- 신규 단락은 항상 **실제 한글 단락을 deepcopy**해서 만들고(워드식으로 좌표를 계산해 넣지 말 것), linesegarray만 제거한다.
- **부차 원인**: charPr의 script별 장평(`<hh:ratio>`)·자간(`<hh:spacing>`) 값을 잘못 매핑하거나 일부 script를 누락하면 간격이 들쭉날쭉해진다 → 새 charPr를 만들지 말고 **기존 charPr를 재사용**.

---

## §4. 미주(尾註)·각주

구조 — **`<hp:ctrl>` 래핑 필수**(누락 시 한글이 인식·하이퍼링크 못함):

```xml
<hp:run><hp:ctrl><hp:endNote …>…</hp:endNote></hp:ctrl></hp:run>
```

- 번호는 **한글이 위치 기준 자동 재계산** → 문서 순서대로 삽입.
- 각주(footNote)도 동일 구조 — `<hp:endNote>`→`<hp:footNote>`, autoNum `numType`을 `ENDNOTE`→`FOOTNOTE`로만 변경(검증됨; 각주=페이지 하단, 미주=문서 끝).
- 삭제할 단락에 endNote가 없으면 미주 번호는 안 깨진다(삭제 전/후 endNote 수 동일 확인).

---

## §5. 표 (생성·삭제·병합·캡션·배치·높이·음영) — 검증됨

구조:
```
<hp:tbl rowCnt colCnt borderFillIDRef>
  (옵션) <hp:caption>
  <hp:sz width=…>                 # 표 전체 폭 = 각 행 셀 width 합
  <hp:pos treatAsChar=…>          # 배치
  <hp:tr><hp:tc>…</hp:tc>…</hp:tr>
```
`tc` = `subList>p>run>t` + `cellAddr(colAddr,rowAddr)` + `cellSpan(colSpan,rowSpan)` + `cellSz(width,height)` + `cellMargin`.

### 신규 생성
기존 표를 deepcopy → `rowCnt`/`colCnt`·`sz width` 설정 → tr 재구성 → 각 tc의 `cellAddr`·`cellSpan(보통 1,1)`·`cellSz`·텍스트 설정 → `linesegarray` 제거 → 새 id.

- **★열 수를 바꾸면 `cellSz width`와 `sz`를 콜럼수에 맞게 재계산**(안 하면 기하구조가 깨져 한글이 거부).
- **★셀 높이: `cellSz`의 height까지 반드시 리셋.** width만 바꾸고 클론 원본의 큰 height(예 5468)를 두면 칸이 비대해진다. **본문 한 줄 = 282, 헤더 = 0(자동).** 한글이 내용에 맞춰 자동 확장하므로 282는 최소값.

```python
def set_cell(tc, col, row, text, W, colspan=1, rowspan=1, height=282):
    tc.find(f'{P}cellAddr').set('colAddr', str(col)); tc.find(f'{P}cellAddr').set('rowAddr', str(row))
    sp=tc.find(f'{P}cellSpan'); sp.set('colSpan', str(colspan)); sp.set('rowSpan', str(rowspan))
    cs=tc.find(f'{P}cellSz'); cs.set('width', str(sum(W[col:col+colspan]))); cs.set('height', str(height))
    p=tc.find(f'.//{P}p')
    for ls in p.findall(f'{P}linesegarray'): p.remove(ls)
    r=p.find(f'{P}run')
    for ch in list(r):
        if ch.tag==f'{P}t': r.remove(ch)
    etree.SubElement(r, f'{P}t').text=text
```

### 세로 병합 (rowSpan)
- 병합 **시작** tc: `cellAddr(col=0,row=R)` + `cellSpan(colSpan=1, rowSpan=N)`.
- **연속 행은 col0 tc를 생략** → 항목이 `col=1`부터 시작(셀 수가 한 칸 적음).

```
행0(헤더): 5칸 (col 0,1,2,3,4)
행1(성별, rowSpan=2): 5칸,  col0=성별
행2(여):    4칸 (col 1,2,3,4)   ← col0 없음(병합에 가려짐)
행3(연령, rowSpan=9): 5칸,  col0=연령
행4~11:     각 4칸
```

### 가로 병합 (colSpan)
시작 tc의 `colSpan`을 키우고, 가려지는 tc는 제거 + `cellSz` 너비 합산.

### 캡션 (검증됨)
```xml
<hp:caption side="TOP">
  <hp:subList><hp:p><hp:run>
    <hp:t>&lt;표 </hp:t>
    <hp:ctrl><hp:autoNum numType="TABLE">…</hp:autoNum></hp:ctrl>   <!-- 그림은 PICTURE -->
    <hp:t>&gt; 캡션텍스트</hp:t>
  </hp:run></hp:p></hp:subList>
</hp:caption>
```
- autoNum은 **미주처럼 `<hp:ctrl>` 래핑**, 한글이 위치 기준 자동 번호.
- **캡션 텍스트 변경 = 마지막 `<hp:t>` 교체.**
- 자동번호 `<표 N>`을 원치 않으면 `caption` 요소를 제거.

### 글자처럼 취급 (treatAsChar)
- **표**: `<hp:pos treatAsChar="1">` = 인라인(글자처럼), `"0"` = 떠있음. 레퍼런스 표는 전부 `1`.
- **그림**: `<hp:pos>`가 **아예 없으면 인라인(글자처럼)이 기본**, `<hp:pos treatAsChar="0">`면 떠있음.

### 헤더 음영
- 헤더 셀 `borderFillIDRef` = 음영 정의, 본문 셀 = 일반(예: 헤더 `borderFill 8`·`charPr 7`, 본문 `borderFill 3`·`charPr 19`).
- **클론 시 헤더/본문 tc 템플릿을 따로 떠서** 써야 본문에 음영이 안 번진다. 헤더 강조는 run `charPrIDRef`를 굵게 정의로.

### 삭제·검증
- 삭제: 표 포함 단락 `getparent().remove()`.
- 검증: `tbl` 수·셀 텍스트·**각 행 셀수**(헤더·병합시작=전체열, 연속행=−1)·`pic` 수.

---

## §6. 이미지 삽입 (검증됨)

- 인라인 그림 참조 `<hc:img binaryItemIDRef="imageN">`은 **hp가 아니라 hc(core) 네임스페이스**(hp로 찾으면 None).
- 등록: `Contents/content.hpf`의 `<opf:manifest>`에
  `<opf:item id="imageN" href="BinData/imageN.png" media-type="image/png" isEmbeded="1"/>` 추가.
  `META-INF/manifest.xml`은 비어 있어 손댈 필요 없음.
- pic는 기존 `<hp:pic>` 런을 deepcopy → `binaryItemIDRef`·`id`·`instid`·크기(`orgSz`/`curSz`/`imgRect` 4점/`imgClip`/`imgDim`/`sz`) 교체.
- 크기 단위 **HWPUNIT = px × 75** (96dpi).
- 재압축 시 새 BinData 파일은 신규 엔트리 → **별도 add(ZIP_DEFLATED)**.
- 인라인 vs 떠있음: §5의 treatAsChar 규칙과 동일(`<hp:pos>` 없으면 인라인).
- 검증: `pic` 수 증가·content.hpf에 id 존재·BinData 포함·`zip.testzip`.

---

## §7. 글자·문단 서식 (검증됨)

- **글자 서식**(크기·글꼴·굵게·기울임·밑줄·글자색 `textColor`·형광펜 `shadeColor`)은 `header.xml`의 **charPr**가 결정.
- **문단 서식**(정렬 `align`·개요수준·줄간격)은 **paraPr**가 결정.
- 적용 = 본문 run의 `charPrIDRef`·단락의 `paraPrIDRef`를 원하는 정의로 변경(**기존 정의 재사용이 가장 안전**).
- 새 서식 = charPr/paraPr를 deepcopy해 속성·기존 자식 수정 후 등록.
- **★중요: header.xml에 charPr/paraPr를 추가하면 `<hh:charProperties>`/`<hh:paraProperties>`의 `itemCnt`를 실제 개수에 맞게 갱신**(불일치 시 한글이 파일 거부).
- 글자크기 `height = pt × 100`. 글꼴은 charPr의 `<hh:fontRef>`가 header fontfaces의 글꼴 id 참조.
- 색·음영·밑줄 type·fontRef는 속성/기존자식 수정이라 안전. **`<hh:bold/>`·`<hh:italic/>` 자식 신규 추가는 스키마 순서를 타니** 기존 bold/italic charPr 재사용이 안전.

---

## §8. 띄어쓰기·재번호

### 띄어쓰기 교정
한글에 숫자·영문이 공백 없이 붙은 오류가 흔함(예: `호소13건`→`호소 13건`, `진단은KDIGO`→`진단은 KDIGO`).

- **절대 `[가-힣][0-9]`/`[가-힣][A-Za-z]` 일괄 정규식 치환 금지** — `제N절`·`제N장`·`사례군N`·`유형N`·`표N`·코드(`N17` 등)·`비타민D` 같은 의도된 표기와, 특히 **표의 인접 셀 경계(거짓양성)**를 망친다.
- 절차: 경계를 문맥과 함께 **전수조사** → 실제 본문 오류만 골라 **구체 문자열 치환 dict** 작성 → `el.text`와 `el.tail` 모두 적용 → 각 키 적용수>0 확인 + 재조사로 남은 경계가 의도된 예외뿐인지 검증.
- 규칙: 숫자+단위(`5건`·`3단계`)는 붙여도 무방. 오류는 한글 어미/단어가 숫자·영문에 직접 붙은 경우.

### 항·장 재번호
- 번호 기반 내부 상호참조가 **stale** 됨(예: 쟁점사례 3→2 재번호 시 본문 `본 절3에서`가 옛 번호를 가리킴).
- 재번호 후 `절N`·`표N`·`유형N` 식 참조를 점검하고, 가능하면 **명칭 기반**(`쟁점사례 항에서`)으로 바꿔 견고화.

---

## §9. 작업 흐름·검증

- 출력을 **연쇄(rtN→rtN+1)로 누적 빌드**, 매 단계 검증:
  - well-formed XML / `endNote` 수 / `pic` 수 / `tbl` 수 / `mimetype` 첫엔트리·STORED / `zip.testzip()`
- present_files로 제시 → **한글 라운드트립**(열림·하이퍼링크·렌더링) 확인 요청.
- 편집 전 대상의 정확한 **단락 인덱스·스타일·원문 t-노드를 먼저 추출**.

### ★유효성 자가검증
- **LibreOffice는 HWPX를 아예 지원 안 함**(원본조차 `source file could not be loaded`) → 변환 기반 검증 불가.
- 대신 **"원본 sectionN.xml을 lxml로 무수정 재직렬화한 것" vs "편집본"을 줄단위 diff**해 의도한 변경만 있는지 증명하면, 원본이 열리는 한 편집본도 안전하다고 간접 보증(단 재압축 이슈는 별개 — 반드시 작업본 기반).

```python
import re, difflib
def norm(xml_bytes):
    s=etree.tostring(etree.fromstring(xml_bytes), encoding='unicode')
    return re.sub('><','>\n<',s).splitlines()
diff=[l for l in difflib.unified_diff(norm(orig), norm(edited), lineterm='')
      if l and l[0] in '+-' and not l.startswith(('+++','---'))]
# diff가 의도한 변경만 포함하는지 육안 확인
```

---

## 부록: 새 id 발급 헬퍼

```python
def make_uid(root):
    ids=set()
    for el in root.iter():
        for a in ('id','instId','instid'):
            v=el.get(a)
            if v and str(v).isdigit(): ids.add(int(v))
    ctr=[max(ids)+5]
    def uid():
        ctr[0]+=2; return ctr[0]
    return uid
```

---

## 부록: Windows 콘솔 한글 출력

Windows에서 Python으로 HWPX 텍스트를 콘솔에 출력할 때 한글이 깨지는 문제:

```python
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
```

스크립트 상단에 넣어야 `print()`/`cat()` 출력이 정상. HWPX 내부 XML은 UTF-8이므로 파싱 자체는 문제없고, Windows stdout의 기본 인코딩(cp949)이 원인.
