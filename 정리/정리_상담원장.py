# 상담원장 정리 스크립트
#
# 결정 기록: my-wiki/decisions/d-006-상담원장-정리-규칙.md
#   d-006 #1  처리상태 «종결» → «완료», «대기» → «보류» (CS기획팀 답)
#   d-006 #2  처리시간 음수는 지우지 않고 값을 그대로 둔다. _표시 칸에 이유
#   d-006 #3  만족도 6~10 은 지우지 않고 값을 그대로 둔다. _표시 칸에 이유. 평균 만족도에서 뺀다
#   d-006 #4  칸 이름은 원본 그대로 (맨 끝에 _표시 칸만 더한다)
#   d-006 #5  고객번호 앞뒤 공백을 없앤다
# 그 밖의 처리는 원본 «작성기준» 시트를 따르거나, d-006 «아직 정하지 않은 것»에 적힌 대로
# 첫 정리 때의 방식을 그대로 둔 것이다.
#
# 원본은 읽기만 한다. 결과:
#   정리본/상담원장_정리.csv   정리본/지운줄.csv   정리본/변경요약.md
#
# 실행: python 정리/정리_상담원장.py            (cleaner 폴더에서)
#       python 정리/정리_상담원장.py --src <원본.xlsx> --out <결과폴더>

import argparse, collections as C, csv, re, sys, io
from pathlib import Path
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
HERE = Path(__file__).resolve().parent.parent          # cleaner/
ap = argparse.ArgumentParser()
ap.add_argument('--src', default=str(HERE.parent / 'messy' / '상담원장_2025.xlsx'))
ap.add_argument('--out', default=str(HERE / '정리본'))
a = ap.parse_args()
SRC, OUT = Path(a.src), Path(a.out)
OUT.mkdir(parents=True, exist_ok=True)

# ---------------- 표기 대응표 ----------------
CH = {'전화':'전화','유선':'전화','전화상담':'전화','TEL':'전화',
      '챗봇':'챗봇','chatbot':'챗봇','Chatbot':'챗봇','챗 봇':'챗봇',
      '앱':'앱','app':'앱','APP':'앱','모바일앱':'앱',
      '이메일':'이메일','e-mail':'이메일','E-mail':'이메일','메일':'이메일'}
REGION = {'경기도':'경기','서울시':'서울','서울특별시':'서울'}
STATUS = {'완료':'완료','처리완료':'완료',
          '종결':'완료',                         # d-006 #1
          '진행중':'진행중','처리중':'진행중',
          '보류':'보류',
          '대기':'보류'}                         # d-006 #1
REASK = {'Y':'Y','y':'Y','예':'Y','TRUE':'Y','N':'N','n':'N','아니오':'N','FALSE':'N'}
FLAG_NEG = '처리시간 음수: 틀린 값이지만 진짜 값을 모름 (d-006 #2)'
FLAG_SAT = '만족도 6~10: NPS를 잘못 넣은 것으로 보이나 미확인, 평균 만족도에서 뺌 (d-006 #3)'

# ---------------- 원본 읽기 ----------------
ws = openpyxl.load_workbook(SRC, read_only=True)['원장']
it = ws.iter_rows(values_only=True)
H = list(next(it))
RAW = [(i, list(r)) for i, r in enumerate(it, start=2) if any(v is not None for v in r)]
OUT_H = H + ['_표시']                             # d-006 #4

def norm(raw):
    """한 줄을 바꾼다. (바뀐 줄, 바꾼 것 목록) 을 돌려준다."""
    r = dict(zip(H, raw)); ch = []; flags = []
    d = re.sub(r'[/.]', '-', r['접수일'])
    if d != r['접수일']: ch.append('접수일 표기 → YYYY-MM-DD')
    r['접수일'] = d
    c = r['고객번호'].strip()                   # d-006 #5
    if c != r['고객번호']: ch.append('고객번호 앞뒤 공백 없앰 (d-006 #5)')
    r['고객번호'] = c
    g = REGION.get(r['지역'], r['지역'])
    if g != r['지역']: ch.append(f'지역 표기 → {g}')
    r['지역'] = g
    t = r['상담원팀'].replace(' ', '')
    if t != r['상담원팀']: ch.append('상담원팀 띄어쓰기 없앰')
    r['상담원팀'] = t
    raw_ch = r['접수채널']; c = CH[raw_ch.strip()]
    if c != raw_ch: ch.append(f'접수채널 표기 → {c}')
    r['접수채널'] = c
    if isinstance(r['월정액'], str):
        r['월정액'] = int(r['월정액'].replace(',', '').replace('원', '')); ch.append('월정액 «N,NNN원» → 숫자')
    s0 = r['처리상태']; s = STATUS[s0.strip()]
    if s != s0:
        if s0.strip() in ('종결', '대기'): ch.append(f'처리상태 {s0.strip()} → {s} (d-006 #1)')
        else: ch.append(f'처리상태 «{s0}» → {s}')
    r['처리상태'] = s
    # 상담시간: 작성기준 «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위»
    if c == '전화' and r['접수일'] >= '2025-07-01':
        r['상담시간'] = round(r['상담시간'] / 60, 2); ch.append('상담시간 초 → 분 (7월 이후 전화)')
    # 처리시간 음수: 값은 그대로, 표시만 (d-006 #2)
    if isinstance(r['처리시간'], (int, float)) and r['처리시간'] < 0:
        flags.append(FLAG_NEG); ch.append('_표시: 처리시간 음수 (d-006 #2)')
    q = r['재문의']
    if q is not None:
        r['재문의'] = REASK[q]
        if r['재문의'] != q: ch.append(f'재문의 표기 → {r["재문의"]}')
    # 만족도
    v = r['만족도']
    if isinstance(v, str):
        m = re.fullmatch(r'([1-5])점', v)
        if m: r['만족도'] = int(m.group(1)); ch.append('만족도 «N점» → 숫자')
        else: r['만족도'] = None; ch.append('만족도 «-»·«미응답» → 빈칸 (작성기준)')
    elif v == 0:
        r['만족도'] = None; ch.append('만족도 0 → 빈칸 (작성기준: 옛 시스템 미응답)')
    elif isinstance(v, int) and v > 5:
        flags.append(FLAG_SAT); ch.append('_표시: 만족도 6~10 (d-006 #3)')   # 값은 그대로
    r['_표시'] = ' / '.join(flags) if flags else None
    return r, ch

# ---------------- 1) 테스트 접수 ----------------
removed, step1 = [], []
for i, raw in RAW:
    if raw[H.index('상담원ID')] == 'AG999': removed.append((i, raw, '테스트 접수 (상담원ID AG999)'))
    else:
        n, ch = norm(raw); step1.append((i, raw, n, ch))

# ---------------- 2) 완전 중복 (표기를 맞춘 뒤) ----------------
first, step2 = {}, []
for i, raw, n, ch in step1:
    t = tuple(n[k] for k in OUT_H)
    if t in first:
        fi, fraw = first[t]
        why = '완전 중복' if fraw == raw else '완전 중복 (날짜 표기만 다름)'
        removed.append((i, raw, f'{why} — 원본 {fi}행과 같음'))
    else:
        first[t] = (i, raw); step2.append((i, raw, n, ch))

# ---------------- 3) 같은 접수번호 → 최종수정일이 가장 늦은 줄 (작성기준) ----------------
grp = C.defaultdict(list)
for x in step2: grp[x[2]['접수번호']].append(x)
kept, ties = [], 0
for v in grp.values():
    v.sort(key=lambda x: (x[2]['최종수정일'], x[0]))
    if len(v) > 1 and v[-1][2]['최종수정일'] == v[-2][2]['최종수정일']: ties += 1
    kept.append(v[-1])
    for i, raw, n, ch in v[:-1]:
        removed.append((i, raw, f'상태 변경 전 줄 (처리상태 {raw[H.index("처리상태")].strip()}, '
                                f'최종수정일 {n["최종수정일"]}) — 최신은 원본 {v[-1][0]}행'))
kept.sort(key=lambda x: (x[2]['접수일'], x[2]['접수시각'], x[2]['접수번호']))
removed.sort(key=lambda x: x[0])

# ---------------- 쓰기 ----------------
cell = lambda v: '' if v is None else v
with open(OUT / '상담원장_정리.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f); w.writerow(OUT_H)
    for _, _, n, _ in kept: w.writerow([cell(n[k]) for k in OUT_H])
with open(OUT / '지운줄.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f); w.writerow(H + ['원본행', '지운이유'])
    for i, raw, why in removed: w.writerow([cell(v) for v in raw] + [i, why])

# ---------------- 지표 ----------------
F = [n for _, _, n, _ in kept]
N = len(F)
avg_t = sum(r['상담시간'] for r in F) / N
ra = [r for r in F if r['재문의'] is not None]
reask = 100 * sum(r['재문의'] == 'Y' for r in ra) / len(ra)
sat = [r['만족도'] for r in F if r['만족도'] is not None and not (r['_표시'] and FLAG_SAT in r['_표시'])]
avg_s = sum(sat) / len(sat)
done = 100 * sum(r['처리상태'] == '완료' for r in F) / N
cust = len({r['고객번호'] for r in F})
cust_raw = len({raw[H.index('고객번호')] for i, raw, n, ch in step1})

# ---------------- 변경 요약 ----------------
before = C.Counter(c for *_, ch in step1 for c in ch)
after = C.Counter(c for *_, ch in kept for c in ch)
rm = C.Counter(re.split(r' — | \(처리상태', why)[0] for _, _, why in removed)
def tbl_rows(keys):
    return '\n'.join(f'| {k} | {before[k]:,} | {after[k]:,} |' for k in keys)
keys = sorted(before, key=lambda k: (-before[k], k))
md = f"""# 상담원장 변경 요약

- 원본: `../messy/상담원장_2025.xlsx` «원장» 시트 (읽기만 함)
- 스크립트: `정리/정리_상담원장.py` · 결정 기록: `my-wiki/decisions/d-006-상담원장-정리-규칙.md`
- 결과: `정리본/상담원장_정리.csv` · `정리본/지운줄.csv`

## 1. 줄 수

| 단계 | 줄 수 |
|---|---:|
| 원본 (머리글 제외) | {len(RAW):,} |
""" + '\n'.join(f'| − {k} | −{v:,} |' for k, v in rm.items()) + f"""
| **지운 줄 합계** (`지운줄.csv`) | **−{len(removed):,}** |
| **정리본** (`상담원장_정리.csv`) | **{N:,}** |

검산: {len(RAW):,} − {len(removed):,} = {len(RAW) - len(removed):,} → 정리본 {N:,}줄과 {'같음' if len(RAW) - len(removed) == N else '다름'}.
같은 접수번호에서 최종수정일이 같아 고를 수 없던 경우: {ties}건.

## 2. 바꾼 값

«처리 단계»는 테스트 접수를 뺀 {len(step1):,}줄(중복을 지우기 전)에서 센 수, «정리본»은 정리본에 남은 {N:,}줄에서 센 수다. 한 줄에서 여러 가지를 바꾸면 각각 센다.

| 바꾼 것 | 처리 단계 | 정리본 |
|---|---:|---:|
{tbl_rows(keys)}

### 상담시간 단위

**상담시간 칸의 단위가 줄마다 달랐다.** 원본 «작성기준» 시트에 «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위로 기록됨»이라고 적혀 있다. 그래서 접수일이 2025-07-01 이후이고 접수채널이 전화인 줄은 **초를 60으로 나눠 분으로 바꿨다** (처리 단계 {before['상담시간 초 → 분 (7월 이후 전화)']:,}줄 · 정리본 {after['상담시간 초 → 분 (7월 이후 전화)']:,}줄). 정리본의 상담시간은 모두 분 단위다. 이때 «전화»에는 원본의 유선 · 전화상담 · TEL · « 전화»도 들어간다. 이것은 데이터를 보고 짐작한 것이고, d-006 «아직 정하지 않은 것»에 적혀 있다.

### 지우지 않고 표시한 줄 (`_표시` 칸)

| 이유 | 정리본 줄 |
|---|---:|
| {FLAG_NEG} | {sum(1 for r in F if r['_표시'] and FLAG_NEG in r['_표시']):,} |
| {FLAG_SAT} | {sum(1 for r in F if r['_표시'] and FLAG_SAT in r['_표시']):,} |

## 3. 다섯 숫자와 고객 수 (정리본 기준)

| 지표 | 값 | 어떻게 셌나 |
|---|---|---|
| 상담 건수 | {N:,}건 | 접수번호 하나 = 1건 |
| 평균 상담시간 | {avg_t:.2f}분 | 전체 {N:,}건 |
| 재문의율 | {reask:.2f}% | 재문의가 적힌 {len(ra):,}건 중 Y (1~2월은 기록 없음이라 뺌) |
| 평균 만족도 | {avg_s:.3f}점 | 1~5점 응답 {len(sat):,}건. `_표시`된 만족도 6~10은 뺌 |
| 완료율 | {done:.2f}% | 처리상태 «완료» / 전체 {N:,}건 (종결은 완료로 합침) |
| 고객 수 | {cust:,}명 | 서로 다른 고객번호 (공백을 없애기 전 처리 단계 기준이었다면 {cust_raw:,}명) |
"""
(OUT / '변경요약.md').write_text(md, encoding='utf-8')

print(f'원본 {len(RAW)} − 지운 줄 {len(removed)} = 정리본 {N}')
print(f'상담 건수 {N} | 평균 상담시간 {avg_t:.4f} | 재문의율 {reask:.4f} ({len(ra)}) | '
      f'평균 만족도 {avg_s:.4f} ({len(sat)}) | 완료율 {done:.4f} | 고객 수 {cust} (공백 그대로면 {cust_raw})')
