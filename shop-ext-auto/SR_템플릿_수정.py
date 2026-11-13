"""쇼핑 접수 엑셀 ↔ SR(스마트 리다이렉트) 쇼핑 캠페인 비교 → 일목표에 맞춰 캠페인별 기본정책(키워드 재료) 연결.

  python SR_템플릿_수정.py [엑셀]        엑셀 생략 = 다운로드 폴더의 최신 쇼핑접수_*.xlsx
  python SR_템플릿_수정.py --dry [엑셀]  바꿀 목록만 보고 안 바꿈

SR_접수.py로 등록하고 몇 분 뒤(SR 리스트에 캠페인이 생긴 뒤) 실행. 다시 실행해도 이미 맞는 캠페인은 건너뜀.
같은 조건 캠페인 건수가 SR과 다르면(없음·모자람·많음) 그 줄은 실패(종료 코드 1), 맞는 줄만 바꿈.
일목표 100 미만 = sr-100미만(KW-022), 100 이상 = sr-100이상(KW-017).
매칭 = 영업점[SR] + 순위키워드 + 등록일(엑셀 날짜), 엑셀·SR 건수가 같아야 함.
SR 계정 = %LOCALAPPDATA%\\shop-ext-chrome\\sr_account.json {"id": "...", "pw": "..."} (공개 저장소라 저장소 밖)
"""
import argparse, collections, datetime, glob, json, os, re, sys
import openpyxl, requests

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
API = 'https://admin-api.smart-redirect.com'
CT = {'campaignType': 'NS_TRAFFIC'}  # SR 쇼핑 모드
MATERIAL = {'under': ('22', 'KW-022'), 'over': ('17', 'KW-017')}  # 키워드 재료 id, code
ACCOUNT_FILE = os.path.join(os.environ['LOCALAPPDATA'], 'shop-ext-chrome', 'sr_account.json')


def want(daily):
    return MATERIAL['over' if daily >= 100 else 'under']


def read_excel(src):
    rows = list(openpyxl.load_workbook(src, read_only=True).active.iter_rows(values_only=True))
    col = {str(h or '').split('\n')[0].strip(): i for i, h in enumerate(rows[0])}
    out = collections.defaultdict(list)  # (영업점[SR], 순위키워드) → [일목표…]
    for r in rows[1:]:
        if r[col['영업점']] and r[col['순위키워드']]:
            key = (str(r[col['영업점']]).strip() + '[SR]', str(r[col['순위키워드']]).strip())
            out[key].append(int(r[col['일목표']] or 0))
    return out


def excel_date(src):
    m = re.search(r'_(\d{6})\.xlsx$', src)  # 쇼핑접수_업체_YYMMDD.xlsx
    if m:
        return datetime.datetime.strptime(m.group(1), '%y%m%d').strftime('%Y-%m-%d')
    m = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(src))
    return m.group(1) if m else datetime.date.fromtimestamp(os.path.getmtime(src)).isoformat()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx', nargs='?')
    ap.add_argument('--dry', action='store_true')
    a = ap.parse_args()
    src = a.xlsx or max((f for f in glob.glob(os.path.join(os.environ['USERPROFILE'], 'Downloads', '쇼핑접수_*.xlsx'))
                         if not f.endswith('_결과.xlsx')), key=os.path.getmtime, default=None)
    if not src:
        sys.exit('엑셀 파일이 없습니다.')
    if not os.path.exists(ACCOUNT_FILE):
        sys.exit(f'SR 계정 파일이 없습니다: {ACCOUNT_FILE}\n  {{"id": "아이디", "pw": "비밀번호"}}')
    acc = json.load(open(ACCOUNT_FILE, encoding='utf-8'))
    day = excel_date(src)
    rows = read_excel(src)
    print(f'엑셀: {src} ({day}, {sum(map(len, rows.values()))}행)')

    s = requests.Session()
    r = s.post(API + '/auth/login', json={'loginId': acc['id'], 'password': acc['pw']}, timeout=20)
    r.raise_for_status()
    s.headers['Authorization'] = 'Bearer ' + r.json()['data']['accessToken']
    camps, page = [], 1
    while True:
        d = s.get(API + '/campaigns', params={**CT, 'limit': 100, 'page': page}, timeout=30).json()
        camps += d['data']
        if page * 100 >= d['meta']['total']:
            break
        page += 1

    sr = collections.defaultdict(list)
    for c in camps:
        if str(c.get('issuedAt', '')).startswith(day):
            sr[(c['branch'], c['rankKeyword'])].append(c)
    missing = changed = 0
    for key, dailies in rows.items():
        found = sr.get(key, [])
        if len(found) != len(dailies):  # 모자라면 아직 안 올라옴, 많으면 같은 날 다른 접수와 섞여 어느 건인지 모름
            print(f'✖ {key[0]} {key[1]}: 엑셀 {len(dailies)}건 / SR {len(found)}건 — 개수가 달라 실패')
            missing += len(dailies)
            continue
        targets = {want(x) for x in dailies}
        if len(targets) > 1:  # 같은 키에 100 미만·이상이 섞이면 어느 캠페인이 어느 일목표인지 모름
            print(f'✖ {key[0]} {key[1]}: 일목표 {dailies} — 100 기준이 섞여 어느 캠페인인지 모름, 실패(직접 설정)')
            missing += len(dailies)
            continue
        mid, code = targets.pop()
        for c in found:
            if c.get('materialCode') == code:
                print(f'= {c["code"]} {key[1]} {c["midId"]}: 이미 {code}')
                continue
            print(f'→ {c["code"]} {key[1]} {c["midId"]}: {c.get("materialCode")} → {code} (일목표 {dailies})')
            if not a.dry:
                s.patch(f'{API}/campaigns/{c["code"]}/policy/material', json={'op': 'relink', 'masterId': mid}, timeout=30).raise_for_status()
            changed += 1
    print(f'{"변경 예정" if a.dry else "변경"} {changed}건' + (f' · 실패 {missing}건(SR에 같은 영업점·순위키워드·등록일 캠페인 건수가 다름 — 아직 안 올라왔으면 몇 분 뒤 다시 실행)' if missing else ''))
    sys.exit(1 if missing else 0)


if __name__ == '__main__':
    main()
