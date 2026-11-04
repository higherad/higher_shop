"""쇼핑 접수 엑셀 → 확장(키워드 추출 · 캠페인 등록) 업로드 → 1 정보 불러오기 → 2 유효키워드 추출 → 3 캠페인 등록.

  python run.py --setup                          전용 크롬 처음 띄우기(확장 설치·설정·영업점 선택은 사람이 1회)
  python run.py [엑셀]   엑셀 생략 = 다운로드 폴더의 최신 쇼핑접수_*.xlsx

3번은 확장의 '이대로 등록할까요?(예상 차감액)' 창에서 사람이 '등록하기'를 눌러야 진행(과금 확인).
영업점·등급은 AGENCY·GRADE 고정(확장 상단 영업점도 자동 선택).
"""
import argparse, csv, glob, json, os, subprocess, sys, tempfile, time, urllib.request
import openpyxl
from playwright.sync_api import sync_playwright
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True); sys.stderr.reconfigure(encoding="utf-8")  # 파이프 실행에서도 바로 출력

EXT_ID = 'kljeaomfnpdlaphobfapgcjkoaaphapo'
EXT_URL = f'chrome-extension://{EXT_ID}/popup.html?view=full'
STORE_URL = f'https://chromewebstore.google.com/detail/{EXT_ID}'
PORT = 9223
# 크롬은 기본 사용자 폴더에 원격 조종을 막아서 전용 폴더를 씀
PROFILE = os.path.join(os.environ['LOCALAPPDATA'], 'shop-ext-chrome')
CHROME = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
STEP_TIMEOUT = 60 * 60  # 단계별 최대 대기(초)
AGENCY, GRADE = '[SR]하이어애드', '스모키'
LAUNCHER = 'http://127.0.0.1:9334'  # apihub-launcher(크롤 창 컨트롤러)
CRAWL_CDP = 'http://127.0.0.1:9333'  # 크롤 창 크롬
# 크롤 창 번호별 네이버 계정 {"1": ["아이디", "비번"], ...} — 저장소 밖(공개 저장소라 비번 커밋 금지)
ACCOUNTS_FILE = os.path.join(PROFILE, 'naver_accounts.json')
LOGIN_TIMEOUT = 10 * 60
HEAD = ['영업점', '등급', '상품URL', '순위키워드', '당일 진행 여부', '일수', '일목표']


def start_chrome(url):
    subprocess.Popen([CHROME, f'--remote-debugging-port={PORT}', f'--user-data-dir={PROFILE}', url])


def build_upload(src, agency, grade, target):
    """어드민 양식(상품URL·순위키워드·당일 진행 여부·일수·일목표) → 확장 업로드 양식."""
    ws = openpyxl.load_workbook(src, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    col = {str(h or '').split('\n')[0].strip(): i for i, h in enumerate(rows[0])}
    out = []
    for r in rows[1:]:
        get = lambda k: r[col[k]] if k in col and col[k] < len(r) and r[col[k]] is not None else ''
        if not get('상품URL'):
            continue
        out.append([agency, grade, get('상품URL'), get('순위키워드'), get('당일 진행 여부') or '익일진행',
                    get('일수'), get('일목표') or target])
    if not out:
        sys.exit('엑셀에 상품URL이 있는 행이 없습니다.')
    wb = openpyxl.Workbook(); wb.active.append(HEAD)
    for r in out:
        wb.active.append(r)
    path = os.path.join(tempfile.gettempdir(), 'shop_ext_upload.xlsx')
    wb.save(path)
    return path, len(out)


def bulk(page):
    return page.evaluate("chrome.storage.session.get('bulk').then(o => o.bulk || {})")


def close_alert(page):
    """안내창만 닫음. 취소 버튼이 보이는 확인창(등록 과금 확인)은 절대 누르지 않음."""
    if page.is_visible('#notice-later'):
        page.click('#notice-later')
    if page.is_visible('#result-overlay') and not page.is_visible('#result-cancel'):
        page.click('#result-ok')


def run_step(page, btn, name, human=False):
    close_alert(page)
    page.click(btn)
    print(f'▶ {name} 시작')
    t0 = time.time()
    while not bulk(page).get('running'):  # 시작 대기
        if human and page.is_visible('#result-cancel'):
            print('  ⏸ 브라우저 확인창에서 금액 확인 후 「등록하기」를 눌러 주세요.')
            while page.is_visible('#result-cancel'):
                time.sleep(1)
            continue
        if time.time() - t0 > (600 if human else 20):
            msg = page.inner_text('#result-body') if page.is_visible('#result-overlay') else page.inner_text('#status')
            sys.exit(f'✖ {name} 시작 안 됨: {msg.strip()}')
        time.sleep(1)
    warned = False
    while (b := bulk(page)).get('running'):
        if page.is_visible('#busy-challenge') and not warned:
            print('  ⚠ 네이버 보안확인이 떴습니다 — 브라우저에서 풀어 주세요.'); warned = True
        p = b.get('progress') or {}
        print(f"\r  {p.get('done', 0)}/{p.get('total', '?')}", end='', flush=True)
        if time.time() - t0 > STEP_TIMEOUT:
            sys.exit(f'\n✖ {name} {STEP_TIMEOUT // 60}분 초과')
        time.sleep(2)
    print(f"\n✔ {name} 끝: {b.get('summary', '')}")


CRAWL_RETRY = 3  # 정보 불러오기 실패 행 재시도 횟수
# 정보 불러오기 실패(오류·다시 불러오기 필요) 행의 MID~옵션키워드를 비움 — 확장은 빈 칸만 다시 불러옴.
#   _crawlSeq·_optionKwSeq는 남겨야 옛 실패 결과가 다시 편입되지 않음
CLEAR_FAILED = """() => chrome.storage.local.get('gridRows').then(({ gridRows: g = [] }) => {
  let n = 0;
  g.forEach(r => {
    if (!String(r.productUrl || '').trim()) return;
    if (!(r._crawlStatus || r._optionKwRetry || !String(r.mid || '').trim() || !String(r.productName || '').trim())) return;
    ['mid', 'productName', 'thumb', 'storeName', 'price', 'hashtag', 'optionKw', '_crawlStatus', '_crawlReason', '_optionKwRetry'].forEach(k => r[k] = '');
    n++;
  });
  return n ? chrome.storage.local.set({ gridRows: g }).then(() => n) : 0;
})"""


def launcher(path, body=None, timeout=30):
    req = urllib.request.Request(LAUNCHER + path, data=None if body is None else json.dumps(body).encode(),
                                 headers={'content-type': 'application/json'})
    return json.load(urllib.request.urlopen(req, timeout=timeout))


def ensure_logins(p):
    """계정이 정해진 크롤 창 중 로그아웃된 창만 그 계정으로 로그인(창 1개 = 계정 1개). 보안문자·기기 등록은 사람이 처리할 때까지 대기."""
    if not os.path.exists(ACCOUNTS_FILE):
        print(f'(네이버 자동 로그인 건너뜀 — {ACCOUNTS_FILE} 없음)')
        return
    acc = json.load(open(ACCOUNTS_FILE, encoding='utf-8'))
    try:
        st = launcher('/wake', {}, 90)  # 꺼진 크롬·닫힌 창 다시 열기
    except Exception as e:
        sys.exit(f'크롤 창 프로그램(apihub-launcher)이 응답하지 않습니다: {e}')
    need = [w['n'] for w in st.get('windows', []) if str(w['n']) in acc and not w.get('loggedIn')]
    if not need:
        return
    br = p.chromium.connect_over_cdp(CRAWL_CDP)
    for n in need:
        uid, pw = acc[str(n)]
        launcher('/goto', {'n': n, 'url': 'https://nid.naver.com/nidlogin.login'})
        launcher('/wait', {'n': n, 'part': 'nidlogin', 'timeoutMs': 15000}, 20)
        launcher('/eval', {'n': n, 'expr': f'window.__shopN = {n}'})  # 창 번호 표시 → CDP 페이지 찾기
        page = next((pg for c in br.contexts for pg in c.pages
                     if 'nidlogin' in pg.url and pg.evaluate('window.__shopN') == n), None)
        if not page:
            print(f'  ⚠ {n}번 창 로그인 화면을 찾지 못함 — 직접 로그인해 주세요.')
            continue
        page.bring_to_front()
        page.click('#id'); page.keyboard.type(uid, delay=120)  # 실제 키 입력(값 대입은 네이버가 자동입력으로 봄)
        page.click('#pw'); page.keyboard.type(pw, delay=120)
        page.keyboard.press('Enter')
        print(f'  {n}번 창 {uid} 로그인 시도')
    t0, warned = time.time(), False
    while True:
        left = [w['n'] for w in launcher('/status').get('windows', []) if w['n'] in need and not w.get('loggedIn')]
        if not left:
            print('✔ 네이버 로그인 완료'); return
        if time.time() - t0 > 20 and not warned:
            print(f'  ⏸ {left}번 창 로그인 대기 — 보안문자·새 기기 등록 화면이면 그 창에서 직접 처리해 주세요.'); warned = True
        if time.time() - t0 > LOGIN_TIMEOUT:
            sys.exit(f'✖ {left}번 창 로그인 {LOGIN_TIMEOUT // 60}분 초과')
        time.sleep(3)


def crawl_with_retry(page):
    run_step(page, '#btn-crawl', '1. 정보 불러오기')
    for k in range(CRAWL_RETRY):
        n = page.evaluate(CLEAR_FAILED)
        if not n:
            return
        print(f'↻ 실패 {n}행 비우고 다시 불러오기 ({k + 1}/{CRAWL_RETRY})')
        page.reload(); page.wait_for_selector('#btn-crawl'); time.sleep(1)  # 확장 화면이 비운 시트를 다시 읽게
        run_step(page, '#btn-crawl', '1. 정보 불러오기(재시도)')
    if page.evaluate(CLEAR_FAILED):
        page.reload()
        sys.exit(f'✖ {CRAWL_RETRY}번 재시도해도 정보 불러오기 실패한 행이 있습니다.')


def save_result(page, src):
    rows = page.evaluate("""Promise.all([chrome.storage.local.get('gridRows'), chrome.storage.session.get('bulk')])
      .then(([g, s]) => (g.gridRows || []).map((r, i) => ({ ...r, ...((s.bulk || {}).rows || [])[i] || {} })))""")
    out = os.path.splitext(src)[0] + '_결과.csv'
    with open(out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['상품URL', '순위키워드', '유효키워드 개수', '등록상태', '캠페인코드', '실패사유'])
        for r in rows:
            if r.get('productUrl'):
                v = r.get('valid') or ''
                n = len(v) if isinstance(v, list) else len([x for x in str(v).split(',') if x.strip()])
                w.writerow([r['productUrl'], r.get('rankKeyword', ''), n, r.get('regStatus') or r.get('regState', ''),
                            r.get('regCode', ''), r.get('regMsg') or r.get('regReason', '')])
    print('결과 저장:', out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx', nargs='?')
    ap.add_argument('--grade', default=GRADE, help=f'NS트래픽 등급 이름(기본 {GRADE})')
    ap.add_argument('--target', type=int, default=100, help='엑셀 일목표가 빈 행에 넣을 값(기본 100)')
    ap.add_argument('--setup', action='store_true')
    a = ap.parse_args()
    if a.setup:
        start_chrome(STORE_URL)
        print('전용 크롬에서: 확장 설치 → 설정(파트너키·프록시) → 확장 화면 상단 영업점 선택. 한 번만 하면 됩니다.')
        return
    if not a.grade or not a.target or a.target < 10:
        sys.exit('--grade 와 --target(10 이상)을 넣어 주세요.')
    src = a.xlsx or max(glob.glob(os.path.join(os.environ['USERPROFILE'], 'Downloads', '쇼핑접수_*.xlsx')),
                        key=os.path.getmtime, default=None)
    if not src:
        sys.exit('엑셀 파일이 없습니다. 어드민 접수 현황에서 엑셀 다운로드 후 실행하세요.')
    print('엑셀:', src)

    with sync_playwright() as p:
        try:
            br = p.chromium.connect_over_cdp(f'http://127.0.0.1:{PORT}')
        except Exception:
            start_chrome('about:blank'); time.sleep(4)
            br = p.chromium.connect_over_cdp(f'http://127.0.0.1:{PORT}')
        ctx = br.contexts[0]
        page = next((pg for pg in ctx.pages if pg.url.startswith(f'chrome-extension://{EXT_ID}/popup.html')), None) or ctx.new_page()
        page.on('dialog', lambda d: d.accept())  # '시트 내용 교체' 확인(과금 무관)
        try:
            page.goto(EXT_URL)
        except Exception:
            sys.exit('전용 크롬에 확장이 설치돼 있지 않습니다. python run.py --setup 으로 설치·설정부터 해 주세요.')
        page.wait_for_selector('#btn-crawl'); time.sleep(1)
        if bulk(page).get('running'):
            sys.exit('확장에서 이미 작업이 진행 중입니다.')
        # 확장 상단 영업점 = AGENCY로 맞춤(과금 대상, 확장이 저장값 partnerKeyIdx로 기억)
        idx = page.evaluate("n => chrome.storage.local.get('partnerKeys').then(s => (s.partnerKeys || []).findIndex(k => k.name === n))", AGENCY)
        if idx < 0:
            sys.exit(f'확장 설정에 영업점 {AGENCY} 이 없습니다.')
        page.evaluate("i => chrome.storage.local.set({ partnerKeyIdx: i })", idx)
        page.reload(); page.wait_for_selector('#btn-crawl'); time.sleep(1)
        agency = page.inner_text('#agency-current').strip()
        if agency != AGENCY:
            sys.exit(f'영업점 선택 실패(화면: {agency})')
        up, n = build_upload(src, agency, a.grade, a.target)
        print(f'영업점 {agency} · 등급 {a.grade} · {n}행 업로드')
        close_alert(page)
        page.set_input_files('#excel-file', up)
        time.sleep(3)
        ensure_logins(p)
        crawl_with_retry(page)
        run_step(page, '#btn-extract', '2. 유효키워드 추출')
        run_step(page, '#btn-register', '3. 캠페인 등록', human=True)
        close_alert(page)
        save_result(page, src)


if __name__ == '__main__':
    main()
