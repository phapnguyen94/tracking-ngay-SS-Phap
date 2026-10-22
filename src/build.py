import pandas as pd, json, os, argparse, glob
from datetime import datetime, timezone, timedelta

ap = argparse.ArgumentParser()
ap.add_argument('--src', default='/mnt/user-data/uploads/')   # thư mục có MUC_TIEU_NGAY_KH.csv + SUB_DMS.csv
ap.add_argument('--orders', nargs='*', default=[])                       # DanhSachChiTietDonHang*.xlsx (đơn T10 đến hiện tại)
ap.add_argument('--r3', default='')                           # CT_-_TH_789*.xlsx (CT/TH T7-T9 theo DDKD)
ap.add_argument('--sched', default='')                        # Visit_Schedule_Report*.xlsx (lịch viếng thật: thứ, tần suất F2/F4/F8, tuần lẻ/chẵn)
ap.add_argument('--data', default='')       # thư mục lưu số liệu chính thức theo ngày
ap.add_argument('--save', action='store_true')   # lưu số liệu các ngày có trong --orders vào --data (nạp chính thức buổi tối)
ap.add_argument('--today', default='')                        # YYYY-MM-DD, mặc định hôm nay (UTC+7)
ap.add_argument('--out', default='dashboard.html')
ap.add_argument('--standalone', action='store_true')   # trang HTML độc lập (GitHub Pages)
ap.add_argument('--password', default='')               # có mật khẩu thì mã hóa dữ liệu AES-GCM
a = ap.parse_args()
here = os.path.dirname(os.path.abspath(__file__))
import calendar
today = datetime.strptime(a.today, '%Y-%m-%d').date() if a.today else datetime.now(timezone(timedelta(hours=7))).date()
today_name = 'Thứ %d' % (today.weekday() + 2) if today.weekday() < 6 else 'Chủ nhật'

k = pd.read_csv(os.path.join(a.src, 'MUC_TIEU_NGAY_KH.csv'), encoding='utf-8-sig')
s = pd.read_csv(os.path.join(a.src, 'MUC_TIEU_NGAY_SUB_DMS.csv'), encoding='utf-8-sig')
dd = sorted(k['DDKD'].unique()); ddi = {d: i for i, d in enumerate(dd)}
sub = sorted(s['Sub division'].unique()); subi = {x: i for i, x in enumerate(sub)}
dms = sorted(s['DMS OBJ'].unique()); dmsi = {x: i for i, x in enumerate(dms)}

# --- đọc & chuẩn hóa file đơn: 2 định dạng ("Danh sách chi tiết đơn hàng" cũ; "BÁO CÁO LINE ITEM" mới từ 22/10)
COLS_STD = ['Mã cửa hàng', 'Ngày tạo đơn', 'Mã SO', 'Mã sản phẩm', 'Thành tiền', 'Khuyến mại', 'Sub Division']
NCANCEL = 0
def read_orders(path):
    global NCANCEL
    raw = pd.read_excel(path, header=None, nrows=12)
    hdr = next(i for i in range(len(raw)) if {'Mã SO', 'Mã đơn hàng'} & set(raw.iloc[i].astype(str)))
    x = pd.read_excel(path, header=hdr)
    if 'Mã đơn hàng' in x.columns:      # LINE ITEM mới: chỉ đơn đã tạo PXK (đang giao / đã giao); Thành tiền = trước CK (khớp file cũ)
        x = x[x['Mã CH'].notna()].rename(columns={'Mã CH': 'Mã cửa hàng', 'Ngày tạo đơn hàng': 'Ngày tạo đơn', 'Mã đơn hàng': 'Mã SO',
                                                  'Thành tiền trước CK': 'Thành tiền', 'Hàng KM': 'Khuyến mại'})
        st = x['Tình trạng đơn hàng'].astype(str) if 'Tình trạng đơn hàng' in x.columns else pd.Series('', index=x.index)
    else:                                # định dạng cũ (có đơn chờ xác nhận, đã hủy)
        st = x['Trạng thái'].astype(str) if 'Trạng thái' in x.columns else pd.Series('', index=x.index)
        if 'Sub Division' not in x.columns: x['Sub Division'] = ''
    cx = st.str.startswith('Đã hủy'); NCANCEL += int(cx.sum())
    x = x[~cx][COLS_STD].copy()
    x['Mã cửa hàng'] = x['Mã cửa hàng'].astype('int64')
    return x

def load_orders(files, data_dir, save):
    def dmy(x): return pd.to_datetime(x['Ngày tạo đơn'], format='%d/%m/%Y %H:%M:%S').dt.normalize()
    parts = []
    if data_dir and os.path.isdir(data_dir):    # số liệu chính thức đã lưu theo ngày
        for fp in sorted(glob.glob(os.path.join(data_dir, 'orders_*.csv.gz'))):
            q = pd.read_csv(fp, dtype={'Sub Division': str}).fillna({'Sub Division': ''}); q['_d'] = dmy(q); parts.append(q)
    for f in files:   # file sau thay toàn bộ các ngày nó phủ (nạp lại trong ngày không bị cộng đôi)
        x = read_orders(f); x['_d'] = dmy(x); lo = x['_d'].min()
        parts = [q[q['_d'] < lo] for q in parts] + [x]
        if save and data_dir:
            os.makedirs(data_dir, exist_ok=True)
            for d, g in x.groupby('_d'):
                g.drop(columns='_d').to_csv(os.path.join(data_dir, 'orders_%s.csv.gz' % d.strftime('%Y-%m-%d')), index=False, compression='gzip')
    print('đơn hủy loại bỏ:', NCANCEL)
    return pd.concat(parts, ignore_index=True).drop(columns='_d')

# --- đơn T10 theo KH
od, tod, tod_n, ord_to, has_today, n_out = {}, {}, {}, '', False, 0
hr = []
if a.orders or a.data:
    o = load_orders(a.orders, a.data, a.save)
    o['d'] = pd.to_datetime(o['Ngày tạo đơn'], format='%d/%m/%Y %H:%M:%S')
    ord_to = o['d'].max().strftime('%d/%m/%Y %H:%M')
    o = o[o['Mã cửa hàng'].isin(k['Mã KH'])] if False else o
    n_out = int((~o['Mã cửa hàng'].isin(k['Mã KH'])).sum())
    o['day'] = o['d'].dt.date
    g = o.groupby(['Mã cửa hàng', 'day']).agg(n=('Mã SO', 'nunique'), amt=('Thành tiền', 'sum')).reset_index()
    # line = số mã sản phẩm khác nhau trên mỗi đơn, bỏ dòng hàng khuyến mại (Khuyến mại = Y)
    ln = o[o['Khuyến mại'] == 'N'][['Mã cửa hàng', 'day', 'Mã SO', 'Mã sản phẩm']].drop_duplicates().groupby(['Mã cửa hàng', 'day']).size()
    g['lines'] = [int(ln.get((r[0], r[1]), 0)) for r in g.itertuples(index=False)]
    for r in g.itertuples(index=False):
        ma = int(r[0])
        od.setdefault(ma, []).append([r[1].day, int(r[2]), int(round(r[3])), int(r[4])])
        if r[1] == today:
            tod[ma] = int(round(r[3])); tod_n[ma] = int(r[2])
    kd = dict(zip(k['Mã KH'].astype(int), k['DDKD'].map(ddi)))
    t = o[(o['day'] == today) & o['Mã cửa hàng'].isin(k['Mã KH'])].copy()
    t['h'] = t['d'].dt.hour
    t['dd'] = t['Mã cửa hàng'].astype(int).map(kd)
    for r in t.groupby(['dd', 'h']).agg(a=('Thành tiền', 'sum'), n=('Mã SO', 'nunique')).reset_index().itertuples(index=False):
        hr.append([int(r[0]), int(r[1]), int(round(r[2])), int(r[3])])
    has_today = any(r[1] == today for r in g.itertuples(index=False))

# --- lịch viếng thật theo Visit Schedule Report: thứ có lịch, tần suất, tuần lẻ/chẵn (ISO week)
DM = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5}
WN = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7']
sch = {}
if a.sched:
    vs = pd.read_excel(a.sched, header=2).drop_duplicates('Outlet_code').set_index('Outlet_code')
    for ma, r in vs.iterrows():
        sch[int(ma)] = ({DM[c] for c in DM if r[c] == 'Y'}, str(r['Frequency']), str(r['ODD_WEEK']))
last_day = calendar.monthrange(today.year, today.month)[1]
def ok_week(par, d):
    w = d.isocalendar()[1]
    return par == 'Both' or (par == 'Odd Week' and w % 2 == 1) or (par == 'Even Week' and w % 2 == 0)
def visits(wds, par, lo, hi):
    out = []
    for dom in range(lo, hi + 1):
        d = today.replace(day=dom)
        if d.weekday() in wds and ok_week(par, d): out.append(dom)
    return out
dates = [{'dom': dom, 'name': WN[today.replace(day=dom).weekday()]} for dom in range(today.day, last_day + 1) if today.replace(day=dom).weekday() < 6]

kh = []
for r in k.itertuples(index=False):
    ma = int(r[0]); days = od.get(ma, [])
    if ma in sch:
        wds, freq, par = sch[ma]
    else:  # không có file lịch: suy từ cột Lịch viếng cũ, coi mọi tuần
        wds = {WN.index(x.strip()) for x in r[6].split(',')}; freq = 'F8' if len(wds) > 1 else 'F4'; par = 'Both'
    vd = visits(wds, par, today.day, last_day); pv = visits(wds, par, 1, today.day - 1)
    lich = ','.join(WN[i] for i in sorted(wds))
    cl0 = int(r[10]); lan = len(vd); cn = int(round(cl0 / lan)) if lan else 0   # CT ngày hôm nay: cố định theo số còn lại đầu ngày
    ta = tod.get(ma, 0)                                  # doanh số đã thực hiện hôm nay -> trừ vào còn lại tháng
    cl = max(cl0 - ta, 0); th_now = int(r[9]) + ta
    lan_f = lan - (1 if (ta > 0 and today.day in vd) else 0)   # KH đã mua hôm nay coi như xong lượt viếng hôm nay
    cf = int(round(cl / lan_f)) if lan_f else 0                # CT ngày/lần cho các ngày sau
    fnum = int(freq[1:]) if freq[1:].isdigit() else 4
    kh.append([ma, r[1], ddi[r[2]], r[3], r[4], r[5], lich, lan, int(r[8]), th_now, cl, cn,
               r[12] if isinstance(r[12], str) else '', r[13] if isinstance(r[13], str) else '', int(r[14]),
               r[15] if isinstance(r[15], str) else '',
               sum(x[1] for x in days), days, tod.get(ma, 0), tod_n.get(ma, 0), fnum, vd, pv, {'Both': 'B', 'Odd Week': 'O', 'Even Week': 'E'}.get(par, 'B'), cf])
S = {}
for r in s.itertuples(index=False):
    S.setdefault(str(int(r[0])), []).append([subi[r[1]], dmsi[r[2]], int(r[3]), int(r[4]), int(r[5])])

# --- Rate 3 tháng theo DDKD
r3 = []
if a.r3:
    b = pd.read_excel(a.r3, header=2)
    b = b[b['SS'].astype(str).str.contains('Nguyễn Văn Pháp')]
    p = b.groupby(['SM NAME', 'MONTH'])[['SO', 'TARGET SO']].sum().reset_index()
    for name in dd:
        row = {'dd': name}
        for mth in ['07', '08', '09']:
            x = p[(p['SM NAME'] == name) & (p['MONTH'] == mth + '/2026')]
            row['t' + mth] = int(x['SO'].sum()); row['c' + mth] = int(x['TARGET SO'].sum())
        r3.append(row)

data = json.dumps({'dd': dd, 'sub': sub, 'dms': dms, 'kh': kh, 's': S, 'r3': r3,
                   'today': today.isoformat(), 'todayName': today_name, 'hasToday': has_today,
                   'hr': hr, 'ordTo': ord_to, 'ordOut': n_out, 'dates': dates, 'hasSched': bool(sch)}, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
if a.password:
    import base64
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    salt, iv, it = os.urandom(16), os.urandom(12), 200000
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=it).derive(a.password.encode('utf-8'))
    ct = AESGCM(key).encrypt(iv, data.replace('<\\/', '</').encode('utf-8'), None)
    b = lambda x: base64.b64encode(x).decode()
    data = json.dumps({'enc': 1, 's': b(salt), 'i': b(iv), 'c': b(ct), 'n': it})
html = open(os.path.join(here, 'template.html'), encoding='utf-8').read().replace('__DATA__', data)
if a.standalone:
    cut = html.index('</style>') + len('</style>')
    fav = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%235b63e6'/%3E%3Cpath d='M8 24V14M14 24V8M20 24v-7M26 24H6' stroke='white' stroke-width='2.6' stroke-linecap='round' fill='none'/%3E%3C/svg%3E"
    head = ('<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<meta name="robots" content="noindex,nofollow"><meta name="theme-color" content="#0e0d11"><meta name="apple-mobile-web-app-capable" content="yes">'
            '<meta name="apple-mobile-web-app-title" content="SS Pháp"><link rel="icon" href="' + fav + '">'
            '<style>[hidden]{display:none!important}img{max-width:100%}:root{padding-bottom:env(safe-area-inset-bottom,0px);-webkit-text-size-adjust:100%}</style>')
    html = head + html[:cut] + '</head><body>' + html[cut:] + '</body></html>\n'

open(a.out, 'w', encoding='utf-8').write(html)
print(a.out, round(len(html) / 1e6, 2), 'MB', len(kh), 'KH', 'today', today, today_name, 'hasToday', has_today, 'ordTo', ord_to, 'KH có đơn', len(od), 'r3', len(r3))
