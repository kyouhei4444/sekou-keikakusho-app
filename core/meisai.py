# -*- coding: utf-8 -*-
"""見積参考資料の明細書をたどって、材料とその数量を拾う。

山梨県（および甲斐市）の見積参考資料は、次の入れ子になっている。

    本工事費内訳書   … 工種ごとの数量（例：軽量盛土 EPS工法 222ｍ３）
      └ 備考「第0016号工種明細書」
    工種明細書       … その工種の内訳。「計」が1単位あたり
      └ 備考「第0004号科目明細書」
    科目明細書       … 見出しの数量（222ｍ３）あたりの材料・作業
      └ 備考「第0 -0008号施工単価表」
    施工単価表       … 2通りある
                        ・「名称・規格など＋数量」… 1単位あたりの数量が載る（もう1段たどる）
                        ・「代表機労材＋構成比」  … 数量は載らない（材料名だけ）

数量は「計」の数量を基準に、親から渡された数量の倍率をかけて合計にする。
（例：現場打天端コンクリート打設は 10ｍ当り0.300ｍ３ → 28ｍ で 0.84ｍ３）

代表機労材しかない工種は数量が出ないので、その工種の数量を入れて
`estimate=True`（要確認）の印を付ける。
"""
import re
import unicodedata

try:
    import pymupdf
except Exception:                       # pragma: no cover
    pymupdf = None

MAX_DEPTH = 8

# 行の名称が作業（＝材料ではない）ことを示す語。名称だけを見る（規格は見ない）
WORK_WORDS = (
    '掘', '盛土', '埋戻', '運搬', '処理', '処分', '設置', '撤去', '据付', '打設',
    '敷均', '転圧', '整形', '整正', '整地', '清掃', '養生', '試験', '仮設', '復旧',
    '警備', '誘導', '土工', '積込', '切断', '破砕', '洗浄', '測量', '準備', '片付',
    '取壊', 'とりこわし', '殻', '再設置', '引抜', '建込', '削孔', '挿入', '定着',
    '練', '足場', '型枠', '張り', '張付', '塗装', '吹付', '締固', '補強', '溶接',
    '組立', '解体', '運転', '給油', '巡回', '点検', '休止', '保守', '移設',
    '路盤', '表層', '基層', '舗装', '伐開', '伐木', '伐採', '除根', '植生',
    '法面整形', '加工', '仕上', '路床', '修繕',
)

# 行そのものを捨てる語
JUNK_WORDS = ('計', '小計', '諸雑費', '積算単価', 'その他', '合計', '端数')

# 名称だけでなく規格も見て作業と判断する語（規格側に作業名が入る様式への手当て）
WORK_WORDS_SPEC = ('とりこわし', '取壊', '破砕', '切断', '積込', '運搬', '処分',
                   '掘削', '伐開', '除根', '清掃', '吹付', '撤去')

# 資材に載せない名称（燃料・経費・手間・機械の消耗品）
SKIP_WORDS = ('軽油', 'ガソリン', '灯油', '混合油', '重油', '計上', '利用料',
              '手間のみ', '諸雑費', '損料', 'ブレード', 'ビット', 'ドリルパイプ',
              'インナーロッド', 'ケーシング')

# 材料ではありえない単位（労務・機械・仮設・回数）
NG_UNITS = ('人', '人日', '日', '％', '%', '供用日', '月', '台日', '回', '台',
            '空m3', '掛m2', '空ｍ３', '掛ｍ２', '人･日', 'ヶ月', 'か月')

NUM = re.compile(r'^[0-9,]+(?:\.[0-9]+)?$')
CODE = re.compile(r'^[A-Z][A-Z0-9]{2,11}$')      # 積算コード（G1050・Y0H242005P2 など）
JP = re.compile(r'[ぁ-んァ-ヶ一-龥]')             # 品名には日本語が入る
NAME_NG = ('各種', '手間のみ', '同上', '標準', '一般')
REF_MEISAI = re.compile(r'第(\d{1,4})号(工種明細書|科目明細書)')
# 施工単価表の参照。「第0 -0010号施工単価表」が「第0 -0010号施工」で切れる資料もある
REF_TANKA = re.compile(r'第0-(\d{1,4})号')


def norm(s):
    return unicodedata.normalize('NFKC', str(s or '')).replace(' ', '').replace('　', '')


def _num(s):
    t = norm(s).replace(',', '')
    if not NUM.match(t):
        return None
    try:
        f = float(t)
    except ValueError:
        return None
    return int(f) if f == int(f) else f


# ---------------------------------------------------------------- ページを行に分ける
def _lines(page, tol=2.5):
    """語を y でまとめて行にする。二重に印字されている語は捨てる。"""
    ws = [w for w in page.get_text('words') if w[4].strip()]
    ws.sort(key=lambda w: (round(w[1], 1), w[0]))
    lines, cur, y = [], [], None
    for w in ws:
        if y is None or abs(w[1] - y) > tol:
            if cur:
                lines.append((y, cur))
            cur, y = [], w[1]
        if cur and abs(cur[-1][0] - w[0]) < 1.0 and cur[-1][4] == w[4]:
            continue
        cur.append(w)
    if cur:
        lines.append((y, cur))
    return lines


# 見出しの文字列と、それが表す列
LABELS = (('qty', ('数量',)), ('unit', ('単位',)), ('price', ('単価',)),
          ('amount', ('金額',)), ('biko', ('備考',)))


def _chars(words):
    """語を1文字ずつに割って x を振る。見出しは1文字ずつ離れて印字されるため。"""
    out = []
    for w in sorted(words, key=lambda w: w[0]):
        t = w[4]
        if not t:
            continue
        step = (w[2] - w[0]) / len(t)
        for i, ch in enumerate(t):
            if ch.strip():
                out.append((w[0] + step * i, w[0] + step * (i + 1), ch))
    return out


def _bounds(lines):
    """見出し行（名称／数量／単位／単価／金額／備考）から列の境界を決める。"""
    for y, words in lines[:14]:
        cs = _chars(words)
        s = ''.join(norm(c[2]) or '_' for c in cs)
        if len(s) != len(cs) or '数量' not in s or '備考' not in s:
            continue
        if '構成比' in s:             # 代表機労材のページは別の様式
            return None
        spans = {}
        for key, pats in LABELS:
            for pat in pats:
                i = s.find(pat)
                if i >= 0:
                    spans[key] = (cs[i][0], cs[i + len(pat) - 1][1])
                    break
        if 'qty' not in spans or 'biko' not in spans:
            continue
        left = [c[1] for c in cs if c[1] < spans['qty'][0]]
        if not left:
            continue
        spans['name'] = (cs[0][0], max(left))
        order = [k for k in ('name', 'qty', 'unit', 'price', 'amount', 'biko') if k in spans]
        edges = {}
        for a, b in zip(order, order[1:]):
            edges[b] = (spans[a][1] + spans[b][0]) / 2.0
        return {'y': y, 'cols': spans, 'edges': edges, 'order': order}
    return None


def _col_of(cx, bounds):
    """語の中心 x から、どの列かを決める。"""
    col = bounds['order'][0]
    for k in bounds['order'][1:]:
        if cx >= bounds['edges'][k]:
            col = k
        else:
            break
    return col


def _split(words, bounds):
    out = {k: [] for k in bounds['order']}
    qty_edge = bounds['edges'].get('qty', 1e9)
    for w in sorted(words, key=lambda w: w[0]):
        col = _col_of((w[0] + w[2]) / 2.0, bounds)
        # 長い名称は中心が数量列に入ってしまう。数字でなく左端が名称欄なら名称とみなす
        if col != 'name' and w[0] < qty_edge and not NUM.match(norm(w[4])):
            col = 'name'
        out[col].append(w)
    return out


# ---------------------------------------------------------------- 行（レコード）を組む
def page_recs(page):
    """1ページを「名称＋規格／数量／単位／備考」のレコードに分解する。

    名称と数量と備考は別の行に印字されるので、名称行が来たら前のレコードを閉じる。
    積算コード（G1050 など）だけの行は、直前のレコードに付ける。
    """
    lines = _lines(page)
    bounds = _bounds(lines)
    if not bounds:
        return []
    recs, cur = [], None

    def cond_of(words):
        """単価表の下に並ぶ条件行（例：B ｺﾝｸﾘｰﾄ規格 =1 18-8-25(20)BB）。"""
        t = [w[4].strip() for w in sorted(words, key=lambda w: w[0]) if w[4].strip()]
        if len(t) >= 4 and len(t[0]) == 1 and t[0].isalpha() and t[2].startswith('='):
            return t[1], ' '.join(t[3:])
        return None

    def flush():
        nonlocal cur
        if cur and (cur['names'] or cur['qty'] is not None):
            recs.append(cur)
        cur = None

    for y, words in lines:
        if y <= bounds['y']:
            continue
        c = cond_of(words)
        if c:
            flush()
            recs.append({'cond': c, 'names': [], 'qty': None, 'unit': '', 'ref': None,
                         'indent': None, 'code': '', 'y': y})
            continue
        cols = _split(words, bounds)
        name_words = cols.get('name') or []
        name_txt = ''.join(w[4] for w in name_words).strip()
        if name_txt.startswith('＊') or set(name_txt) == {'＊'}:
            flush()
            continue
        if name_words:
            # 積算コードの行は見出しに付く。ブロック番号を読んだ直後にだけ出る
            code_only = (bool(CODE.match(norm(name_txt))) and cur is not None
                         and bool(cur['ref']) and cur['qty'] is None)
            done = cur and (cur['qty'] is not None or cur['unit'] or cur['ref'])
            if cur is None or (done and not code_only):
                flush()
                cur = {'names': [], 'qty': None, 'unit': '', 'ref': None,
                       'indent': name_words[0][0], 'code': '', 'y': y}
            if code_only:
                cur['code'] = norm(name_txt)
            else:
                cur['names'].append(name_txt)
                if cur['indent'] is None:
                    cur['indent'] = name_words[0][0]
        qty_txt = ''.join(w[4] for w in (cols.get('qty') or [])).strip()
        if qty_txt:
            v = _num(qty_txt)
            if v is not None:
                if cur is None:
                    cur = {'names': [], 'qty': None, 'unit': '', 'ref': None,
                           'indent': None, 'code': '', 'y': y}
                if cur['qty'] is not None:
                    flush()
                    cur = {'names': [], 'qty': None, 'unit': '', 'ref': None,
                           'indent': None, 'code': '', 'y': y}
                cur['qty'] = v
        unit_txt = ''.join(w[4] for w in (cols.get('unit') or [])).strip()
        if unit_txt and cur is not None and not cur['unit']:
            cur['unit'] = unit_txt
        biko = norm(''.join(w[4] for w in (cols.get('biko') or [])))
        if biko and cur is not None and not cur['ref']:
            m = REF_MEISAI.search(biko)
            if m:
                cur['ref'] = ('工種' if m.group(2) == '工種明細書' else '科目') + m.group(1)
            else:
                m = REF_TANKA.search(biko)
                if m:
                    cur['ref'] = '単価' + m.group(1)
    flush()
    return recs


def rec_name(rec):
    return (rec['names'][0] if rec['names'] else '').strip()


def rec_spec(rec):
    return ' '.join(x.strip() for x in rec['names'][1:] if x.strip())


def is_junk(rec):
    n = norm(rec_name(rec))
    if not n:
        return True
    return any(n == w or n.startswith(w) for w in JUNK_WORDS)


# ---------------------------------------------------------------- 資料全体を読む
def read_book(pdf):
    """見積参考資料を読んで、内訳書の行と明細書のブロックを返す。"""
    doc = pymupdf.open(pdf)
    naiyou, blocks = [], {}
    cur = None                 # 明細書はページをまたいで続くので、ブロックを持ち越す
    for pi in range(doc.page_count):
        page = doc[pi]
        head = norm(page.get_text('text')[:400])
        is_naiyou = '本工事費内訳書' in head
        kind = None
        if '工種明細書' in head:
            kind = '工種'
        elif '科目明細書' in head:
            kind = '科目'
        elif '施工単価表' in head:
            kind = '単価'
        if not is_naiyou and kind is None:
            continue
        recs = page_recs(page)
        if not recs:
            continue
        if cur is not None and cur['kind'] != kind:
            cur = None            # 種類が変わったら持ち越さない
        if is_naiyou:
            for r in recs:
                r['page'] = pi + 1
            naiyou += recs
            continue
        # 明細書・単価表は「第NNNN号…」ごとのブロックに切る
        for r in recs:
            r['page'] = pi + 1
            if r.get('cond'):
                if cur is not None:
                    cur.setdefault('conds', {})[r['cond'][0]] = r['cond'][1]
                continue
            head_ref = r['ref'] and r['indent'] is not None and r['indent'] < 70
            if head_ref:
                bid = r['ref']
                cur = {'id': bid, 'kind': kind, 'name': rec_name(r), 'code': r['code'],
                       'qty': r['qty'], 'unit': r['unit'], 'page': pi + 1, 'items': []}
                blocks[bid] = cur
                continue
            if cur is None:
                continue
            cur['items'].append(r)
    tanka = read_tanka(doc)
    doc.close()
    for b in blocks.values():
        b['basis'] = _basis(b)
    return {'naiyou': naiyou, 'blocks': blocks, 'tanka': tanka}


TITLE_TANKA = re.compile(r'第0-(\d{1,4})号')


def read_tanka(doc):
    """代表機労材の施工単価表を号ごとに読む。数量は載っていないので名前だけ。

    明細書の作業行（例：下層路盤 58ｍ２）から、その単価表に入っている材料
    （再生クラッシャーラン など）を引くために使う。
    """
    from . import kikaishizai as kz
    out, cur = {}, None
    for pi in range(doc.page_count):
        page = doc[pi]
        if not kz.has_tanka(page.get_text('text')):
            continue
        cx = kz._code_x(page)
        nx0, nx1 = kz._name_x_range(page)
        if cx is None:
            continue
        for ws in kz._lines(page):
            joined = norm(''.join(w[4] for w in ws))
            if joined.startswith('＊＊＊'):
                continue
            code = next((w[4] for w in ws
                         if abs((w[0] + w[2]) / 2 - cx) < 30 and kz.CODE.match(w[4])), None)
            m = TITLE_TANKA.search(joined)
            if m and not code:      # 見出し行（明細の行には積算コードが付く）
                key = m.group(1).zfill(4)
                cur = out.setdefault(key, {'id': key, 'name': '', 'zairyo': [], 'kikai': []})
                head = [w[4] for w in ws
                        if w[0] < cx - 40 and '単価' not in norm(w[4]) and '第' not in w[4]]
                if head and not cur['name']:
                    cur['name'] = ''.join(head).strip()
                continue
            if cur is None or not code:
                continue
            area = sorted([w for w in ws if nx0 <= w[0] <= nx1], key=lambda w: w[0])
            left, right, gap = [], [], False
            for k, w in enumerate(area):
                if k and (w[0] - area[k - 1][2]) > 12:
                    gap = True
                (right if gap else left).append(w[4])
            if not left:
                continue
            kind = kz.classify(code)
            if kind not in ('機械', '材料'):
                continue
            row = {'name': ''.join(left), 'spec': ''.join(right), 'code': code}
            bag = cur['zairyo'] if kind == '材料' else cur['kikai']
            if not any(norm(x['name']) == norm(row['name'])
                       and norm(x['spec']) == norm(row['spec']) for x in bag):
                bag.append(row)
    return out


def _basis(block):
    """その明細書の行が「何単位あたり」かを決める。見出しの数量、無ければ「計」。"""
    if block.get('qty'):
        return block['qty']
    kei = None
    for r in block['items']:
        n = norm(rec_name(r))
        if n.startswith('計') and r['qty']:
            kei = r['qty']
        elif n.startswith('小計') and r['qty'] and kei is None:
            kei = r['qty']
    return kei or 1


# ---------------------------------------------------------------- 明細書をたどる
def _label(bid, block):
    if block['kind'] == '単価':
        return '第0-%s号施工単価表' % bid[2:]
    return '第%s号%s明細書' % (bid[2:], block['kind'])


def walk(book, out, bid, need, unit, trail, depth=0, seen=None):
    """ブロックをたどって、末端の行を out に積む。need は合計数量。"""
    blocks = book['blocks']
    b = blocks.get(bid)
    if b is None or depth > MAX_DEPTH:
        return
    seen = (seen or set()) | {bid}
    basis = b.get('basis') or 1
    scale = (need / basis) if (need and basis) else 1.0
    here = trail + ['%s %s' % (_label(bid, b), b['name'])]
    for r in b['items']:
        if is_junk(r):
            continue
        total = (r['qty'] * scale) if r['qty'] is not None else None
        ref = r['ref']
        nxt = blocks.get(ref) if ref else None
        if nxt is not None and ref not in seen:
            if nxt['kind'] in ('工種', '科目'):
                walk(book, out, ref, total, r['unit'], here, depth + 1, seen)
                continue
        spec = rec_spec(r)
        for k, v in (b.get('conds') or {}).items():      # 単価表の条件から規格を補う
            base = norm(k).replace('規格', '')
            if ('規格' not in norm(k) or not base
                    or base not in norm(rec_name(r))
                    or norm(v) in norm(spec + rec_name(r))):
                continue
            spec = (spec + ' ' + v).strip()
        row = {
            'name': rec_name(r), 'spec': spec, 'unit': r['unit'],
            'qty': total, 'raw': r['qty'], 'scale': scale,
            'basis': '%s %s%s当り' % (b['name'], _fmt(basis), b['unit'] or unit or ''),
            'trail': here, 'page': r.get('page'),
            'tanka': ref if ref and ref.startswith('単価') else '',
            'work': b['name'], 'via': _label(bid, b),
        }
        out.append(row)
        # 数量の入った施工単価表があれば、その中の材料まで降りる（明細書から1段だけ）
        if (nxt is not None and ref not in seen and nxt['kind'] == '単価'
                and b['kind'] != '単価'):
            row['expanded'] = True
            start = len(out)
            walk(book, out, ref, total, r['unit'], here, depth + 1, seen)
            if any(is_material(x) for x in out[start:]):
                row['superseded'] = True      # 内訳が取れたので、この行自体は出さない


def _fmt(v):
    if v is None:
        return ''
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return '{:,}'.format(v)


def is_material(row):
    """材料かどうか。名称に作業の語があれば材料ではない。"""
    n = norm(row.get('name'))
    if not n or row.get('qty') in (None, 0):
        return False
    if not JP.search(n) or n in NAME_NG:       # 規格だけの行を拾わない
        return False
    if any(w in n for w in JUNK_WORDS if len(w) > 1):
        return False
    if any(w in n + norm(row.get('spec')) for w in SKIP_WORDS):
        return False
    if norm(row.get('unit')) in [norm(u) for u in NG_UNITS]:
        return False
    if '材料費' in n:
        return True
    if '費' in n:                      # 処分費・リース費など
        return False
    if any(w in n for w in WORK_WORDS):
        return False
    both = n + norm(row.get('spec'))
    return not any(w in both for w in WORK_WORDS_SPEC)


def pack(rows):
    """同じ品名・規格・単位をまとめる。摘要（工種）と根拠は並べて残す。"""
    bag = {}
    for r in rows:
        key = (norm(r['name']), norm(r['spec']), norm(r['unit']))
        hit = bag.get(key)
        if hit is None:
            bag[key] = dict(r, qty=r['qty'], works=[r['work']], bases=[r['basis']],
                            trails=[' → '.join(r['trail'])])
            continue
        if r['qty'] is not None:
            hit['qty'] = (hit['qty'] or 0) + r['qty']
        if r['work'] not in hit['works']:
            hit['works'].append(r['work'])
        if r['basis'] not in hit['bases']:
            hit['bases'].append(r['basis'])
        t = ' → '.join(r['trail'])
        if t not in hit['trails']:
            hit['trails'].append(t)
    out = []
    for r in bag.values():
        q = r['qty']
        if isinstance(q, float):
            q = round(q, 3)
            if q == int(q):
                q = int(q)
        r['qty'] = q
        r['work'] = '／'.join(r['works'][:4])
        r['basis'] = '／'.join(r['bases'][:3])
        r['trail_text'] = r['trails'][0]
        out.append(r)
    return out


def collect(pdf):
    """内訳書 → 工種明細書 → 科目明細書 とたどって、材料と作業を拾う。"""
    if pymupdf is None:
        return {'error': 'pymupdf が入っていません'}
    book = read_book(pdf)
    if not book['blocks'] or not book['naiyou']:
        return {'error': '明細書が見つかりませんでした'}

    rows = []
    stack = []            # 内訳書の親（費目・工種・種別）
    entries = 0
    for r in book['naiyou']:
        if is_junk(r):
            continue
        name = rec_name(r)
        ind = r['indent']
        if ind is not None:
            stack = [s for s in stack if s[0] < ind - 2]
            stack.append((ind, name))
        if not r['ref'] or not r['ref'].startswith(('工種', '科目')):
            continue
        entries += 1
        trail = ['／'.join(s[1] for s in stack[:-1])] if len(stack) > 1 else []
        walk(book, rows, r['ref'], r['qty'], r['unit'], trail)

    direct = [x for x in rows if is_material(x) and not x.get('superseded')]
    works = [x for x in rows if not is_material(x)]
    extra = _from_tanka(book, works, direct)
    mats = pack(direct) + pack(extra)
    return {'rows': rows, 'materials': mats, 'works': pack(works),
            'direct': len(pack(direct)), 'tanka': len(pack(extra)),
            'blocks': len(book['blocks']), 'entries': entries,
            'naiyou': len(book['naiyou'])}


def _from_tanka(book, works, direct):
    """作業の行から、その施工単価表（代表機労材）に載っている材料を補う。

    単価表には数量が無いので、数量は**その作業の数量**を入れる（＝工種数量）。
    材料そのものの数量ではないので、要確認として印を付ける。
    """
    index = book.get('tanka') or {}
    have = {norm(x['name']) for x in direct}
    out = []
    for w in works:
        if w.get('expanded') or not w.get('tanka') or w.get('qty') in (None, 0):
            continue
        blk = index.get(w['tanka'][2:].zfill(4))
        if not blk:
            continue
        for z in blk['zairyo']:
            if norm(z['name']) in have:
                continue
            row = dict(w, name=z['name'], spec=z['spec'], code=z['code'], estimate=True,
                       basis='%s %s%s（工種数量）' % (w['name'], _fmt(w['qty']), w['unit']),
                       work=w['name'])
            if is_material(row):
                out.append(row)
    return out
