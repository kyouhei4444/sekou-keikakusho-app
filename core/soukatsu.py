# -*- coding: utf-8 -*-
"""段階2：設計図書の判別と、本工事費内訳書からの工種・数量の読み取り。

発注者ごとに様式が違うので、**列の位置を決め打ちしない**。
表の罫線から列を割り出し、見出し（費目／単位／数量…）で列の意味を決める。
階層は名称の字下げ（x座標）の順位で決める。
"""
import os
import re
import sys

from . import config, excel, util

try:
    import pymupdf
except Exception:
    pymupdf = None

LEVELS = ['B', 'C', 'D', 'E', 'G', 'G', 'G', 'G']
NUM = re.compile(r'^[0-9,]+(\.[0-9]+)?$')
ZEN = str.maketrans('０１２３４５６７８９．，－', '0123456789.,-')


def _clean(s):
    return re.sub(r'[\s　]+', ' ', str(s or '')).strip()


# ---------------------------------------------------------------- 図書の判別
KINDS = [
    ('工事数量総括表', ['工　事　数　量　総　括　表', '工事数量総括表']),
    ('公表用設計書', ['公 表 用 設 計 書', '公表用設計書']),
    ('特記仕様書', ['特記仕様書', '特 記 仕 様 書']),
    ('見積参考資料', ['見積参考資料', '施工単価表', '科目明細書']),
    ('数量計算書', ['数量計算書', '数 量 計 算 書']),
    ('設計書', ['設計書', '総括表']),
    ('入札公告', ['入札公告', '公告', '告示']),
    ('位置図', ['位置図']),
]

META_LABELS = {
    '工事名': ['工事名', '工 事 名'],
    '工事場所': ['工事場所', '施工場所', '工事箇所'],
    '事業名': ['事業名'],
    '発注者': ['発注元', '発注者'],
    '河川路線名': ['河川名、路線名等', '河川名', '路線名'],
}


def scan_folder(folder):
    if pymupdf is None:
        return {'error': 'pymupdf が入っていません（pip install pymupdf）'}
    if not folder or not os.path.isdir(folder):
        return {'error': 'フォルダーが見つかりません: %s' % folder}
    pdfs = []
    for root, dirs, names in os.walk(folder):
        for n in sorted(names):
            if n.lower().endswith('.pdf'):
                pdfs.append(os.path.join(root, n))
        if root != folder:
            dirs[:] = []          # 1階層下までにする
    if not pdfs:
        return {'error': 'PDFが見つかりません: %s' % folder}

    files, meta = [], {}
    for p in pdfs[:40]:
        try:
            d = pymupdf.open(p)
        except Exception as e:
            files.append({'name': os.path.basename(p), 'path': p, 'kind': '開けません',
                          'pages': 0, '内訳あり': False, 'error': str(e)})
            continue
        head = ''
        for i in range(min(3, d.page_count)):
            head += d[i].get_text('text')
        has_text = bool(head.strip())
        kind = '不明'
        for name, keys in KINDS:
            if any(k in head for k in keys):
                kind = name
                break
        # 施工単価表は資料の後ろの方にあることが多いので、全ページ見る（上限あり）。
        # 文字間隔の違いを吸収するため、空白を除いてから探す。
        from . import kikaishizai
        has_uchiwake = False
        has_tanka = False
        for i in range(min(400, d.page_count)):
            t = d[i].get_text('text')
            if not has_uchiwake:
                n = kikaishizai.norm(t)
                if '本工事費内訳書' in n or '費目・工種' in n:
                    has_uchiwake = True
            if not has_tanka and kikaishizai.has_tanka(t):
                has_tanka = True
            if has_uchiwake and has_tanka:
                break
        if not meta:
            meta = read_meta(d)
        if kind == '不明' and has_tanka:
            kind = '見積参考資料'
        files.append({'name': os.path.basename(p), 'path': p, 'kind': kind,
                      'pages': d.page_count, '内訳あり': has_uchiwake, '単価表あり': has_tanka,
                      '文字あり': has_text})
        d.close()
    files.sort(key=lambda f: (not f.get('内訳あり'), not f.get('単価表あり'),
                              f['kind'] == '不明', f['name']))
    return {'files': files, 'meta': meta}


def read_meta(doc):
    """先頭ページから工事名・場所などを拾う。ラベルの右側の文字を値にする。"""
    meta = {}
    for pi in range(min(3, doc.page_count)):
        for line in _lines(doc[pi]):
            words = sorted(line, key=lambda w: w[0])
            head = _clean(words[0][4]).replace(' ', '')
            rest = _clean(' '.join(w[4] for w in words[1:]))
            if not rest:
                continue
            for key, labels in META_LABELS.items():
                if key in meta:
                    continue
                if any(head.startswith(_clean(l).replace(' ', '')) for l in labels):
                    meta[key] = re.sub(r'^（当初）', '', rest).strip()
    return meta


# ---------------------------------------------------------------- 内訳書の解析
def _lines(page, tol=3.0):
    """単語を行ごとにまとめる。"""
    buckets = {}
    for w in page.get_text('words'):
        buckets.setdefault(round(w[1] / tol), []).append(w)
    return [buckets[k] for k in sorted(buckets)]


def _rules(page):
    """罫線（縦・横）の位置。細長い矩形も罫線として扱う。"""
    hs, vs = [], []
    for dr in page.get_drawings():
        for it in dr['items']:
            if it[0] == 'l':
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 0.8 and abs(a.x - b.x) > 50:
                    hs.append((a.y + b.y) / 2)
                if abs(a.x - b.x) < 0.8 and abs(a.y - b.y) > 20:
                    vs.append((a.x + b.x) / 2)
            elif it[0] == 're':
                r = it[1]
                if r.height < 1.5 and r.width > 50:
                    hs.append((r.y0 + r.y1) / 2)
                if r.width < 1.5 and r.height > 20:
                    vs.append((r.x0 + r.x1) / 2)
    return _cluster(hs), _cluster(vs)


def _cluster(vals, tol=3.0):
    out = []
    for v in sorted(vals):
        if not out or v - out[-1] > tol:
            out.append(v)
        else:
            out[-1] = (out[-1] + v) / 2
    return out


def _columns(page, vlines):
    """見出し行から列の意味（名称／単位／数量／摘要…）を決める。"""
    bounds = [0.0] + list(vlines) + [page.rect.width]
    for line in _lines(page):
        txt = _clean(' '.join(w[4] for w in sorted(line, key=lambda w: w[0]))).replace(' ', '')
        if '費目' not in txt:
            continue
        cols = {}
        for w in line:
            cx = (w[0] + w[2]) / 2
            for i in range(len(bounds) - 1):
                if bounds[i] <= cx < bounds[i + 1]:
                    cols.setdefault(i, {'x0': bounds[i], 'x1': bounds[i + 1], 'text': ''})
                    cols[i]['text'] += w[4]
                    break
        info = {'name_x1': None, 'qty': None, 'unit': None, 'notes': [], 'y': line[0][1]}
        for i, c in cols.items():
            t = c['text'].replace('　', '').replace(' ', '')
            cx = (c['x0'] + c['x1']) / 2
            if '費目' in t:
                info['name_x1'] = c['x1']
            elif '数量' in t or t == '数量':
                info['qty'] = cx
            elif '単位' in t:
                info['unit'] = cx
            elif '摘要' in t or '備考' in t:
                info['notes'].append(cx)      # 摘要も備考も階層のヒントになる
        if info['name_x1'] and (info['qty'] or info['unit']):
            return info
    return None


def parse_naiyou(pdf):
    """本工事費内訳書 → 明細の一覧。"""
    if pymupdf is None:
        return {'error': 'pymupdf が入っていません'}
    if not os.path.isfile(pdf):
        return {'error': 'ファイルがありません: %s' % pdf}
    doc = pymupdf.open(pdf)
    records, indents = [], []
    meta = read_meta(doc)
    pages_used = 0

    for pi in range(doc.page_count):
        page = doc[pi]
        text = page.get_text('text')
        if '本工事費内訳書' not in text and '費目・工種' not in text:
            continue
        hl, vl = _rules(page)
        col = _columns(page, vl)
        if not col or len(hl) < 3:
            continue
        pages_used += 1
        pitch = min((b - a) for a, b in zip(hl[:-1], hl[1:])) if len(hl) > 1 else 27.0
        bands = list(zip(hl[:-1], hl[1:]))
        bands.append((hl[-1], hl[-1] + pitch * 1.2))   # 最終行が罫線の外に出る様式への保険
        for y0, y1 in bands:
            rec = {'name': [], 'spec': '', 'qty': None, 'unit': '', 'note': '', 'x': None,
                   'page': pi + 1}
            for w in page.get_text('words'):
                cy = (w[1] + w[3]) / 2
                if not (y0 < cy <= y1):
                    continue
                cx = (w[0] + w[2]) / 2
                t = w[4]
                if cx < col['name_x1']:
                    rec['name'].append((w[1], w[0], t))
                    if rec['x'] is None or w[0] < rec['x']:
                        rec['x'] = w[0]
                else:
                    tgt = _nearest(cx, col)
                    if tgt == 'qty':
                        rec['qty'] = (rec['qty'] or '') + t
                    elif tgt == 'unit':
                        rec['unit'] += t
                    elif tgt == 'note':
                        rec['note'] += t
            if not rec['name'] and rec['qty'] is None and not rec['unit']:
                continue
            rec['name'], rec['spec'] = _split_name(rec['name'])
            if _is_junk(rec['name']):
                continue
            records.append(rec)
            if rec['x'] is not None and rec['name']:
                indents.append(round(rec['x'], 0))

    doc.close()
    if not records:
        return {'error': 'この資料からは内訳書の表を読み取れませんでした（画像PDFの可能性）'}

    # 名称だけの行と数量だけの行に分かれている様式（県）を1件にまとめる
    merged = []
    for r in records:
        if not r['name'] and merged and merged[-1]['qty'] is None:
            merged[-1]['qty'] = r['qty']
            merged[-1]['unit'] = merged[-1]['unit'] or r['unit']
            merged[-1]['note'] = merged[-1]['note'] or r['note']
        else:
            merged.append(r)

    steps = _cluster([i for i in indents], tol=3.0)
    offset = _calibrate(merged, steps)
    rows = []
    for r in merged:
        name = _clean(r['name'])
        if not name or '費目' in name or '内訳書' in name:
            continue
        star = name.startswith('＊') or name.startswith('*')
        name = re.sub(r'[＊*]+', '', name).strip()
        if not name:
            continue
        level = _level(star, r, steps, offset)
        qty, unit = _qty_unit(r)
        rows.append({'level': level, 'name': name, 'spec': _clean(r.get('spec')),
                     'qty': qty, 'unit': unit, 'page': r.get('page')})

    return {'rows': rows, 'meta': meta, 'source': '%d ページから読み取り' % pages_used}


JUNK = [re.compile(x) for x in (
    r'^\d+\s*頁$', r'^頁\s*[-‐−]?\s*\d*$', r'^[-‐−]?[\d,]+$',
    r'内訳書', r'費目', r'^（当初）', r'^第\s*\d+\s*号')]


def _is_junk(name):
    """ページ番号・見出し・ページ見返しなど、明細でない行。"""
    t = _clean(name).replace(' ', '')
    if not t:
        return False          # 数量だけの行は後で前の行に合流させる
    return any(p.search(t) for p in JUNK)


def _split_name(words):
    """名称セルの中身を 名称／規格 に分ける。
    2行目が1行目より右に出ていれば折り返し（名称の続き）、
    同じ字下げなら規格（市の様式）とみなす。"""
    if not words:
        return '', ''
    rows = {}
    for y, x, t in words:
        rows.setdefault(round(y / 3.0), []).append((x, t))
    keys = sorted(rows)
    def join(k):
        return ' '.join(t for _, t in sorted(rows[k]))
    def left(k):
        return min(x for x, _ in rows[k])
    name, spec = join(keys[0]), []
    base = left(keys[0])
    for k in keys[1:]:
        if left(k) > base + 4:
            name += join(k)          # 折り返し
        else:
            spec.append(join(k))     # 規格
    return name, ' '.join(spec)


def _nearest(cx, col):
    cands = [('qty', col.get('qty')), ('unit', col.get('unit'))]
    cands += [('note', c) for c in (col.get('notes') or [])]
    best, bd = None, 1e9
    for k, c in cands:
        if c is None:
            continue
        d = abs(cx - c)
        if d < bd:
            best, bd = k, d
    return best if bd < 60 else None


ORDER = ['B', 'C', 'D', 'E', 'G']


def _rank(x, steps):
    if x is None or not steps:
        return None
    return min(range(len(steps)), key=lambda i: abs(steps[i] - x))


def _calibrate(records, steps):
    """摘要欄の「費目行／工種行／種別行」から、字下げ順位と階層のずれを求める。
    市の様式のように途中の階層が無い内訳書でも、正しい階層に割り当てられる。"""
    votes = {}
    for r in records:
        lv = _note_level(r.get('note', ''))
        if not lv or lv == 'B':
            continue
        rk = _rank(r.get('x'), steps)
        if rk is None:
            continue
        d = ORDER.index(lv) - rk
        votes[d] = votes.get(d, 0) + 1
    if not votes:
        return 0
    return max(votes.items(), key=lambda kv: kv[1])[0]


def _note_level(note):
    if '費目行' in note:
        return 'B'
    if '工種行' in note:
        return 'D'
    if '種別行' in note:
        return 'E'
    return None


def _level(star, rec, steps, offset=0):
    if star:
        return 'B'
    lv = _note_level(rec.get('note', ''))
    if lv:
        return lv
    rk = _rank(rec.get('x'), steps)
    if rk is None:
        return 'G'
    return ORDER[max(0, min(rk + offset, len(ORDER) - 1))]


def _qty_unit(rec):
    q = _clean(rec['qty'] or '').translate(ZEN).replace(' ', '')
    u = _clean(rec['unit'] or '').replace(' ', '')
    if not q and u:
        m = re.match(r'^([０-９0-9]+)?(式|ｍ|m|ｍ２|ｍ３|箇所|本|枚|基|ｔ|t|組|個)$', u)
        if m:
            q = (m.group(1) or '1').translate(ZEN)
            u = m.group(2)
    if q and not NUM.match(q):
        m = re.search(r'[0-9,]+(\.[0-9]+)?', q)
        q = m.group(0) if m else ''
    if not q:
        return None, u
    try:
        v = float(q.replace(',', ''))
        return (int(v) if v == int(v) else v), u
    except Exception:
        return None, u


# ---------------------------------------------------------------- 出力
def make_files(prj, rows, targets, log=print):
    settings, ng = config.check()
    if ng:
        return {'error': '\n'.join(ng)}
    dst_dir = prj.get('出力先')
    if not dst_dir:
        return {'error': '出力先フォルダーが指定されていません'}
    os.makedirs(dst_dir, exist_ok=True)
    made = []

    if '1-2' in targets:
        try:
            made.append(_naiyou_xlsx(prj, rows, dst_dir, log))
        except Exception as e:
            return {'error': '（1-2）工事内容 の作成に失敗：%s' % e}
    if '1' in targets:
        try:
            made.append(_gaiyou_xls(prj, rows, dst_dir, settings, log))
        except Exception as e:
            return {'error': '（1）工事概要 の作成に失敗：%s' % e}
    return {'files': [os.path.basename(x) for x in made], 'folder': dst_dir}


def _naiyou_xlsx(prj, rows, dst_dir, log):
    import json
    out = os.path.join(dst_dir, '（1-2）工事内容.xlsx')
    items = []
    for r in rows:
        it = {'level': (r.get('level') or 'G').upper(), 'name': r.get('name', '')}
        if r.get('spec'):
            it['spec'] = r['spec']
        q = r.get('qty')
        if q not in (None, ''):
            try:
                v = float(str(q).replace(',', ''))
                it['qty'] = int(v) if v == int(v) else v
            except Exception:
                pass
        if r.get('unit'):
            it['unit'] = r['unit']
        items.append(it)
    payload = {'output': out, 'items': items,
               'template': config.path('data', 'templates', '（1-2）工事内容_雛形.xlsx')}
    p = excel.write_json(payload, 'naiyou')
    # exe 化すると sys.executable は自分自身になるので、別プロセスではなく同じプロセスで動かす
    import runpy
    argv = sys.argv
    try:
        sys.argv = ['build_koji_naiyou.py', p]
        try:
            runpy.run_path(config.path('engine', 'build_koji_naiyou.py'), run_name='__main__')
        except SystemExit as e:
            if e.code:
                raise RuntimeError('（1-2）の生成スクリプトが終了コード %s で止まりました' % e.code)
        log('（1-2）工事内容.xlsx  %d 明細' % len(items))
    finally:
        sys.argv = argv
        try:
            os.remove(p)
        except Exception:
            pass
    return out


def _gaiyou_xls(prj, rows, dst_dir, settings, log):
    """（1）工事概要.xls を雛形から作る。工期は日付シリアル、請負金額は空欄。"""
    src = os.path.join(settings['雛形フォルダー'], '（1）工事概要.xls')
    if not os.path.isfile(src):
        raise RuntimeError('雛形がありません: %s' % src)
    dst = os.path.join(dst_dir, '（1）工事概要.xls')

    gaiyou = [x for x in (prj.get('工事概要') or '').splitlines() if x.strip()][:6]
    cells = {
        'C4': prj.get('事業名') or None,
        'C7': prj.get('工事名') or None,
        'C10': prj.get('工事場所') or None,
        'E13': _serial(prj.get('工期自')),
        'E15': _serial(prj.get('工期至')),
        'D18': None,        # 請負金額は落札後に確定する
    }
    for i in range(6):
        cells['C%d' % (21 + i)] = gaiyou[i] if i < len(gaiyou) else None

    edits = [{'file': dst, 'sheet': '（1）工事概要', 'cells': cells}]

    mae = _maegaki(src, prj.get('工事名') or '')
    if mae:
        edits.append({'file': dst, 'sheet': 'まえがき', 'cells': {'B4': mae}})

    excel.run_ps('apply_edits.ps1', {'copies': [{'src': src, 'dst': dst}], 'edits': edits},
                 log=lambda s: log('  ' + s))
    return dst


def _serial(d):
    """Excel の日付シリアル値。"""
    dt = util.parse_date(d)
    if not dt:
        return None
    import datetime
    return (dt - datetime.date(1899, 12, 30)).days


def _maegaki(src, koji):
    """まえがき B4 の【工事名】だけ差し替える。"""
    if not koji:
        return None
    try:
        import xlrd
        t = xlrd.open_workbook(src).sheet_by_name('まえがき').cell_value(3, 1)
    except Exception:
        return None
    if not t or '【' not in t:
        return None
    return re.sub(r'【[^】]*】', '【%s】' % koji, str(t), count=1)
