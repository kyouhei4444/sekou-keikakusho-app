# -*- coding: utf-8 -*-
"""段階2：（5）主要機械・（6）主要資材の候補を設計図書から拾う。

**（6）主要資材**は見積参考資料の明細書をたどって拾う（`meisai.py`）。

    本工事費内訳書（工種と数量）
      → 備考「第0016号工種明細書」
    工種明細書（1単位あたりの内訳）
      → 備考「第0004号科目明細書」
    科目明細書（見出しの数量あたりの材料）
      → 備考「第0 -0008号施工単価表」

数量が出るのはこの経路だけなので、これを本筋とする。明細書が無い資料では、
従来の方法（代表機労材の材料名＋本工事費内訳書からの引き当て）に切り替える。

**（5）主要機械**は施工単価表の「代表機労材」から拾う。
積算コードの接頭で機械／労務／材料が分かる。

      M* K*  → 機械（損料・賃料）
      T*     → 材料（TSX は軽油・ガソリンなので主要資材には載せない）
      R* E*  → 労務・その他（対象外）

代表機労材が無い資料（甲斐市の設計書など）は、本文から機械名を拾う。
台数は資料に出ないので画面で入力する。
"""
import os
import re
import runpy
import sys
import unicodedata

from . import config, excel

try:
    import pymupdf
except Exception:
    pymupdf = None

CODE = re.compile(r'^[A-Z]{1,3}[0-9A-Z]{2,}$')
FUEL = ('TSX',)

# ②で使う機械名のことば。半角カナ・全角どちらでも当たるよう NFKC で正規化して探す
MACHINE_WORDS = [
    'ラフテレーンクレーン', 'トラッククレーン', 'クレーン機能付', 'クレーン付',
    'アスファルトフィニッシャ', 'コンクリートポンプ車', 'コンクリートカッタ',
    'ダンプトラック', 'モータグレーダ', 'ホイールローダ', 'タイヤローラ', 'ロードローラ',
    'クラムシェル', 'コンプレッサ', 'バックホウ', 'ブルドーザ', 'スクレーパ',
    '高所作業車', '振動ローラ', 'フィニッシャ', 'タンパ及びランマ',
    'クレーン', 'グレーダ', 'ショベル', 'トラック', 'ローラ', 'ランマ', 'タンパ',
    'カッタ', 'ブレーカ', 'ポンプ', 'ミキサ', 'ウインチ', '杭打機', 'オーガ',
    'バイブロ', '散水車', '発電機', '圧砕', '削岩', '積込機',
]
MACHINE_WORDS_SORTED = sorted(MACHINE_WORDS, key=len, reverse=True)

# 資材とみなす単位（②の資材候補で使う）
ZAIRYO_UNITS = ('ｍ', 'm', 'ｍ２', 'm2', 'ｍ３', 'm3', 'ｔ', 't', '枚', '本', '個', '組',
                '基', '箇所', 'kg', 'ｋｇ', 'L', 'ℓ')


def norm(s):
    """半角カナ・全角英数・空白の違いを吸収する。"""
    return unicodedata.normalize('NFKC', str(s or '')).replace(' ', '').replace('　', '')


def classify(code):
    m = re.match(r'^([A-Z]+)', code or '')
    if not m:
        return 'その他'
    pre = m.group(1)
    if pre in FUEL:
        return '燃料'
    if pre[0] in ('M', 'K'):
        return '機械'
    if pre[0] == 'T':
        return '材料'
    return 'その他'


def has_tanka(page_text):
    """施工単価表（代表機労材）のページか。文字間隔の違いを吸収して判定する。"""
    t = norm(page_text)
    return ('代表機労材' in t) or ('施工単価表' in t and '構成比' in t)


def _lines(page, tol=3.0):
    b = {}
    for w in page.get_text('words'):
        b.setdefault(round(w[1] / tol), []).append(w)
    return [sorted(b[k], key=lambda w: w[0]) for k in sorted(b)]


def _code_x(page):
    """積算コードが並ぶ列の x を、実際の出現位置から求める（様式差を吸収）。"""
    xs = [(w[0] + w[2]) / 2 for w in page.get_text('words') if CODE.match(w[4])]
    if len(xs) < 3:
        return None
    xs.sort()
    return xs[len(xs) // 2]      # 中央値


def _name_x_range(page):
    """代表機労材の名称が並ぶ列の範囲。見出し行から求め、無ければ既定値。"""
    for ws in _lines(page):
        t = norm(''.join(w[4] for w in ws))
        if '代表機労材' in t and '構成比' in t:
            head = [w for w in ws if '代表機労材' in norm(w[4])]
            kousei = [w for w in ws if '構成比' in norm(w[4])]
            if head and kousei:
                return head[0][0] - 20, kousei[0][0] - 10
    return 50.0, 95.0


def parse_tanka(pdf):
    """（5）（6）の候補を拾う。資材は明細書をたどり、機械は代表機労材から拾う。"""
    if pymupdf is None:
        return {'error': 'pymupdf が入っていません'}
    if not os.path.isfile(pdf):
        return {'error': 'ファイルがありません: %s' % pdf}
    doc = pymupdf.open(pdf)

    if not any(doc[i].get_text('text').strip() for i in range(min(5, doc.page_count))):
        doc.close()
        return {'error': 'この資料は文字が入っていないPDF（スキャン画像）なので読み取れません。'
                         '機械・資材は手入力してください。'}

    kikai, zairyo = {}, {}
    pages = 0
    koushu = ''          # 単価表はページをまたぐので、見出しを持ち越す
    for pi in range(doc.page_count):
        page = doc[pi]
        if not has_tanka(page.get_text('text')):
            continue
        pages += 1
        koushu = _scan_tanka_page(page, kikai, zairyo, koushu)
    km = _pack(kikai, '機械') if pages else _fallback_kikai(doc)
    doc.close()

    zm, msg_z = _zairyo_from_meisai(pdf)
    if not zm:
        # 明細書が読めない資料は、従来の方法（代表機労材＋内訳書の引き当て）
        zm = _pack(zairyo, '材料') if pages else _fallback_zairyo(pdf)
        rows = _naiyou_rows(pdf)
        linked, used = _link_qty(zm, rows)
        extra = _extra_zairyo(rows, used, zm)
        zm += extra
        msg_z = '明細書が読めないため、施工単価表の材料名に内訳書の数量を引き当て'
        if extra:
            msg_z += '／内訳書の材料 %d 件を追加' % len(extra)
        msg_z += '（数量は工種数量。要確認）'

    if not km and not zm:
        return {'error': 'この資料からは機械・資材を拾えませんでした。'
                         '施工単価表（代表機労材）も明細書も見つかりません。'
                         '別の資料（見積参考資料・設計書）を選ぶか、手入力してください。'}
    msg_k = ('%d ページの施工単価表（代表機労材）から機械' % pages if pages
             else '施工単価表が無いため、本文から機械名を拾った候補【要確認】')
    return {'kikai': km, 'zairyo': zm, 'fallback': pages == 0,
            'source': msg_k + '／資材は' + msg_z}


def _zairyo_from_meisai(pdf):
    """見積参考資料の明細書をたどって（6）主要資材を作る。"""
    try:
        from . import meisai
        m = meisai.collect(pdf)
    except Exception as e:
        return [], '明細書を読めませんでした（%s）' % e
    if m.get('error') or not m.get('materials'):
        return [], ''
    out = []
    for x in m['materials']:
        out.append({
            'name': x['name'], 'spec': x['spec'], 'unit': x['unit'], 'qty': x['qty'],
            'code': x.get('code', ''), 'work': x.get('work', ''),
            'basis': x.get('basis', ''), 'trail': x.get('trail_text', ''),
            'estimate': bool(x.get('estimate')),
        })
    est = sum(1 for x in out if x['estimate'])
    msg = '明細書（内訳書→工種明細書→科目明細書→施工単価表）をたどって %d 件' % len(out)
    if est:
        msg += ('。うち %d 件は単価表に数量が無いため、その工種の数量を入れた【要確認】' % est)
    return out, msg


def _scan_tanka_page(page, kikai, zairyo, koushu=''):
    cx = _code_x(page)
    nx0, nx1 = _name_x_range(page)
    if cx is None:
        return koushu
    cur = None
    for ws in _lines(page):
        joined = norm(''.join(w[4] for w in ws))
        if joined.startswith('＊＊＊') or joined.startswith('***'):
            cur = None
            continue
        if joined.endswith('単価表') or '号施工単価' in joined:
            head = [w[4] for w in ws
                    if w[0] < cx - 40 and '単価' not in norm(w[4]) and '第' not in w[4]]
            if head:
                koushu = ''.join(head)
            cur = None
            continue

        code = next((w[4] for w in ws if abs((w[0] + w[2]) / 2 - cx) < 30 and CODE.match(w[4])), None)
        # 名称欄は「名称＋規格」の1列。語の間が大きく空いたところで名称と規格に分ける
        area = sorted([w for w in ws if nx0 <= w[0] <= nx1], key=lambda w: w[0])
        left, right, gap = [], [], False
        for k, w in enumerate(area):
            if k and (w[0] - area[k - 1][2]) > 12:
                gap = True
            (right if gap else left).append(w[4])
        lx = area[0][0] if area else None

        if code and left:
            kind = classify(code)
            cur = {'name': ''.join(left), 'spec': ''.join(right), 'code': code, 'kind': kind,
                   'koushu': koushu}
            if kind in ('機械', '材料'):
                bag = kikai if kind == '機械' else zairyo
                key = norm(cur['name']) + '|' + norm(cur['spec'])
                if key in bag:
                    cur = bag[key]
                    if koushu and koushu not in cur['koushu']:
                        cur['koushu'] += '／' + koushu
                else:
                    bag[key] = cur
            continue

        if cur and (left or right) and lx is not None and lx >= nx0 + 5:
            add = ''.join(left) + ''.join(right)
            if add and add not in cur['spec']:
                cur['spec'] = (cur['spec'] + ' ' + add).strip()
    return koushu


def _pack(bag, kind):
    out = []
    for v in bag.values():
        if v['kind'] != kind:
            continue
        name, spec = v['name'], v['spec']
        if spec and name.endswith(spec):
            name = name[:-len(spec)].strip()      # 名称欄に規格まで入っている様式への手当て
        work = v['koushu']
        if '単価表' in norm(work) or '明細書' in norm(work):
            work = ''
        row = {'name': name, 'spec': spec, 'code': v.get('code', ''), 'work': work}
        if kind == '機械':
            s = name + ' ' + spec
            row['low_noise'] = '○' if '低騒音' in s else ''
            row['emission'] = '○' if '排ガス' in s else ''
            row['unit'] = '台'
            row['qty'] = 1
        else:
            row['unit'] = ''
            row['qty'] = ''
        out.append(row)
    return out


def _naiyou_rows(pdf):
    """同じ資料に入っている本工事費内訳書の明細。"""
    try:
        from . import soukatsu
        r = soukatsu.parse_naiyou(pdf)
    except Exception:
        return []
    return [x for x in r.get('rows', []) if x.get('qty') not in (None, '')]


def _is_zairyo_row(row):
    """内訳書の明細のうち、材料として（6）に載せてよさそうな行か。"""
    if (row.get('level') or '') not in ('E', 'F', 'G'):
        return False
    unit = str(row.get('unit') or '').strip()
    if not unit or unit in ('式', '人', '日', '人日', '台'):
        return False
    nm = str(row.get('name') or '')
    return not any(k in nm for k in WORK_WORDS)


def _extra_zairyo(rows, used, already):
    """単価表に出てこない材料（内訳書の細別）を足す。"""
    seen = {norm(z.get('name')) for z in already}
    out = []
    for r in rows:
        if not _is_zairyo_row(r):
            continue
        key = norm(r['name'])
        if key in seen:
            continue
        seen.add(key)
        out.append({'name': r['name'], 'spec': r.get('spec', ''), 'code': '',
                    'unit': r.get('unit', ''), 'qty': r.get('qty'),
                    'work': '', 'basis': '%s %s%s' % (r['name'], r['qty'], r.get('unit') or '')})
    return out


def _link_qty(zairyo, rows):
    """資材の数量を、同じ資料の本工事費内訳書から引き当てる。

    施工単価表には数量（構成比だけ）が無い。単価表の見出し＝その材料を使う細別なので、
    細別名で内訳書を引くと数量が分かる。**工種の数量**なので、材料そのものの数量とは
    限らない（例：生コンは「ブロック製作 86個」が当たる）。根拠を控えて要確認とする。
    """
    used = set()
    if not zairyo or not rows:
        return 0, used
    idx = {}
    for x in rows:
        idx.setdefault(norm(x['name']), x)

    n = 0
    for z in zairyo:
        works = [w for w in (z.get('work') or '').split('／') if w.strip()]
        hits = []
        for w in works:
            key = norm(w)
            hit = idx.get(key)
            if hit is None:
                cand = [v for k, v in idx.items() if key and (key in k or k in key)]
                hit = cand[0] if cand else None
            if hit and hit not in hits:
                hits.append(hit)
        if not hits:
            continue
        # 単位がそろっていれば合計、違えば先頭を使う
        units = {str(h.get('unit') or '') for h in hits}
        if len(units) == 1 and len(hits) > 1:
            try:
                z['qty'] = sum(float(h['qty']) for h in hits)
                if z['qty'] == int(z['qty']):
                    z['qty'] = int(z['qty'])
            except Exception:
                z['qty'] = hits[0]['qty']
        else:
            z['qty'] = hits[0]['qty']
        z['unit'] = hits[0].get('unit', '')
        z['basis'] = '／'.join('%s %s%s' % (h['name'], h['qty'], h.get('unit') or '') for h in hits)
        for h in hits:
            used.add(id(h))
        n += 1
    return n, used


# ---------------------------------------------------------------- ② 代表機労材が無い資料
def _machine_hit(token):
    """語の中に機械名があれば (名称, 規格) に分ける。
    名称はキーワードそのもの、規格はその後ろ（前に付く工種名は捨てる）。"""
    t = norm(token)
    for w in MACHINE_WORDS_SORTED:
        i = t.find(w)
        if i < 0:
            continue
        if w == 'ローラ' and i > 0 and t[i - 1] == 'ク':
            continue          # 「クローラ（履帯）」は機械名ではない
        spec = t[i + len(w):].strip('()（）[]［］{}・,、.。 　')
        if spec and not re.search(r'[0-9０-９]|型|式|級|吊', spec):
            spec = ''         # 規格らしくない文字列は捨てる
        return w, spec
    return None


NOT_KOUSHU = ('費目', '工種', '明細', '単価', '数量', '金額', '摘要', '備考', '合計', '計')
WORK_WORDS = ('掘削', '埋戻', '運搬', '処理', '処分', '設置', '撤去', '据付', '打設', '敷均',
              '転圧', '整形', '清掃', '養生', '試験', '仮設', '復旧', '警備', '誘導', '土工',
              '積込', '切断', '破砕', '洗浄', '測量', '準備', '片付', '取壊', '殻', '再設置',
              '引抜', '建込', '製作')


def _looks_like_koushu(line):
    """工種・細別の名前らしい行か（見出し行や数字行を除く）。"""
    if not line or len(line) > 30 or len(line) < 2:
        return False
    if any(k in line for k in NOT_KOUSHU):
        return False
    return bool(re.search(r'[ぁ-んァ-ヶ一-龥]', line))


def _fallback_kikai(doc):
    """本文から機械名を拾う。直前の行（細別名）を工種として添える。"""
    found = {}
    for pi in range(doc.page_count):
        page = doc[pi]
        if not any(norm(w) in norm(page.get_text('text')) for w in ('機', 'バックホウ')):
            pass
        prev = ''
        for ws in _lines(page):  # noqa: E501
            texts = [w[4] for w in ws]
            line = norm(''.join(texts))
            hit_any = False
            for tok in texts:
                for part in re.split(r'[+＋、,／/]', tok):
                    hit = _machine_hit(part)
                    if not hit:
                        continue
                    hit_any = True
                    name, spec = hit
                    if len(name) < 3:
                        continue
                    key = norm(name) + '|' + norm(spec)
                    rec = found.get(key)
                    if rec is None:
                        found[key] = {'name': name, 'spec': spec, 'koushu': prev, 'kind': '機械',
                                      'code': ''}
                    elif prev and prev not in rec['koushu']:
                        rec['koushu'] = (rec['koushu'] + '／' + prev).strip('／')
            if not hit_any and _looks_like_koushu(line):
                prev = ''.join(texts).strip()
    return _pack(found, '機械')


def _fallback_zairyo(pdf):
    """本工事費内訳書の明細のうち、単位が材料らしいものを資材候補にする。"""
    try:
        from . import soukatsu
        r = soukatsu.parse_naiyou(pdf)
    except Exception:
        return []
    if r.get('error'):
        return []
    out, seen = [], set()
    for row in r.get('rows', []):
        unit = str(row.get('unit') or '').strip()
        if not unit or unit not in ZAIRYO_UNITS:
            continue
        if (row.get('level') or '') not in ('E', 'F', 'G'):
            continue
        nm = str(row.get('name') or '')
        if any(k in nm for k in WORK_WORDS):
            continue          # 「管路掘削」などの作業は資材ではない
        key = norm(nm)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append({'name': row.get('name', ''), 'spec': row.get('spec', ''), 'code': '',
                    'unit': unit, 'qty': row.get('qty'), 'work': ''})
    return out


# ---------------------------------------------------------------- 出力
def make_files(prj, kikai, zairyo, log=print):
    dst_dir = prj.get('出力先')
    if not dst_dir:
        return {'error': '出力先フォルダーが指定されていません'}
    os.makedirs(dst_dir, exist_ok=True)
    made = []
    ryaku = (prj.get('工事略称') or '').strip()
    suffix = '（%s）' % ryaku if ryaku else ''

    if kikai:
        out = os.path.join(dst_dir, '（5）主要機械%s.xlsx' % suffix)
        payload = {
            'output': out,
            'note': '低騒音、排出ガス対策型の建設機械を使用する。',
            'machines': [{
                'name': r.get('name', ''), 'spec': r.get('spec', ''),
                'unit': r.get('unit') or '台', 'qty': _num(r.get('qty'), 1),
                'work': r.get('work', ''),
                'low_noise': r.get('low_noise', ''), 'emission': r.get('emission', ''),
            } for r in kikai],
            'notes': [
                '（5）主要機械　算出根拠・留意事項　――　%s' % (prj.get('工事名') or ''),
                '',
                '■ 出典',
                '・設計図書から抽出した機械の一覧である。',
                '■ 台数について【要確認】',
                '・台数は資料に出ないため、施工計画に合わせて入力した値である。',
                '■ 規格について【要確認】',
                '・容量クラスは積算コード内の定義でPDFに印字されないため、'
                '確認できる型式・排ガス／騒音区分のみ記載した。',
            ],
        }
        made.append(_run('build_shuyou_kikai.py', payload, out, log))

    if zairyo:
        out = os.path.join(dst_dir, '（6）主要資材%s.xlsx' % suffix)
        payload = {
            'output': out,
            'items': [{
                'name': r.get('name', ''), 'spec': r.get('spec', ''),
                'unit': r.get('unit', ''), 'qty': _num(r.get('qty'), None),
                'remark': r.get('work', ''),
            } for r in zairyo],
            'header': {
                '工事名': prj.get('工事名', ''), '発注者': prj.get('発注者', ''),
                '工事場所': prj.get('工事場所', ''),
                '様式': '「（6）主要資材.xls」に準拠',
            },
            'basis': [{
                'name': r.get('name', ''), 'spec': r.get('spec', ''),
                'formula': r.get('basis', ''),
                'qty': _num(r.get('qty'), None),
                'source': (r.get('trail') or '見積参考資料の明細書より')
                          + ('　※単価表に数量が無いため、その工種の数量【要確認】'
                             if r.get('estimate') else ''),
            } for r in zairyo if r.get('basis') or r.get('trail')],
            'notes': [
                '■ 出典',
                '・見積参考資料の明細書をたどって拾った材料の一覧である。',
                '　本工事費内訳書（工種・数量）→ 工種明細書 → 科目明細書 → 施工単価表。',
                '・軽油・ガソリン・手間のみ・諸雑費は除いた。',
                '■ 数量の出し方',
                '・明細書の行の数量に、親から渡した数量の倍率をかけて合計にした。',
                '　（例：現場打天端コンクリート打設は 10ｍ当り0.300ｍ３ × 28ｍ ＝ 0.84ｍ３）',
                '・算出根拠シートの「算出式」は、その行が何単位あたりかを示す。',
                '■ ※印の行について【要確認】',
                '・施工単価表が「代表機労材」だけで数量が無い工種は、',
                '　**その工種の数量**を入れてある（材料そのものの数量ではない）。',
                '　（例：下層路盤58ｍ２ → 再生クラッシャーラン 58ｍ２と表示）',
                '■ 不足分について【要確認】',
                '・図面の標準数量表・数量計算書にしかない材料（水抜きパイプ・目地材など）',
                '　は含まれていない。図面と突き合わせて足すこと。',
                '・標準単価・市場単価の工種は内訳が印字されないため、材料が出ないことがある。',
            ],
        }
        made.append(_run('build_shuyou_shizai.py', payload, out, log))

    if not made:
        return {'error': '明細がありません'}
    return {'files': [os.path.basename(x) for x in made], 'folder': dst_dir}


def _num(v, default):
    if v in (None, ''):
        return default
    try:
        f = float(str(v).replace(',', ''))
        return int(f) if f == int(f) else f
    except Exception:
        return default


def _run(script, payload, out, log):
    p = excel.write_json(payload, os.path.splitext(script)[0])
    argv = sys.argv
    try:
        sys.argv = [script, p]
        try:
            runpy.run_path(config.path('engine', script), run_name='__main__')
        except SystemExit as e:
            if e.code:
                raise RuntimeError('%s が終了コード %s で止まりました' % (script, e.code))
        log('%s' % os.path.basename(out))
    finally:
        sys.argv = argv
        try:
            os.remove(p)
        except Exception:
            pass
    return out
