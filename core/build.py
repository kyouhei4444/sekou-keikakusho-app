# -*- coding: utf-8 -*-
"""段階1：雛形に工事データを差し込んで施工計画書一式を作る。"""
import os
import re
import shutil

from . import config, excel, util

PLACEHOLDER = re.compile(r'^\{([^{}]+)\}$')


# ---------------------------------------------------------------- 差し込み値
def make_values(prj):
    """工事データから {…} の置き換え表を作る。無い値は None（＝セルを消す）。"""
    t = prj.get('体制', {}) or {}
    kensa = (t.get('社内検査員') or []) + ['', '', '', '']
    kantoku = (prj.get('監督員') or []) + [{}, {}, {}]

    def s(v):
        v = '' if v is None else str(v).strip()
        return v or None

    def tel(name, given=''):
        return s(given) or s(config.staff_tel(name or ''))

    dairi = t.get('現場代理人', '')
    shunin = t.get('主任技術者', '')

    v = {
        '工事名': s(prj.get('工事名')),
        '会社名': s(prj.get('会社名')),
        '工事場所': s(prj.get('工事場所')),
        '発注者': s(prj.get('発注者')),
        '発注者課': s(prj.get('発注者課')),
        '工期表示': s(util.kouki_hyouji(prj.get('工期自'), prj.get('工期至'))),
        '現場代理人': s(dairi),
        '現場代理人TEL': tel(dairi, t.get('現場代理人TEL', '')),
        '主任技術者': s(shunin),
        '主任技術者TEL': tel(shunin, t.get('主任技術者TEL', '')),
    }
    v['発注者改行入り'] = s('\n'.join([x for x in [v['発注者'], v['発注者課']] if x]))

    for k in ('事務担当', '品質管理担当', '出来形管理担当', '工程管理担当', '安全管理担当',
              '統括安全衛生責任者', '安全管理者', '衛生管理者', '安全巡視員',
              '火気取締責任者', '電気取扱責任者', '車両取扱責任者', '機械取扱責任者',
              '作業指揮者1名称', '作業指揮者1', '作業指揮者2名称', '作業指揮者2',
              '事故対策指揮者', '被災者救出係', '事故現場指揮者', '現場交通管理係',
              '事故原因調査係', '情報連絡係', '災害対策係'):
        v[k] = s(t.get(k))
    for i in range(4):
        v['社内検査員%d' % (i + 1)] = s(kensa[i])
    for i in range(3):
        v['監督員%d' % (i + 1)] = s((kantoku[i] or {}).get('氏名'))

    # （4）指定機械
    sitei = prj.get('指定機械', {}) or {}
    if sitei.get('あり'):
        v['指定機械文言'] = s(sitei.get('文言'))
    else:
        v['指定機械文言'] = '%s において指定機械はありません。' % (v['工事名'] or '本工事')

    # （9）安全方針
    hoshin = (prj.get('安全方針') or []) + ['', '', '', '']
    for i in range(4):
        v['安全方針%d' % (i + 1)] = s(hoshin[i])

    # （11）搬出先
    han = prj.get('搬出先', {}) or {}
    for key, tag in (('Co、As殻', 'Co殻'), ('汚泥', '汚泥'), ('残土', '残土'), ('残土（2）', '残土2')):
        d = han.get(key, {}) or {}
        v['搬出先_%s_処分先' % tag] = s(d.get('処分先'))
        v['搬出先_%s_所在地' % tag] = s(d.get('所在地'))
        v['搬出先_%s_距離' % tag] = s(d.get('距離'))
        v['搬出先_%s_内容' % tag] = s(d.get('内容'))
    return v


def resolve(text, values):
    """'{工事名}' → 値。文字列中の {…} も置き換える。全体が未設定なら None。"""
    if text is None:
        return None
    t = str(text)
    m = PLACEHOLDER.match(t)
    if m:
        return values.get(m.group(1))
    out = re.sub(r'\{([^{}]+)\}', lambda x: values.get(x.group(1)) or '', t)
    # 字下げ（全角空白）は雛形の体裁なので strip しない。中身が空のときだけ None
    return out if out.strip() else None


def kousei_gaisha(prj):
    """下請＋委託をまとめた構成会社一覧。"""
    rows = []
    for x in (prj.get('下請') or []):
        if str(x.get('会社', '')).strip():
            rows.append({'会社': x.get('会社'), '担当': x.get('担当'), '電話': x.get('電話')})
    for x in (prj.get('委託') or []):
        if str(x.get('会社', '')).strip():
            rows.append({'会社': x.get('会社'), '担当': x.get('担当'), '電話': x.get('電話')})
    return rows


# ---------------------------------------------------------------- 生成
def build(prj, ids=None, log=print, verify=True):
    settings, ng = config.check()
    if ng:
        raise excel.EngineError('\n'.join(ng))

    src_dir = settings['雛形フォルダー']
    dst_dir = prj.get('出力先') or ''
    if not dst_dir:
        raise excel.EngineError('出力先フォルダーが指定されていません')
    os.makedirs(dst_dir, exist_ok=True)

    tmap = config.load_json('template_map.json', {})
    files = [f for f in tmap.get('files', []) if (ids is None or f['id'] in ids)]
    if not files:
        raise excel.EngineError('作る項目が選ばれていません')

    values = make_values(prj)
    lists = {
        '下請': [x for x in (prj.get('下請') or []) if str(x.get('会社', '')).strip() or str(x.get('工種', '')).strip()],
        '委託': [x for x in (prj.get('委託') or []) if str(x.get('会社', '')).strip() or str(x.get('種別', '')).strip()],
        '構成会社': kousei_gaisha(prj),
    }

    copies, edits = [], []
    made = []

    for f in files:
        src = os.path.join(src_dir, f['file'])
        dst = os.path.join(dst_dir, f['file'])
        if not os.path.isfile(src):
            log('  ! 雛形がありません: %s（とばします）' % f['file'])
            continue
        copies.append({'src': src, 'dst': dst})
        made.append(f)

        for e in f.get('edits', []):
            cells = {}
            for addr, tpl in (e.get('cells') or {}).items():
                cells[addr] = resolve(tpl, values)

            # 繰り返し（下請・委託・構成会社）
            for rep in (e.get('repeat') or []):
                rows = lists.get(rep.get('list'), [])
                for i, slot in enumerate(rep.get('slots', [])):
                    row = rows[i] if i < len(rows) else None
                    for field, addr in slot.items():
                        if field.startswith('_'):
                            continue
                        val = None
                        if row is not None:
                            val = row.get(field)
                            val = str(val).strip() if val is not None else ''
                            val = val or None
                        cells[addr] = val

            # 月見出し（安全衛生管理計画表）
            mon = e.get('months')
            if mon:
                cols = mon.get('cols', [])
                hdr = util.month_headers(prj.get('工期自'), prj.get('工期至'), len(cols))
                for c, h in zip(cols, hdr):
                    cells['%s%d' % (c, mon.get('row', 5))] = h or None

            for addr in (e.get('clear') or []):
                cells[addr] = None

            item = {'file': dst, 'sheet': e.get('sheet'), 'cells': cells}
            if e.get('text'):
                item['text'] = e['text']
            edits.append(item)

        # （13）現場環境改善
        k = f.get('kankyo')
        if k:
            chosen = set(prj.get('現場環境改善') or [])
            cells = {}
            for it in k.get('items', []):
                key = '%d' % it['row']
                mark = k.get('suffix', '') if (key in chosen or it['text'] in chosen) else ''
                cells['%s%d' % (k.get('column', 'D'), it['row'])] = it['text'] + mark
            edits.append({'file': dst, 'sheet': k.get('sheet'), 'cells': cells})

    if not made:
        raise excel.EngineError('雛形が1つも見つかりませんでした。設定の雛形フォルダーを確認してください')

    log('■ 雛形のコピーと差し込み（%d ファイル）' % len(made))
    excel.run_ps('apply_edits.ps1', {'copies': copies, 'edits': edits}, log=lambda s: log('  ' + s))

    # ---- 氏名のドロップダウン
    names = config.staff_names()
    dfiles = []
    for f in made:
        d = f.get('dropdown')
        if not d:
            continue
        entry = {
            'file': os.path.join(dst_dir, f['file']),
            'listSheet': d['listSheet'],
            'listCol': d['listCol'],
            'colWidth': 18,
            'dropdowns': [{'sheet': d['listSheet'], 'cells': d['cells']}],
        }
        if d.get('printArea'):
            entry['printArea'] = d['printArea']
        if d.get('cells2'):
            entry['dropdowns'].append({'sheet': d['cells2']['sheet'], 'cells': d['cells2']['cells']})
        dfiles.append(entry)
    if dfiles and names:
        log('■ 氏名のドロップダウン（%d ファイル・%d 名）' % (len(dfiles), len(names)))
        excel.run_ps('set_dropdown.ps1', {'names': names, 'files': dfiles}, log=lambda s: log('  ' + s))

    # ---- 組織表の系統線
    trees = []
    for f in made:
        t = f.get('tree')
        if not t:
            continue
        n = len(lists['下請'])
        kids = t['children'][:n]
        if not kids:
            continue
        trees.append({
            'file': os.path.join(dst_dir, f['file']),
            'sheet': t['sheet'],
            'deleteNamePrefix': '下請業者線',
            'trees': [{'name': '下請業者線', 'parent': t['parent'], 'children': kids,
                       'color': t.get('color', 'blue'), 'weight': t.get('weight', 1.25)}],
        })
    if trees:
        log('■ 下請業者の系統線')
        excel.run_ps('draw_tree.ps1', {'files': trees}, log=lambda s: log('  ' + s))

    # ---- 印刷体裁
    ffiles = []
    for f in made:
        if f.get('fit'):
            ffiles.append({'file': os.path.join(dst_dir, f['file']), 'sheets': f['fit']})
    if ffiles:
        log('■ 印刷体裁')
        excel.run_ps('fit_print.ps1', {'files': ffiles}, log=lambda s: log('  ' + s))

    # ---- 検証用PDF
    images = []
    if verify and settings.get('検証画像を作る', True):
        pdf_dir = os.path.join(dst_dir, '_確認用')
        jobs = []
        for f in made:
            for sh in (f.get('fit') or []):
                jobs.append({'file': os.path.join(dst_dir, f['file']), 'sheet': sh['sheet'],
                             'name': f['id'] + '_' + re.sub(r'[\\/:*?"<>|]', '_', str(sh['sheet']))})
        if jobs:
            log('■ 確認用PDFの書き出し')
            excel.run_ps('to_pdf.ps1', {'out': pdf_dir, 'jobs': jobs}, log=lambda s: log('  ' + s))
            images = to_png(pdf_dir, log)

    log('')
    log('完了：%s に %d ファイル' % (dst_dir, len(made)))
    return {'folder': dst_dir, 'files': [f['file'] for f in made], 'images': images}


def to_png(pdf_dir, log=print):
    """確認用PDFをPNGにする（画面に出して目視するため）。"""
    try:
        import pymupdf
    except Exception:
        log('  ! pymupdf が無いので画像化はとばしました')
        return []
    out = []
    for name in sorted(os.listdir(pdf_dir)):
        if not name.lower().endswith('.pdf'):
            continue
        try:
            doc = pymupdf.open(os.path.join(pdf_dir, name))
            for i, page in enumerate(doc):
                dst = os.path.join(pdf_dir, '%s_p%d.png' % (os.path.splitext(name)[0], i + 1))
                page.get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6)).save(dst)
                out.append(dst)
                if i >= 2:
                    break
            doc.close()
        except Exception as e:
            log('  ! 画像化できませんでした: %s (%s)' % (name, e))
    return out


def copy_only_list():
    tmap = config.load_json('template_map.json', {})
    return [f for f in tmap.get('files', []) if f.get('copyOnly')]
