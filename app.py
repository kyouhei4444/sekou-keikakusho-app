# -*- coding: utf-8 -*-
"""施工計画書アプリ — ローカルサーバー。
ブラウザが画面、Excel COM（engine/*.ps1）が出力エンジン。

    python app.py           … 起動してブラウザを開く
    python app.py --noopen  … ブラウザを開かない
"""
import io
import json
import mimetypes
import os
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import build as builder      # noqa: E402
from core import config, excel         # noqa: E402

try:
    from core import soukatsu          # 段階2（無くても動く）
    from core import kikaishizai
except Exception:
    soukatsu = None
    kikaishizai = None

try:
    from core import shousa            # 設計照査（無くても動く）
except Exception:
    shousa = None

JOB = {'running': False, 'log': [], 'result': None, 'error': None, 'title': ''}
JOB_LOCK = threading.Lock()
CRLF = chr(13) + chr(10)
VERSION = '2026-10-05b'   # 画面右上に出す版。配布物の取り違え防止

# 雛形への差し込みではなく、設計書から読み取った明細で作る項目。
# 「作る項目」タブにも並べたいので、template_map.json とは別にここで定義する。
#   needs … 取り込みタブのどのデータが要るか（rows=工種・数量／kikai・zairyo=単価表）
EXTRA_ITEMS = [
    {'id': '01', 'label': '（1）工事概要', 'needs': 'rows'},
    {'id': '01-2', 'label': '（1-2）工事内容', 'needs': 'rows'},
    {'id': '05', 'label': '（5）主要機械', 'needs': 'kikai'},
    {'id': '06', 'label': '（6）主要資材', 'needs': 'zairyo'},
    {'id': 'shousa', 'label': '設計照査結果', 'needs': 'shousa'},
]
EXTRA_IDS = [x['id'] for x in EXTRA_ITEMS]


def jlog(s):
    with JOB_LOCK:
        JOB['log'].append(str(s))


def pick_folder(start=''):
    """フォルダー選択ダイアログ。
    ブラウザの裏に隠れないよう、最前面のオーナーウィンドウを作ってから開く。
    日本語や引用符でつまずかないよう、一時 .ps1（UTF-8 BOM）に書いて -File で呼ぶ。"""
    sel = start.replace("'", "''") if start else ''
    lines = [
        "Add-Type -AssemblyName System.Windows.Forms",
        # ダイアログがブラウザの裏に回らないよう、最前面の見えないウィンドウを親にする
        "$f = New-Object System.Windows.Forms.Form",
        "$f.TopMost = $true",
        "$f.ShowInTaskbar = $false",
        "$f.Opacity = 0",
        "$f.Width = 1",
        "$f.Height = 1",
        "$f.Show()",
        "$f.Activate()",
        "$d = New-Object System.Windows.Forms.FolderBrowserDialog",
        "$d.Description = 'フォルダーを選んでください'",
        "$d.ShowNewFolderButton = $true",
    ]
    if sel:
        lines.append("$d.SelectedPath = '%s'" % sel)
    lines += [
        "$r = $d.ShowDialog($f)",
        "$f.Close()",
        "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath }",
    ]
    ps = CRLF.join(lines) + CRLF
    fd, path = tempfile.mkstemp(suffix='.ps1', prefix='pick_')
    os.close(fd)
    with io.open(path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(ps)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
                            '-File', path], capture_output=True, timeout=300)
        out = excel._decode(r.stdout).strip()
        err = excel._decode(r.stderr).strip()
        if r.returncode != 0 or (err and not out):
            return {'error': 'フォルダー選択の画面を開けませんでした。'
                             '入力欄にパスを直接貼り付けてください。 ' + (err[:300] or '')}
        return {'folder': out}
    except Exception as e:
        return {'error': 'フォルダー選択の画面を開けませんでした（%s）。'
                         '入力欄にパスを直接貼り付けてください。' % e}
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


def drives():
    out = []
    for c in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
        d = c + ':' + os.sep
        if os.path.isdir(d):
            out.append(d)
    return out


def api_listdir(body):
    """アプリ内のフォルダー選択用。OSのダイアログはブラウザの裏に回るので使わない。"""
    p = (body.get('path') or '').strip().strip('"')
    if not p:
        p = config.load_settings().get('工事フォルダー親') or os.path.expanduser('~')
    p = os.path.abspath(p)
    if not os.path.isdir(p):
        up = os.path.dirname(p)
        if os.path.isdir(up):
            p = up
        else:
            return {'error': '見つかりません: %s' % p, 'path': '', 'parent': '',
                    'dirs': [], 'drives': drives()}
    parent = os.path.dirname(p)
    if parent == p:
        parent = ''
    dirs = []
    try:
        for n in sorted(os.listdir(p)):
            if n.startswith('$') or n.startswith('.'):
                continue
            full = os.path.join(p, n)
            if os.path.isdir(full):
                dirs.append({'name': n, 'path': full})
    except PermissionError:
        return {'error': 'このフォルダーは開けません（アクセス権がありません）',
                'path': p, 'parent': parent, 'dirs': [], 'drives': drives()}
    return {'path': p, 'parent': parent, 'dirs': dirs, 'drives': drives()}


def projects_dir():
    d = config.path('projects')
    os.makedirs(d, exist_ok=True)
    return d


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', '_', str(s)).strip() or 'no-name'


def item_order(item_id):
    """'01' '01-2' '09a' … を （0）（1）（1-2）…（16） の順に並べる。"""
    m = re.match(r'(\d+)(?:-(\d+))?([a-z]*)', str(item_id))
    if not m:
        return (99, 0, str(item_id))
    return (int(m.group(1)), int(m.group(2) or 0), m.group(3) or '')


# ---------------------------------------------------------------- API
def api_init():
    settings = config.load_settings()
    ok, ver = excel.excel_available()
    tmap = config.load_json('template_map.json', {})
    items = [{'id': f['id'], 'label': f['label'], 'copyOnly': bool(f.get('copyOnly')),
              'kankyo': f.get('kankyo', {}).get('items') if f.get('kankyo') else None}
             for f in tmap.get('files', [])]
    if soukatsu is not None and kikaishizai is not None:
        items += [dict(x) for x in EXTRA_ITEMS
                  if x['id'] != 'shousa' or shousa is not None]
    items.sort(key=lambda x: item_order(x['id']))
    return {
        'settings': settings,
        'staff': config.load_json('staff.json', {}).get('staff', []),
        'company': config.load_json('company.json', {}),
        'items': items,
        'excel': {'ok': ok, 'version': ver},
        'projects': sorted(os.path.splitext(x)[0] for x in os.listdir(projects_dir())
                           if x.endswith('.json')),
        'phase2': soukatsu is not None and kikaishizai is not None,
        'shousa': shousa is not None,
        'version': VERSION,
    }


def build_extra(prj, eids, body, files, warns):
    """設計書から読み取った明細で（1）（1-2）（5）（6）を作る。"""
    targets = [t for i, t in (('01', '1'), ('01-2', '1-2')) if i in eids]
    if targets:
        rows = body.get('rows') or []
        if rows:
            jlog('')
            jlog('■ 設計書の明細から（1）（1-2）（%d 明細）' % len(rows))
            r = soukatsu.make_files(prj, rows, targets, log=lambda s: jlog('  ' + s))
            if r.get('error'):
                warns.append(r['error'])
                jlog('  ! ' + r['error'])
            files += r.get('files', [])
        else:
            warns.append('（1）（1-2）は工種・数量が取り込まれていないので作りませんでした')
            jlog('  ! 工種・数量が未取込のため（1）（1-2）はとばしました')

    if '05' in eids or '06' in eids:
        kikai = (body.get('kikai') or []) if '05' in eids else []
        zairyo = (body.get('zairyo') or []) if '06' in eids else []
        if kikai or zairyo:
            jlog('')
            jlog('■ 設計書の明細から（5）（6）（機械 %d 件／資材 %d 件）' % (len(kikai), len(zairyo)))
            r = kikaishizai.make_files(prj, kikai, zairyo, log=lambda s: jlog('  ' + s))
            if r.get('error'):
                warns.append(r['error'])
                jlog('  ! ' + r['error'])
            files += r.get('files', [])
        else:
            warns.append('（5）（6）は機械・資材が取り込まれていないので作りませんでした')
            jlog('  ! 機械・資材が未取込のため（5）（6）はとばしました')


def build_shousa(prj, body, files, warns):
    """取り込みタブで行った設計照査の結果を 設計照査結果.xlsx にする。"""
    res = body.get('shousa')
    if not res:
        warns.append('設計照査結果は、照査がまだなので作りませんでした')
        jlog('  ! 設計照査が未実施のため とばしました')
        return
    jlog('')
    jlog('■ 設計照査結果')
    try:
        out = shousa.make_xlsx(prj, res, prj.get('出力先'), log=lambda s: jlog('  ' + s))
        files.append(os.path.basename(out))
    except Exception as e:
        warns.append('設計照査結果 の作成に失敗：%s' % e)
        jlog('  ! %s' % e)


def api_build(body):
    prj = body.get('project') or {}
    ids = body.get('ids')
    with JOB_LOCK:
        if JOB['running']:
            return {'error': 'いま別の生成が動いています'}
        JOB.update({'running': True, 'log': [], 'result': None, 'error': None,
                    'title': prj.get('工事名', '')})

    def work():
        try:
            eids = [i for i in (ids or []) if i in EXTRA_IDS]
            tids = None if ids is None else [i for i in ids if i not in EXTRA_IDS]
            files, warns, res = [], [], {}
            if tids is None or tids:
                res = builder.build(prj, ids=tids, log=jlog, verify=body.get('verify', True))
                files += res.get('files', [])
            if eids and soukatsu is not None and kikaishizai is not None:
                build_extra(prj, [i for i in eids if i != 'shousa'], body, files, warns)
                if 'shousa' in eids and shousa is not None:
                    build_shousa(prj, body, files, warns)
                jlog('')
                jlog('合計 %d ファイル' % len(files))
            with JOB_LOCK:
                JOB['result'] = {'folder': res.get('folder') or prj.get('出力先', ''),
                                 'files': files, 'images': res.get('images', []),
                                 'warnings': warns}
        except Exception as e:
            jlog('')
            jlog('■ エラー: %s' % e)
            with JOB_LOCK:
                JOB['error'] = str(e)
            traceback.print_exc()
        finally:
            with JOB_LOCK:
                JOB['running'] = False

    threading.Thread(target=work, daemon=True).start()
    return {'started': True}


def api_scan(body):
    if soukatsu is None:
        return {'error': '段階2の機能が読み込めていません'}
    return soukatsu.scan_folder(body.get('folder', ''))


def api_parse(body):
    if soukatsu is None:
        return {'error': '段階2の機能が読み込めていません'}
    return soukatsu.parse_naiyou(body.get('file', ''))


def api_make_naiyou(body):
    if soukatsu is None:
        return {'error': '段階2の機能が読み込めていません'}
    return soukatsu.make_files(body.get('project') or {}, body.get('rows') or [],
                               body.get('targets') or [], log=jlog)


def api_tanka(body):
    if kikaishizai is None:
        return {'error': '段階2の機能が読み込めていません'}
    return kikaishizai.parse_tanka(body.get('file', ''))


def api_make_kikai(body):
    if kikaishizai is None:
        return {'error': '段階2の機能が読み込めていません'}
    return kikaishizai.make_files(body.get('project') or {},
                                  body.get('kikai') or [], body.get('zairyo') or [], log=jlog)


def api_shousa(body):
    if shousa is None:
        return {'error': '設計照査の機能が読み込めていません'}
    return shousa.run(body.get('files') or [], body.get('project') or {})


def api_make_shousa(body):
    if shousa is None:
        return {'error': '設計照査の機能が読み込めていません'}
    prj = body.get('project') or {}
    if not prj.get('出力先'):
        return {'error': '出力先フォルダーが指定されていません'}
    try:
        out = shousa.make_xlsx(prj, body.get('result') or {}, prj['出力先'], log=jlog)
    except PermissionError:
        return {'error': '設計照査結果.xlsx を Excel で開いていると上書きできません。閉じてからやり直してください'}
    return {'files': [os.path.basename(out)], 'folder': prj['出力先']}


ROUTES_GET = {
    '/api/init': lambda q: api_init(),
    '/api/job': lambda q: dict(JOB),
    '/api/project': lambda q: json.load(io.open(
        os.path.join(projects_dir(), safe_name(q.get('name', [''])[0]) + '.json'), encoding='utf-8')),
}


def handle_post(path, body):
    if path == '/api/settings':
        return config.save_settings(body)
    if path == '/api/staff':
        return config.save_json('staff.json', {'staff': body.get('staff', [])})
    if path == '/api/company':
        return config.save_json('company.json', body)
    if path == '/api/project':
        name = safe_name(body.get('工事名') or body.get('_name') or '')
        with io.open(os.path.join(projects_dir(), name + '.json'), 'w', encoding='utf-8') as f:
            json.dump(body, f, ensure_ascii=False, indent=2)
        return {'saved': name}
    if path == '/api/build':
        return api_build(body)
    if path == '/api/listdir':
        return api_listdir(body)
    if path == '/api/pickfolder':
        return pick_folder(body.get('start', ''))
    if path == '/api/open':
        p = body.get('path', '')
        if os.path.exists(p):
            os.startfile(p)
            return {'opened': p}
        return {'error': '見つかりません: %s' % p}
    if path == '/api/scan':
        return api_scan(body)
    if path == '/api/parse':
        return api_parse(body)
    if path == '/api/naiyou':
        return api_make_naiyou(body)
    if path == '/api/tanka':
        return api_tanka(body)
    if path == '/api/kikai':
        return api_make_kikai(body)
    if path == '/api/shousa':
        return api_shousa(body)
    if path == '/api/shousa_xlsx':
        return api_make_shousa(body)
    return {'error': 'unknown api: %s' % path}


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False, default=str))

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        try:
            if u.path in ROUTES_GET:
                return self._json(ROUTES_GET[u.path](q))
            if u.path == '/api/image':
                p = q.get('path', [''])[0]
                if os.path.isfile(p):
                    with open(p, 'rb') as f:
                        return self._send(200, f.read(), 'image/png')
                return self._send(404, b'not found', 'text/plain')
            return self._static(u.path)
        except Exception as e:
            traceback.print_exc()
            return self._json({'error': str(e)}, 500)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        try:
            n = int(self.headers.get('Content-Length') or 0)
            body = json.loads(self.rfile.read(n).decode('utf-8')) if n else {}
            return self._json(handle_post(u.path, body))
        except Exception as e:
            traceback.print_exc()
            return self._json({'error': str(e)}, 500)

    def _static(self, path):
        rel = 'index.html' if path in ('/', '') else path.lstrip('/')
        p = config.path('web', *rel.split('/'))
        if not os.path.isfile(p):
            return self._send(404, b'not found', 'text/plain')
        ctype = mimetypes.guess_type(p)[0] or 'application/octet-stream'
        if ctype.startswith('text/') or ctype in ('application/javascript', 'application/json'):
            ctype += '; charset=utf-8'
        with open(p, 'rb') as f:
            return self._send(200, f.read(), ctype)


def main():
    s = config.load_settings()
    port = int(s.get('ポート', 8765))
    srv = None
    for p in range(port, port + 20):
        try:
            srv = ThreadingHTTPServer(('127.0.0.1', p), Handler)
            port = p
            break
        except OSError:
            continue
    if srv is None:
        print('空いているポートが見つかりませんでした')
        return
    url = 'http://127.0.0.1:%d/' % port
    print('施工計画書アプリ を起動しました  →  %s' % url)
    print('（終了するには、このウィンドウを閉じてください）')
    if '--noopen' not in sys.argv:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main()
