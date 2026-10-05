# -*- coding: utf-8 -*-
"""設定とマスタの読み書き。exe化しても data/ engine/ web/ は exe の隣を見る。"""
import io
import json
import os
import sys

DEFAULT_SETTINGS = {
    "雛形フォルダー": "",
    "必携フォルダー": "",
    "工事フォルダー親": "",
    "ポート": 8765,
    "検証画像を作る": True
}


def app_dir():
    """exe（または app.py）が置かれているフォルダー。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def path(*parts):
    return os.path.join(app_dir(), *parts)


def _read(p, default):
    try:
        with io.open(p, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def _write(p, obj):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with io.open(p, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def guess_folders():
    """雛形・必携フォルダーを探す（初回起動時の自動設定）。

    **アプリの隣（配布フォルダー内）を最優先**にし、無ければ上のフォルダーへ遡る。
    配布セットには `施工計画書雛形/` を同梱しているので、通常はそれが選ばれる。
    """
    found = {}
    bases, b = [], app_dir()
    while True:
        bases.append(b)
        parent = os.path.dirname(b)
        if parent == b or len(bases) > 8:
            break
        b = parent

    for b in bases:
        for name, key in (('施工計画書雛形', '雛形フォルダー'), ('建設工事必携R7.10', '必携フォルダー')):
            p = os.path.join(b, name)
            if key not in found and os.path.isdir(p):
                found[key] = p

    # 工事フォルダーの親は「工事フォルダーが並んでいる場所」。
    # アプリの隣に雛形を同梱しているときに、配布フォルダー自身を指さないようにする。
    if '必携フォルダー' in found:
        found['工事フォルダー親'] = os.path.dirname(found['必携フォルダー'])
    else:
        for b in bases[1:]:
            if os.path.isdir(os.path.join(b, '施工計画書雛形')):
                found['工事フォルダー親'] = b
                break
        else:
            found['工事フォルダー親'] = os.path.dirname(app_dir())
    return found


def load_settings():
    s = dict(DEFAULT_SETTINGS)
    s.update(_read(path('settings.json'), {}))
    changed = False
    # 別のパソコンに移したとき、前のPCのパスが残っていると再検出できないので消す
    for k in ('雛形フォルダー', '必携フォルダー', '工事フォルダー親'):
        if s.get(k) and not os.path.isdir(s[k]):
            s[k] = ''
            changed = True
    for k, v in guess_folders().items():
        if not s.get(k):
            s[k] = v
            changed = True
    if changed:
        save_settings(s)
    return s


def save_settings(s):
    _write(path('settings.json'), s)
    return s


def load_json(name, default=None):
    """data/ のマスタを読む。無ければ見本（*.sample.json）を読む。
    社員名簿・協力会社は公開リポジトリに入れないので、初回は見本から始まる。
    画面で保存すると本物（staff.json など）が作られ、以後はそちらを読む。"""
    p = path('data', name)
    if not os.path.isfile(p):
        root, ext = os.path.splitext(name)
        p = path('data', root + '.sample' + ext)
    return _read(p, default if default is not None else {})


def save_json(name, obj):
    _write(path('data', name), obj)
    return obj


def staff_names():
    return [x.get('氏名', '') for x in load_json('staff.json', {}).get('staff', []) if x.get('氏名')]


def staff_tel(name):
    for x in load_json('staff.json', {}).get('staff', []):
        if x.get('氏名') == name:
            return x.get('電話', '')
    return ''


def check():
    """使える状態か点検して、足りないものを返す。"""
    s = load_settings()
    ng = []
    if not s.get('雛形フォルダー') or not os.path.isdir(s['雛形フォルダー']):
        ng.append('雛形フォルダーが見つかりません（設定タブで指定してください）')
    if not os.path.isdir(path('engine')):
        ng.append('engine フォルダーがありません')
    return s, ng
