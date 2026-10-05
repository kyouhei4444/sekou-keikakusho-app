# -*- coding: utf-8 -*-
"""Excel 操作は engine/ の PowerShell スクリプトに任せる。
.xls を編集できるのは Excel COM だけなので、Python から直接は扱わない。"""
import io
import json
import os
import subprocess
import tempfile

from . import config


class EngineError(Exception):
    pass


def _decode(b):
    if not b:
        return ''
    for enc in ('cp932', 'utf-8'):
        try:
            return b.decode(enc)
        except Exception:
            pass
    return b.decode('cp932', 'replace')


def write_json(obj, prefix='job'):
    fd, p = tempfile.mkstemp(suffix='.json', prefix=prefix + '_')
    os.close(fd)
    with io.open(p, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    return p


def run_ps(script, payload=None, args=None, log=None):
    """engine/<script> を実行する。payload があれば -Json で渡す。"""
    exe = config.path('engine', script)
    if not os.path.isfile(exe):
        raise EngineError('スクリプトがありません: %s' % exe)
    cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', exe]
    tmp = None
    if payload is not None:
        tmp = write_json(payload, os.path.splitext(script)[0])
        cmd += ['-Json', tmp]
    if args:
        cmd += list(args)
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=1800)
    finally:
        if tmp and os.path.isfile(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass
    out = _decode(r.stdout).strip()
    err = _decode(r.stderr).strip()
    if log is not None:
        for line in out.splitlines():
            log(line)
    if r.returncode != 0:
        raise EngineError('%s が失敗しました\n%s\n%s' % (script, out, err))
    if err and log is not None:
        for line in err.splitlines():
            log('  ! ' + line)
    return out


def excel_available():
    """Excel（COM）が使えるか。"""
    ps = ('try { $x = New-Object -ComObject Excel.Application; $v = $x.Version; '
          '$x.Quit(); Write-Output $v } catch { Write-Output "NG" }')
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-Command', ps],
                           capture_output=True, timeout=120)
        v = _decode(r.stdout).strip()
        return (v and v != 'NG'), v
    except Exception as e:
        return False, str(e)
