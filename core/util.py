# -*- coding: utf-8 -*-
"""日付・文字列まわりの小物。"""
import datetime
import re

# 元号（施行日, 略号, 開始年）
_ERAS = [
    (datetime.date(2019, 5, 1), 'R', 2018),
    (datetime.date(1989, 1, 8), 'H', 1988),
    (datetime.date(1926, 12, 25), 'S', 1925),
]


def parse_date(s):
    """'2026-10-08' / '2026/10/8' / 'R8.10.8' を date にする。空なら None。"""
    if not s:
        return None
    if isinstance(s, datetime.date):
        return s
    s = str(s).strip()
    m = re.match(r'^([RHS])\s*(\d{1,2})[.\-/年]\s*(\d{1,2})[.\-/月]\s*(\d{1,2})', s)
    if m:
        mark, yy, mm, dd = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
        for _, k, base in _ERAS:
            if k == mark:
                return datetime.date(base + yy, mm, dd)
    m = re.match(r'^(\d{4})[.\-/年]\s*(\d{1,2})[.\-/月]\s*(\d{1,2})', s)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def wareki(d, with_day=True):
    """date → 'R8.10.8'（和暦）。"""
    d = parse_date(d)
    if not d:
        return ''
    for start, mark, base in _ERAS:
        if d >= start:
            y = d.year - base
            return '%s%d.%d.%d' % (mark, y, d.month, d.day) if with_day else '%s%d.%d' % (mark, y, d.month)
    return d.strftime('%Y.%m.%d')


def wareki_month(d):
    """date → 'R8.10月'（安全衛生管理計画表の月見出し）。"""
    d = parse_date(d)
    if not d:
        return ''
    for start, mark, base in _ERAS:
        if d >= start:
            return '%s%d.%d月' % (mark, d.year - base, d.month)
    return d.strftime('%Y.%m月')


def month_headers(start, end, count):
    """工期から月見出しを count 個作る。足りない分は空文字。"""
    s, e = parse_date(start), parse_date(end)
    out = []
    if s and e:
        y, m = s.year, s.month
        while len(out) < count:
            out.append(wareki_month(datetime.date(y, m, 1)))
            if (y, m) >= (e.year, e.month):
                break
            m += 1
            if m > 12:
                m = 1
                y += 1
    while len(out) < count:
        out.append('')
    return out[:count]


def kouki_hyouji(start, end):
    """'R8.10.8～R9.6.30'。"""
    a, b = wareki(start), wareki(end)
    return ('%s～%s' % (a, b)) if (a and b) else (a or b)


def blank(v):
    return v is None or str(v).strip() == ''
