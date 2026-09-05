#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Считает рабочие дни по производственному календарю РФ.

Примеры:
    python3 scripts/workdays.py 2026-09-15 +5      -> дата через 5 рабочих дней
    python3 scripts/workdays.py 2026-11-20 -5      -> дата за 5 рабочих дней до
    python3 scripts/workdays.py 2026-09-15 2026-11-20   -> сколько рабочих дней между

Календарь берётся из файла scripts/календарь.txt — его можно править руками.
"""
import sys, os, datetime as dt

CAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'календарь.txt')

def _load():
    holidays, workdays = set(), set()
    if os.path.exists(CAL):
        for line in open(CAL, encoding='utf-8'):
            line = line.split('#')[0].strip()
            if not line:
                continue
            if line.startswith('рабочий '):
                workdays.add(line.split()[1])
            else:
                holidays.add(line)
    return holidays, workdays

HOLIDAYS, EXTRA_WORKDAYS = _load()

def is_workday(d):
    s = d.isoformat()
    if s in EXTRA_WORKDAYS:
        return True
    if s in HOLIDAYS:
        return False
    return d.weekday() < 5

def add(d, n):
    """+n рабочих дней (n>0) или -n (n<0). День отсчёта не считается."""
    step = 1 if n >= 0 else -1
    left = abs(n)
    while left:
        d += dt.timedelta(days=step)
        if is_workday(d):
            left -= 1
    return d

def between(a, b):
    """Сколько рабочих дней от a до b (не включая a, включая b)."""
    n, d = 0, a
    while d < b:
        d += dt.timedelta(days=1)
        if is_workday(d):
            n += 1
    return n

def parse(s):
    s = s.strip().replace('/', '.').replace('-', '.')
    parts = s.split('.')
    if len(parts[0]) == 4:
        y, m, day = parts
    else:
        day, m, y = parts
    return dt.date(int(y), int(m), int(day))

def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    d = parse(sys.argv[1])
    arg = sys.argv[2]
    if arg.startswith(('+', '-')):
        print(add(d, int(arg)).isoformat())
    else:
        print(between(d, parse(arg)))

if __name__ == '__main__':
    main()
