#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Обновляет производственный календарь из открытых данных.

    python3 scripts/update_calendar.py 2026
    python3 scripts/update_calendar.py 2026 2027

Источник — репозиторий xmlcalendar (https://github.com/xmlcalendar/data), тот же календарь,
что публикуется на xmlcalendar.ru: официальные праздники и переносы выходных по годам.
Скрипт переписывает блок нужного года в scripts/календарь.txt, остальные годы не трогает.

Нужен интернет. Если его нет — календарь можно править руками, формат описан в самом файле.
"""
import sys, os, re, subprocess

URL = 'https://raw.githubusercontent.com/xmlcalendar/data/master/ru/%s/calendar.xml'
CAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'календарь.txt')


def download(year):
    url = URL % year
    try:
        import urllib.request
        with urllib.request.urlopen(url, timeout=20) as r:
            return r.read().decode('utf-8')
    except Exception:
        out = subprocess.run(['curl', '-sS', '--max-time', '20', url],
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if out.returncode or not out.stdout:
            raise SystemExit('Не удалось скачать календарь на %s год.\n%s'
                             % (year, out.stderr.decode('utf-8', 'ignore').strip()))
        return out.stdout.decode('utf-8')


def parse(xml, year):
    """t=1 — нерабочий, t=2 — сокращённый (для нас рабочий), t=3 — рабочая суббота/воскресенье."""
    holidays, workdays = [], []
    for m in re.finditer(r'<day\s+d="(\d\d)\.(\d\d)"\s+t="(\d)"([^/>]*)/>', xml):
        mm, dd, t, rest = m.groups()
        date = '%s-%s-%s' % (year, mm, dd)
        note = ''
        f = re.search(r'f="(\d\d)\.(\d\d)"', rest)
        if f:
            note = '   # перенос с %s.%s' % (f.group(2), f.group(1))
        if t == '1':
            holidays.append(date + note)
        elif t == '3':
            workdays.append('рабочий ' + date + note)
    return holidays, workdays


def write(year, holidays, workdays):
    lines = open(CAL, encoding='utf-8').read().split('\n') if os.path.exists(CAL) else []
    out, skip = [], False
    for line in lines:
        if re.match(r'^#\s*---\s*%s\s*---' % year, line):
            skip = True
            continue
        if skip:
            if re.match(r'^#\s*---\s*\d{4}\s*---', line):
                skip = False
            elif re.match(r'^(рабочий\s+)?%s-' % year, line.strip()) or not line.strip():
                continue
            else:
                skip = False
        out.append(line)
    while out and not out[-1].strip():
        out.pop()
    out += ['', '# --- %s --- (обновлено из xmlcalendar.ru)' % year] + holidays + workdays + ['']
    open(CAL, 'w', encoding='utf-8').write('\n'.join(out))


def main():
    years = sys.argv[1:]
    if not years:
        raise SystemExit(__doc__)
    for year in years:
        if not re.match(r'^\d{4}$', year):
            raise SystemExit('Год указывается четырьмя цифрами, например 2026')
        h, w = parse(download(year), year)
        if not h:
            print('%s год: данных пока нет. Календарь на следующий год публикуется после того,\n'
                  '  как правительство утвердит переносы выходных. Прежний блок оставлен без изменений.' % year)
            continue
        write(year, h, w)
        print('%s год: нерабочих дней %d, рабочих суббот/воскресений %d' % (year, len(h), len(w)))
    print('Календарь обновлён: %s' % CAL)


if __name__ == '__main__':
    main()
