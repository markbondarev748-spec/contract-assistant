#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Реестр исходящих писем: что подготовлено, в какой версии, что отправлено.

    python3 scripts/registry.py                        — показать реестр
    python3 scripts/registry.py L-003 на_вычитке       — сменить статус
    python3 scripts/registry.py L-003 отправлено --исх 1057 --дата 2026-09-18
    python3 scripts/registry.py L-003 --примечание "правки от юриста учтены"

Статусы: черновик · на_вычитке · готово · отправлено
Записи создаются сами, когда письмо готовится через make_letter.py с указанием --тема.
"""
import sys, os, csv, datetime as dt

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REG = os.path.join(BASE, '4_План_и_журнал', 'реестр_писем.csv')
COLS = ['ID', 'Дата подготовки', 'Версия', 'Тема', 'Кому', 'Пункт плана',
        'Исх. номер', 'Статус', 'Файл', 'Примечание']
СТАТУСЫ = ('черновик', 'на_вычитке', 'готово', 'отправлено')


def read():
    if not os.path.exists(REG):
        return []
    with open(REG, encoding='utf-8-sig') as f:
        return [r for r in csv.DictReader(f, delimiter=';') if r.get('ID')]


def write(rows):
    os.makedirs(os.path.dirname(REG), exist_ok=True)
    with open(REG, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter=';')
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, '') for c in COLS})


def register(файл, тема, кому='', пункт='', версия=1, исх=''):
    """Заводит письмо в реестре или обновляет его версию. Возвращает ID."""
    rows = read()
    ключ = os.path.basename(файл).split('_v')[0]
    for r in rows:
        if os.path.basename(r.get('Файл', '')).split('_v')[0] == ключ:
            r.update({'Версия': str(версия), 'Файл': файл,
                      'Дата подготовки': dt.date.today().isoformat(),
                      'Статус': 'черновик' if версия == 1 else 'на_вычитке'})
            if исх:
                r['Исх. номер'] = исх
            write(rows)
            return r['ID']
    ид = 'L-%03d' % (len(rows) + 1)
    rows.append({'ID': ид, 'Дата подготовки': dt.date.today().isoformat(),
                 'Версия': str(версия), 'Тема': тема, 'Кому': кому, 'Пункт плана': пункт,
                 'Исх. номер': исх, 'Статус': 'черновик', 'Файл': файл, 'Примечание': ''})
    write(rows)
    return ид


def show(rows):
    if not rows:
        print('Реестр пуст. Он заполняется сам, когда письма готовятся через make_letter.py.')
        return
    print('%-6s %-11s %-4s %-8s %-38s %s' % ('ID', 'Дата', 'вер.', 'Исх.', 'Тема', 'Статус'))
    for r in rows:
        print('%-6s %-11s v%-3s %-8s %-38s %s' % (
            r['ID'], r['Дата подготовки'], r['Версия'], r['Исх. номер'] or '—',
            r['Тема'][:38], r['Статус']))
    n = sum(1 for r in rows if r['Статус'] != 'отправлено')
    if n:
        print('\nВ работе: %d. Отправленные письма отмечайте статусом «отправлено».' % n)


def main():
    a = sys.argv[1:]
    rows = read()
    if not a:
        show(rows)
        return
    ид = a[0].upper()
    строка = next((r for r in rows if r['ID'].upper() == ид), None)
    if not строка:
        raise SystemExit('В реестре нет письма %s. Список: python3 scripts/registry.py' % ид)
    for x in a[1:]:
        if x in СТАТУСЫ:
            строка['Статус'] = x
    if '--исх' in a:
        строка['Исх. номер'] = a[a.index('--исх') + 1]
    if '--дата' in a:
        строка['Дата подготовки'] = a[a.index('--дата') + 1]
    if '--примечание' in a:
        строка['Примечание'] = a[a.index('--примечание') + 1]
    write(rows)
    print('%s: статус «%s», исх. № %s' % (строка['ID'], строка['Статус'], строка['Исх. номер'] or '—'))
    if строка['Статус'] == 'отправлено':
        print('Не забудьте отметить факт отправки в журнале событий, если пункт плана закрыт.')


if __name__ == '__main__':
    main()
