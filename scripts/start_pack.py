#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Готовит сразу пачку писем — например, все письма первых дней контракта.

    python3 scripts/start_pack.py --список          # что войдёт в пакет, без создания файлов
    python3 scripts/start_pack.py                   # подготовить письма старта контракта
    python3 scripts/start_pack.py --этап 1          # письма по этапу 1
    python3 scripts/start_pack.py --все             # все пункты плана, у которых есть шаблон

Общие значения можно задать сразу для всех писем:
    python3 scripts/start_pack.py --значение "дата_контракта=18.09.2026"

Письма сохраняются в 5_Готовые_письма как …_v1.docx и заводятся в реестр со статусом «черновик».
Поля, которые скрипт не знает (исходящие номера, количество листов), остаются видимыми
как {{...}} — их заполняет человек или Клод при подготовке конкретного письма.
"""
import sys, os, json, datetime as dt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_letter import auto_values, fill_docx, find_fields, next_version, esc
import registry

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONTRACT = os.path.join(BASE, '4_План_и_журнал', 'контракт.json')
OUT = os.path.join(BASE, '5_Готовые_письма')


def тема_из_шаблона(путь):
    имя = os.path.splitext(os.path.basename(путь))[0]
    if имя[:2].isdigit() and имя[2:3] == '_':
        имя = имя[3:]
    return имя.replace('_', ' ')


def отобрать(d, режим, этап):
    """Группирует обязательства по шаблону: одно письмо может закрывать несколько пунктов."""
    письма = {}
    for o in d.get('обязательства', []):
        if o.get('справочно') or not o.get('шаблон'):
            continue
        ид = o['id']
        if режим == 'старт' and not ид.startswith('П-'):
            continue
        if режим == 'этап':
            этапы = o.get('этапы') or []
            if этап not in этапы and not ид.startswith('П-'):
                continue
        письма.setdefault(o['шаблон'], {'пункты': [], 'кому': o.get('кому', ''),
                                        'основание': o.get('основание', '')})
        письма[o['шаблон']]['пункты'].append(ид)
    return письма


def main():
    a = sys.argv[1:]
    if not os.path.exists(CONTRACT):
        raise SystemExit('Нет паспорта контракта. Сначала: python3 scripts/check.py')
    d = json.load(open(CONTRACT, encoding='utf-8'))

    режим, этап = 'старт', None
    if '--все' in a or '--всё' in a:
        режим = 'все'
    if '--этап' in a:
        режим, этап = 'этап', int(a[a.index('--этап') + 1])

    values = auto_values()
    for i, x in enumerate(a):
        if x == '--значение':
            k, _, v = a[i + 1].partition('=')
            values[k.strip()] = v
    if '--этап' in a:
        values.setdefault('этап', str(этап))

    письма = отобрать(d, режим, этап)
    if not письма:
        raise SystemExit('Не нашлось пунктов с шаблонами. Посмотрите: python3 scripts/check.py')

    print('В пакет войдёт писем: %d\n' % len(письма))
    for шаблон, инфо in sorted(письма.items()):
        print('  %-9s %s' % (', '.join(инфо['пункты']), тема_из_шаблона(шаблон)))
    if '--список' in a:
        print('\nЭто предварительный список. Чтобы подготовить письма, запустите без --список.')
        return

    print()
    сегодня = dt.date.today().isoformat()
    for шаблон, инфо in sorted(письма.items()):
        тема = тема_из_шаблона(шаблон)
        имя = '%s_%s_%s.docx' % (сегодня, '-'.join(инфо['пункты']),
                                 тема.lower().replace(' ', '-')[:40].rstrip('-'))
        путь, версия = next_version(os.path.join(OUT, имя))
        fill_docx(os.path.join(BASE, шаблон), путь, values)
        rel = os.path.relpath(путь, BASE)
        ид = registry.register(rel, тема, инфо['кому'], ', '.join(инфо['пункты']), версия)
        осталось = find_fields(путь)
        print('%s  %s' % (ид, rel))
        if осталось:
            print('    заполнить: %s' % ', '.join('{{%s}}' % f for f in осталось))

    print('\nПисьма готовы как черновики. Дальше: заполнить оставшиеся поля, вычитать, '
          'подписать и отправить.\nРеестр: python3 scripts/registry.py')


if __name__ == '__main__':
    main()
