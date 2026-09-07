#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Размечает шаблон письма: заменяет заглушки на поля {{...}}.

    python3 scripts/mark_template.py "2_Шаблоны_писем/новое_письмо.docx"
        — только показать, что нашлось (ничего не меняет)

    python3 scripts/mark_template.py "2_Шаблоны_писем/новое_письмо.docx" --применить
        — заменить однозначные заглушки: исходящий номер, дата письма, номер и дата контракта

    python3 scripts/mark_template.py "…docx" --применить --замена "ХХ.ХХ.2026={{дата_заявки}}"
        — плюс точечные замены для остального (дат, номеров приложений и т.п.)

Форматирование, бланк и подписи не затрагиваются — меняется только текст заглушек.
"""
import sys, os, re, json, shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_letter import rewrite_docx, replace_xml, find_fields, check_docx
from read_doc import docx_text

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONTRACT = os.path.join(BASE, '4_План_и_журнал', 'контракт.json')

# Заглушки, которые встречаются в письмах и однозначно понятны
ПРОБЕЛ = '[ \u00a0\u2009\u202f]*'   # обычный, неразрывный и узкие пробелы — Word ставит разные
ОДНОЗНАЧНЫЕ = [
    (r'Исх\.' + ПРОБЕЛ + r'№' + ПРОБЕЛ + r'(NNNN|N{3,}|_{3,})', 'Исх. № {{исх_номер}}'),
    (r'от' + ПРОБЕЛ + r'«(ХХХ|_{1,})»' + ПРОБЕЛ + r'(ХХХХ|_{3,})' + ПРОБЕЛ + r'20\d\d' + ПРОБЕЛ + r'г\.', 'от {{дата}}'),
]
# Заглушки дат и номеров, для которых имя поля выбирает человек
НЕОДНОЗНАЧНЫЕ = [r'Х{2}\.Х{2}\.20\d\d', r'Х{2}\.0?Х\.20\d\d', r'_{2}\.[_0-9]{2}\.20\d\d',
                 r'_{3,}', r'Х{3,}', r'№ *_{3,}']


def контекст(текст, m, ширина=45):
    a = max(0, m.start() - ширина)
    b = min(len(текст), m.end() + ширина)
    return ('…' if a else '') + текст[a:b].replace('\n', ' ') + ('…' if b < len(текст) else '')


def автозамены(текст):
    """Возвращает список (что найдено, на что заменить)."""
    pairs = []
    for шаблон, замена in ОДНОЗНАЧНЫЕ:
        for m in re.finditer(шаблон, текст):
            pairs.append((m.group(0), замена))
    # номер контракта из паспорта
    if os.path.exists(CONTRACT):
        номер = json.load(open(CONTRACT, encoding='utf-8')).get('контракт', {}).get('номер')
        if номер and номер in текст:
            for m in re.finditer(r'от' + ПРОБЕЛ + r'(_{3,}|Х{2}\.0?Х\.20\d\d|Х{2}\.\d\d\.20\d\d)' + ПРОБЕЛ + r'№' + ПРОБЕЛ + re.escape(номер), текст):
                pairs.append((m.group(0), 'от {{дата_контракта}} № {{номер_контракта}}'))
            pairs.append((номер, '{{номер_контракта}}'))
    # длинные фразы заменяем раньше коротких, иначе короткая съест часть длинной
    return sorted(set(pairs), key=lambda p: -len(p[0]))


def main():
    a = sys.argv[1:]
    if not a:
        raise SystemExit(__doc__)
    путь = a[0]
    check_docx(путь)
    текст = docx_text(путь)

    pairs = автозамены(текст)
    print('Файл: %s\n' % os.path.basename(путь))
    if pairs:
        print('Однозначные заглушки (заменяются автоматически):')
        for что, на in pairs:
            print('  «%s»  ->  %s' % (что, на))
    else:
        print('Однозначных заглушек не нашлось.')

    ручные = []
    for шаблон in НЕОДНОЗНАЧНЫЕ:
        for m in re.finditer(шаблон, текст):
            если_уже = any(m.group(0) in что for что, _ in pairs)
            if not если_уже:
                ручные.append((m.group(0), контекст(текст, m)))
    if ручные:
        print('\nОстальные заглушки — имя поля выберите сами:')
        видели = set()
        for что, ctx in ручные:
            if (что, ctx) in видели:
                continue
            видели.add((что, ctx))
            print('  «%s» в тексте: %s' % (что, ctx))
        print('\n  Пример: --замена "ХХ.ХХ.2026={{дата_начала_услуг}}"')

    свои = []
    for i, x in enumerate(a):
        if x == '--замена':
            k, _, v = a[i + 1].partition('=')
            свои.append((k, v))

    if '--применить' not in a:
        print('\nЭто предварительный просмотр. Чтобы применить, добавьте --применить')
        return

    выход = a[a.index('--выход') + 1] if '--выход' in a else путь
    if выход == путь:
        shutil.copy2(путь, путь + '.bak')
    rewrite_docx(путь, выход + '.tmp', lambda xml: replace_xml(xml, свои + pairs))
    shutil.move(выход + '.tmp', выход)
    поля = find_fields(выход)
    print('\nГотово: %s' % выход)
    print('Поля в шаблоне: %s' % (', '.join('{{%s}}' % f for f in поля) or 'нет'))
    if выход == путь:
        print('Исходный файл сохранён рядом как %s.bak' % os.path.basename(путь))
    print('Проверьте письмо: python3 scripts/read_doc.py "%s"' % выход)


if __name__ == '__main__':
    main()
