#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Готовит официальное письмо в Word.

1) Посмотреть, какие поля есть в шаблоне:
       python3 scripts/make_letter.py --поля "2_Шаблоны_писем/письмо.docx"

2) Заполнить шаблон и сохранить готовое письмо:
       python3 scripts/make_letter.py --шаблон "2_Шаблоны_писем/письмо.docx" \
           --выход "5_Готовые_письма/2026-09-18_исх-123.docx" \
           --значение исх_номер=123 --значение текст="Уважаемый Евгений Юрьевич! ..."

3) Вести версии и реестр (письма вычитывают в несколько кругов):
       --версия            сохранит как …_v1.docx, следующий вызов — _v2, _v3
       --тема "О ..." --кому "Минздрав ПК" --пункт П-02
                           запишет письмо в 4_План_и_журнал/реестр_писем.csv

4) Если готового шаблона нет — собрать письмо из текста:
       python3 scripts/make_letter.py --из-текста черновик.txt \
           --выход "5_Готовые_письма/письмо.docx"

Поля в шаблоне размечаются двойными фигурными скобками: {{исх_номер}}, {{текст}}.
Разово поправить любой текст прямо в шаблоне (например число листов приложения):
    --замена "на 4 л.=на 6 л." 
Значения из 4_План_и_журнал/контракт.json подставляются сами:
  {{номер_контракта}} {{дата_контракта}} {{предмет_контракта}} {{цена_контракта}}
  {{заказчик}} {{подписант_заказчика}} {{исполнитель}} {{подписант_исполнителя}}
  {{должность_подписанта}} {{телефон}} {{исполнитель_письма}} {{дата}} {{сегодня}}
"""
import sys, os, re, json, shutil, zipfile, datetime as dt

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONTRACT = os.path.join(BASE, '4_План_и_журнал', 'контракт.json')
МЕСЯЦЫ = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
          'августа', 'сентября', 'октября', 'ноября', 'декабря']


def auto_values():
    v = {}
    if os.path.exists(CONTRACT):
        d = json.load(open(CONTRACT, encoding='utf-8'))
        k, s = d.get('контракт', {}), d.get('стороны', {})
        z, i = s.get('заказчик', {}), s.get('исполнитель', {})
        v.update({
            'номер_контракта': k.get('номер', ''),
            'дата_контракта': k.get('дата_заключения', ''),
            'предмет_контракта': k.get('предмет', ''),
            'цена_контракта': k.get('цена', ''),
            'заказчик': z.get('наименование', ''),
            'подписант_заказчика': z.get('подписант', ''),
            'исполнитель': i.get('наименование', ''),
            'подписант_исполнителя': i.get('подписант', ''),
            'должность_подписанта': i.get('должность_подписанта', ''),
            'телефон': i.get('телефон', ''),
            'исполнитель_письма': i.get('исполнитель_письма', ''),
        })
    v = dict((k, x) for k, x in v.items() if str(x).strip())  # пустые не подставляем — останутся видны как {{...}}
    t = dt.date.today()
    v['дата'] = '«%02d» %s %d г.' % (t.day, МЕСЯЦЫ[t.month - 1], t.year)
    v['сегодня'] = t.isoformat()
    return v


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def replace_xml(xml, pairs):
    """Заменяет пары (что, на что) в тексте документа.

    Word рвёт текст абзаца на куски, и фраза может быть разделена между ними. Поэтому:
    целые куски правим по одному (форматирование внутри абзаца сохраняется), а если
    хоть одна фраза разорвана — склеиваем абзац целиком и применяем замены к нему,
    иначе короткая фраза успевает сработать раньше длинной и портит замену."""
    def sub_text(t):
        for a, b in pairs:
            t = t.replace(a, b)
        return t

    out, pos = [], 0
    for m in re.finditer(r'<w:p[ >].*?</w:p>', xml, re.S):
        para = m.group(0)
        chunks = re.findall(r'<w:t[^>]*>(.*?)</w:t>', para, re.S)
        joined = ''.join(chunks)
        разорванные = [a for a, _ in pairs
                       if a in joined and not any(a in c for c in chunks)]
        if разорванные:
            filled = sub_text(joined)
            first = [True]

            def repl(x):
                if first[0]:
                    first[0] = False
                    return x.group(1) + filled + x.group(3)
                return x.group(1) + x.group(3)
            para = re.sub(r'(<w:t[^>]*>)(.*?)(</w:t>)', repl, para, flags=re.S)
        elif any(a in joined for a, _ in pairs):
            para = re.sub(r'(<w:t[^>]*>)(.*?)(</w:t>)',
                          lambda x: x.group(1) + sub_text(x.group(2)) + x.group(3),
                          para, flags=re.S)
        out.append(xml[pos:m.start()])
        out.append(para)
        pos = m.end()
    out.append(xml[pos:])
    return ''.join(out)


def fill_xml(xml, values):
    return replace_xml(xml, [('{{%s}}' % k, esc(str(v))) for k, v in values.items()])


PARTS = ('word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml',
         'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml')


def rewrite_docx(src, dst, transform):
    """Копирует .docx, пропуская текстовые части через transform(xml) -> xml."""
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in PARTS:
                data = transform(data.decode('utf-8')).encode('utf-8')
            zout.writestr(item, data)
    zin.close()


def fill_docx(template, output, values, pairs=()):
    check_docx(template)
    rewrite_docx(template, output,
                 lambda xml: replace_xml(fill_xml(xml, values), list(pairs)))


def check_docx(path):
    if not os.path.exists(path):
        raise SystemExit('Файл не найден: %s' % path)
    if not path.lower().endswith(('.docx', '.dotx')):
        raise SystemExit('Шаблон должен быть в формате .docx (сейчас: %s).\n'
                         'Старый .doc нужно пересохранить в Word как .docx.' % os.path.splitext(path)[1])
    if not zipfile.is_zipfile(path):
        raise SystemExit('Файл «%s» не читается как документ Word. Откройте его в Word и пересохраните как .docx.' % os.path.basename(path))


def find_fields(template):
    check_docx(template)
    with zipfile.ZipFile(template) as z:
        names = [n for n in z.namelist() if n.startswith('word/') and n.endswith('.xml')]
        found = []
        for n in names:
            xml = z.read(n).decode('utf-8', 'ignore')
            text = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', xml, re.S))
            found += re.findall(r'\{\{([^}]+)\}\}', text)
    seen, res = set(), []
    for f in found:
        if f not in seen:
            seen.add(f)
            res.append(f)
    return res


def write_docx(body_xml, output, landscape=False):
    """Собирает минимальный файл .docx из готового содержимого <w:body>."""
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            '<w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="1134"/></w:sectPr>'
            if landscape else
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s%s</w:body></w:document>' % (body_xml, sect))
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        z.writestr('word/document.xml', doc.encode('utf-8'))


def para(text, size=28, bold=False, align='both', indent=709):
    """Абзац Times New Roman. size — половины пункта (28 = 14 пт)."""
    rpr = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="%d"/>%s' % (
        size, '<w:b/>' if bold else '')
    ppr = '<w:pPr><w:spacing w:after="60"/><w:jc w:val="%s"/>%s</w:pPr>' % (
        align, '<w:ind w:firstLine="%d"/>' % indent if indent else '')
    return ('<w:p>%s<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, rpr, esc(text)))


def docx_from_text(text, output):
    body = ''.join(para(line.strip()) for line in text.split('\n'))
    write_docx(body, output)


def next_version(path):
    """Возвращает путь со следующим номером версии: письмо.docx -> письмо_v1.docx, _v2, _v3…"""
    d = os.path.dirname(os.path.abspath(path))
    base, ext = os.path.splitext(os.path.basename(path))
    base = re.sub(r'_v\d+$', '', base)
    n = 0
    if os.path.isdir(d):
        for f in os.listdir(d):
            m = re.match(re.escape(base) + r'_v(\d+)' + re.escape(ext) + '$', f)
            if m:
                n = max(n, int(m.group(1)))
    return os.path.join(os.path.dirname(path), '%s_v%d%s' % (base, n + 1, ext)), n + 1


def main():
    a = sys.argv[1:]
    if not a:
        raise SystemExit(__doc__)
    if '--поля' in a:
        tpl = a[a.index('--поля') + 1]
        fields = find_fields(tpl)
        if fields:
            print('Поля в шаблоне «%s»:' % os.path.basename(tpl))
            for f in fields:
                print('  {{%s}}' % f)
        else:
            print('В шаблоне нет полей вида {{...}}. Варианты:\n'
                  '  — открыть шаблон и заменить изменяемые места на {{имя_поля}};\n'
                  '  — либо собрать письмо из текста: --из-текста черновик.txt --выход ...')
        return
    if '--выход' not in a:
        raise SystemExit('Не указан --выход (куда сохранить письмо)')
    out = a[a.index('--выход') + 1]
    версия = 1
    if '--версия' in a:
        out, версия = next_version(out)
    values = auto_values()
    for i, x in enumerate(a):
        if x == '--значение':
            k, _, v = a[i + 1].partition('=')
            values[k.strip()] = v
    pairs = []
    for i, x in enumerate(a):
        if x == '--замена':
            k, _, v = a[i + 1].partition('=')
            pairs.append((k, esc(v)))
    if '--из-текста' in a:
        src = a[a.index('--из-текста') + 1]
        docx_from_text(open(src, encoding='utf-8').read(), out)
    elif '--шаблон' in a:
        fill_docx(a[a.index('--шаблон') + 1], out, values, pairs)
    else:
        raise SystemExit('Нужен --шаблон или --из-текста')
    print('Письмо сохранено: %s' % out)
    if '--тема' in a:
        import registry
        rel = os.path.relpath(os.path.abspath(out), BASE)
        ид = registry.register(rel, a[a.index('--тема') + 1],
                               a[a.index('--кому') + 1] if '--кому' in a else '',
                               a[a.index('--пункт') + 1] if '--пункт' in a else '',
                               версия, values.get('исх_номер', ''))
        print('В реестре: %s, версия %d' % (ид, версия))
    left = find_fields(out)
    if left:
        print('ВНИМАНИЕ, остались незаполненные поля: %s' % ', '.join('{{%s}}' % f for f in left))


if __name__ == '__main__':
    main()
