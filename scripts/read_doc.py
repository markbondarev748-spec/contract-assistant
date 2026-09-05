#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Читает .docx / .odt / .txt / .md и печатает текст.

Нужен, чтобы Клод мог прочитать контракт, ТЗ и приложения:
    python3 scripts/read_doc.py "1_Документы_контракта/ТЗ.odt"
    python3 scripts/read_doc.py "1_Документы_контракта/ТЗ.odt" --строки 900-1000

Работает на голом Python 3.8+, ничего устанавливать не надо.
"""
import sys, os, re, html, zipfile

def _clean(t):
    t = re.sub(r'[ \t]+\n', '\n', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()

def docx_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8', 'ignore')
    xml = re.sub(r'</w:p>', '\n', xml)
    xml = re.sub(r'</w:tr>', '\n', xml)
    xml = re.sub(r'</w:tc>', ' | ', xml)
    xml = re.sub(r'<w:tab[^>]*/>', '\t', xml)
    xml = re.sub(r'<w:br[^>]*/>', '\n', xml)
    return _clean(html.unescape(re.sub(r'<[^>]+>', '', xml)))

def odt_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('content.xml').decode('utf-8', 'ignore')
    xml = re.sub(r'</text:(p|h)>', '\n', xml)
    xml = re.sub(r'</table:table-row>', '\n', xml)
    xml = re.sub(r'</table:table-cell>', ' | ', xml)
    xml = re.sub(r'<text:tab[^>]*/>', '\t', xml)
    xml = re.sub(r'<text:line-break[^>]*/>', '\n', xml)
    return _clean(html.unescape(re.sub(r'<[^>]+>', '', xml)))

def read_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.docx', '.dotx', '.docm'):
        return docx_text(path)
    if ext in ('.odt', '.ott'):
        return odt_text(path)
    if ext in ('.txt', '.md', '.csv'):
        with open(path, encoding='utf-8', errors='ignore') as f:
            return f.read()
    if ext == '.doc':
        raise SystemExit('Формат .doc (старый Word) не читается. Пересохраните файл как .docx.')
    raise SystemExit('Неизвестный формат: %s. Поддерживаются .docx, .odt, .txt, .md' % ext)

def main():
    args = [a for a in sys.argv[1:]]
    if not args:
        raise SystemExit(__doc__)
    path = args[0]
    if not os.path.exists(path):
        raise SystemExit('Файл не найден: %s' % path)
    text = read_any(path)
    lines = text.split('\n')
    if '--строки' in args:
        rng = args[args.index('--строки') + 1]
        a, _, b = rng.partition('-')
        lines = lines[int(a) - 1: int(b or a)]
    if '--сколько-строк' in args:
        print(len(text.split('\n')))
        return
    sys.stdout.write('\n'.join(lines))

if __name__ == '__main__':
    main()
