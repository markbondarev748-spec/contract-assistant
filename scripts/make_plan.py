#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собирает план работ по контракту: обязательства + журнал событий = что и когда делать.

    python3 scripts/make_plan.py                  # пересчитать план
    python3 scripts/make_plan.py --в-word         # ещё и сохранить план в Word
    python3 scripts/make_plan.py --в-confluence   # таблица для вставки в Confluence
    python3 scripts/make_plan.py --в-календарь    # файл .ics для Outlook (напоминание за день)
    python3 scripts/make_plan.py --всё            # все выгрузки сразу
    python3 scripts/make_plan.py --на-дату 2026-10-01

Читает:  4_План_и_журнал/контракт.json  и  4_План_и_журнал/журнал_событий.csv
Пишет:   4_План_и_журнал/план.csv  (открывается в Excel)  и сводку в терминал.
"""
import sys, os, csv, json, hashlib, datetime as dt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import workdays as wd
from make_letter import write_docx, para, esc

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR = os.path.join(BASE, '4_План_и_журнал')
CONTRACT = os.path.join(DIR, 'контракт.json')
JOURNAL = os.path.join(DIR, 'журнал_событий.csv')
PLAN = os.path.join(DIR, 'план.csv')
PLAN_DOCX = os.path.join(DIR, 'план.docx')
PLAN_WIKI = os.path.join(DIR, 'план_для_confluence.txt')
PLAN_ICS = os.path.join(DIR, 'сроки.ics')

COLS = [('Срок', 1300), ('Статус', 2000), ('Что сделать', 6800), ('Основание', 1900), ('Канал', 2000)]


def _cell(text, w, bold=False, shade=False):
    tcpr = ('<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>'
            % (w, '<w:shd w:val="clear" w:fill="E8E8E8"/>' if shade else ''))
    return '<w:tc>%s%s</w:tc>' % (tcpr, para(text, size=20, bold=bold, align='left', indent=0))


def to_docx(rows, contract, today):
    """Складывает тот же план в файл Word — таблицей, альбомная страница."""
    k = contract.get('контракт', {})
    head = [para('План работ по контракту', size=28, bold=True, align='center', indent=0)]
    подзаг = 'Контракт № %s от %s' % (k.get('номер') or '______', k.get('дата_заключения') or '__________')
    head.append(para(подзаг, size=24, align='center', indent=0))
    head.append(para('Предмет: %s' % k.get('предмет', ''), size=20, align='left', indent=0))
    head.append(para('Составлен на %s' % today.isoformat(), size=20, align='left', indent=0))

    borders = ('<w:tblBorders>' + ''.join(
        '<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % b
        for b in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>')
    trs = ['<w:tr><w:trPr><w:tblHeader/></w:trPr>%s</w:tr>'
           % ''.join(_cell(n, w, bold=True, shade=True) for n, w in COLS)]
    for r in rows:
        cells = ''.join(_cell(r[n] if n != 'Срок' else (r[n] or '—'), w) for n, w in COLS)
        trs.append('<w:tr>%s</w:tr>' % cells)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>%s</w:tblPr>%s</w:tbl>'
           % (borders, ''.join(trs)))

    tail = [para('', size=20, indent=0),
            para('Пункты со статусом «ждём событие» получат дату, как только соответствующий '
                 'факт будет отмечен в журнале событий.', size=18, align='left', indent=0)]
    write_docx(''.join(head) + tbl + ''.join(tail), PLAN_DOCX, landscape=True)


def load_journal():
    """Возвращает {код события: дата}. Строки с id обязательства = отметка о выполнении."""
    events = {}
    if not os.path.exists(JOURNAL):
        return events
    with open(JOURNAL, encoding='utf-8-sig') as f:
        for row in csv.reader(f, delimiter=';'):
            if not row or not row[0].strip() or row[0].strip().startswith('#'):
                continue
            code = row[0].strip()
            if code in ('код', 'событие'):
                continue
            date = (row[1] if len(row) > 1 else '').strip()
            if not date:
                continue
            try:
                events[code] = wd.parse(date)
            except Exception:
                print('  ! не понял дату в журнале: %s;%s' % (code, date))
    return events


def expand(obligations):
    """Размножает обязательства, отмеченные полем "этапы", по каждому этапу."""
    out = []
    for o in obligations:
        stages = o.get('этапы')
        if not stages:
            out.append(dict(o, _этап=None))
            continue
        for st in stages:
            c = json.loads(json.dumps(o))
            c['id'] = '%s.%s' % (o['id'], st)
            c['_этап'] = st
            s = c.get('срок', {})
            if 'событие' in s:
                s['событие'] = s['событие'].replace('{этап}', str(st))
            if isinstance(s.get('дата_этапа'), dict) and s['дата_этапа'].get('этап') == 'текущий':
                s['дата_этапа']['этап'] = st
            c['что'] = c['что'].replace('{этап}', str(st))
            out.append(c)
    return out


def stage_date(contract, ref):
    for st in contract['этапы']:
        if st['номер'] == ref.get('этап'):
            v = st.get(ref.get('поле', 'оказать_до'))
            return wd.parse(v) if v else None
    return None


def compute(o, contract, events):
    """-> (дата или None, пояснение)"""
    s = o.get('срок', {})
    kind = s.get('тип')
    if kind == 'от_события':
        ev = s.get('событие')
        if ev in events:
            return wd.add(events[ev], int(s.get('смещение', 0))), 'через %s р.д. после «%s» (%s)' % (
                s.get('смещение'), ev, events[ev].isoformat())
        return None, 'ждём событие «%s» — отметьте его в журнале' % ev
    if kind == 'до_даты':
        base = stage_date(contract, s.get('дата_этапа', {}))
        if base:
            return wd.add(base, -int(s.get('смещение', 0))), 'за %s р.д. до %s' % (s.get('смещение'), base.isoformat())
        return None, 'не задана дата этапа'
    if kind == 'фиксированная':
        return wd.parse(s['дата']), 'фиксированный срок'
    return None, 'по мере необходимости'


def status(date, done, today, why=''):
    if done:
        return 'ВЫПОЛНЕНО'
    if date is None:
        return 'по необходимости' if why.startswith('по мере') else 'ждём событие'
    if date < today:
        return 'ПРОСРОЧЕНО'
    left = wd.between(today, date)
    if left <= 3:
        return 'СРОЧНО (%s р.д.)' % left
    return 'запланировано (%s р.д.)' % left


def to_confluence(rows, contract, today):
    """Таблица для вставки в Confluence: вики-разметка и markdown."""
    k = contract.get('контракт', {})
    cols = ['Срок', 'Статус', 'Что сделать', 'Основание', 'Канал']
    заголовок = 'План работ по контракту № %s от %s (составлен %s)' % (
        k.get('номер') or '______', k.get('дата_заключения') or '__________', today.isoformat())

    def строка(r, sep):
        return sep + sep.join((r[c] or '—').replace('|', '/') for c in cols) + sep

    wiki = ['h2. ' + заголовок, '', '||' + '||'.join(cols) + '||']
    wiki += [строка(r, '|') for r in rows]

    md = ['## ' + заголовок, '', '| ' + ' | '.join(cols) + ' |', '|' + '---|' * len(cols)]
    md += ['| ' + ' | '.join((r[c] or '—').replace('|', '/') for c in cols) + ' |' for r in rows]

    text = ('=== ВАРИАНТ 1: вики-разметка (Confluence Server/Data Center) ===\n'
            'В редакторе страницы: Вставить > Разметка > Confluence Wiki\n\n'
            + '\n'.join(wiki) +
            '\n\n\n=== ВАРИАНТ 2: markdown (Confluence Cloud) ===\n'
            'В редакторе страницы: Вставить > Разметка > Markdown\n\n'
            + '\n'.join(md) + '\n')
    open(PLAN_WIKI, 'w', encoding='utf-8').write(text)


def _ics_line(name, value):
    """Складывает строку по правилам iCalendar (не длиннее 75 символов)."""
    v = (str(value).replace('\\', '\\\\').replace(';', '\\;')
         .replace(',', '\\,').replace('\n', '\\n'))
    line = '%s:%s' % (name, v)
    out, cur = [], ''
    for ch in line:
        if len(cur.encode('utf-8')) + len(ch.encode('utf-8')) > 73:
            out.append(cur)
            cur = ' '
        cur += ch
    out.append(cur)
    return '\r\n'.join(out)


def to_ics(rows, contract, today):
    """Ключевые даты в файл .ics — импортируется в Outlook или любой календарь."""
    stamp = dt.datetime.now().strftime('%Y%m%dT%H%M%SZ')
    out = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//contract-assistant//RU',
           'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
           _ics_line('X-WR-CALNAME', 'Сроки по контракту %s'
                     % (contract.get('контракт', {}).get('номер') or ''))]

    def событие(date, summary, description, uid):
        d = wd.parse(date) if isinstance(date, str) else date
        # UID должен состоять из латиницы, поэтому берём устойчивый отпечаток ключа:
        # при повторном импорте календарь обновит событие, а не создаст дубль
        uid = 'ca-' + hashlib.md5(uid.encode('utf-8')).hexdigest()[:16]
        out.extend(['BEGIN:VEVENT', _ics_line('UID', uid + '@contract-assistant'),
                    'DTSTAMP:' + stamp,
                    'DTSTART;VALUE=DATE:' + d.strftime('%Y%m%d'),
                    'DTEND;VALUE=DATE:' + (d + dt.timedelta(days=1)).strftime('%Y%m%d'),
                    _ics_line('SUMMARY', summary), _ics_line('DESCRIPTION', description),
                    'BEGIN:VALARM', 'TRIGGER:-P1D', 'ACTION:DISPLAY',
                    _ics_line('DESCRIPTION', summary), 'END:VALARM', 'END:VEVENT'])

    n = 0
    for st in contract.get('этапы', []):
        for поле, подпись in (('оказать_до', 'оказать услуги'), ('сдать_и_оплатить_до', 'сдача и оплата')):
            if st.get(поле):
                n += 1
                событие(st[поле], 'Этап %s: %s — крайний срок' % (st['номер'], подпись),
                        st.get('что', ''), 'stage%s-%s' % (st['номер'], поле))
    for r in rows:
        if r['Срок'] and not r['Статус'].startswith(('справочно', 'ВЫПОЛНЕНО')):
            n += 1
            событие(r['Срок'], '%s — %s' % (r['ID'], r['Что сделать'][:60]),
                    '%s\nОснование: %s\nКанал: %s' % (r['Что сделать'], r['Основание'], r['Канал']),
                    'item-' + r['ID'])
    out.append('END:VCALENDAR')
    open(PLAN_ICS, 'w', encoding='utf-8', newline='').write('\r\n'.join(out) + '\r\n')
    return n


def main():
    today = dt.date.today()
    if '--на-дату' in sys.argv:
        today = wd.parse(sys.argv[sys.argv.index('--на-дату') + 1])
    if not os.path.exists(CONTRACT):
        raise SystemExit('Нет файла %s' % CONTRACT)
    try:
        contract = json.load(open(CONTRACT, encoding='utf-8'))
    except ValueError as e:
        raise SystemExit('Файл контракт.json повреждён (ошибка разметки: %s).\n'
                         'Скорее всего, при ручной правке пропущена запятая или кавычка.' % e)
    events = load_journal()

    rows = []
    for o in expand(contract.get('обязательства', [])):
        date, why = compute(o, contract, events)
        done = o['id'] in events
        rows.append({
            '_справочно': bool(o.get('справочно')),
            'Срок': date.isoformat() if date else '',
            'Статус': status(date, done, today, why),
            'ID': o['id'],
            'Что сделать': o['что'],
            'Кто': o.get('кто', ''),
            'Кому': o.get('кому', ''),
            'Основание': o.get('основание', ''),
            'Канал': o.get('канал', ''),
            'Шаблон письма': o.get('шаблон', ''),
            'Как посчитан срок': why,
            'Примечание': o.get('срок', {}).get('комментарий', ''),
        })

    if not rows:
        raise SystemExit('В контракт.json нет ни одного обязательства.\n'
                         'Попросите Клода разобрать контракт и ТЗ из папки 1_Документы_контракта.')

    order = {'ПРОСРОЧЕНО': 0, 'СРОЧНО': 1}
    rows.sort(key=lambda r: (r['Срок'] == '', r['Срок'], order.get(r['Статус'].split(' (')[0], 2)))

    for r in rows:
        if r.pop('_справочно'):
            r['Статус'] = 'справочно · ' + r['Статус']

    with open(PLAN, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=';')
        w.writeheader()
        w.writerows(rows)

    print('План пересчитан на %s -> %s' % (today.isoformat(), PLAN))
    все = '--всё' in sys.argv or '--все' in sys.argv
    if все or '--в-word' in sys.argv:
        to_docx(rows, contract, today)
        print('План в Word -> %s' % PLAN_DOCX)
    if все or '--в-confluence' in sys.argv:
        to_confluence(rows, contract, today)
        print('Таблица для Confluence -> %s' % PLAN_WIKI)
    if все or '--в-календарь' in sys.argv:
        n = to_ics(rows, contract, today)
        print('Календарь (%d событий, напоминание за день) -> %s' % (n, PLAN_ICS))
    print()
    свои = [r for r in rows if not r['Статус'].startswith('справочно')]
    hot = [r for r in свои if r['Статус'].startswith(('ПРОСРОЧЕНО', 'СРОЧНО'))]
    soon = [r for r in свои if r['Статус'].startswith('запланировано')][:5]
    wait = [r for r in свои if r['Статус'] == 'ждём событие']
    anytime = [r for r in свои if r['Статус'] == 'по необходимости']
    if hot:
        print('ТРЕБУЕТ ВНИМАНИЯ:')
        for r in hot:
            print('  [%s] %s — %s (%s)' % (r['Срок'] or '—', r['Статус'], r['Что сделать'][:80], r['Основание']))
        print()
    if soon:
        print('БЛИЖАЙШЕЕ:')
        for r in soon:
            print('  %s  %s (%s)' % (r['Срок'], r['Что сделать'][:80], r['Основание']))
        print()
    if wait:
        print('ЖДЁТ СОБЫТИЯ (дату посчитать нельзя, пока не отметите факт в журнале): %d шт.' % len(wait))
        for r in wait[:5]:
            print('  %s — %s' % (r['ID'], r['Как посчитан срок']))
    if anytime:
        print('\nБЕЗ ЖЁСТКОГО СРОКА (делается по ситуации): %d шт.' % len(anytime))
    done_n = sum(1 for r in rows if r['Статус'].endswith('ВЫПОЛНЕНО'))
    справ = len(rows) - len(свои)
    print('\nВсего пунктов: %d (из них %d справочно — обязанности Заказчика), выполнено: %d'
          % (len(rows), справ, done_n))


if __name__ == '__main__':
    main()
