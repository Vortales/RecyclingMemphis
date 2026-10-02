# -*- coding: utf-8 -*-
"""Развёрнутые редакции законопроектов «для чтения».

К каждому эталонному законопроекту (ACT. 00N — *.docx, формат Приложения №2) строится отдельный документ:
  * текст каждой предлагаемой статьи — дословно как в правой колонке эталона («предложенный вариант»);
  * после каждого абзаца, содержащего ссылку на другую статью/часть/пункт, — справка с ПОЛНЫМ текстом
    упомянутой нормы (действующая редакция из файлов законов либо редакция по соответствующему ACT);
  * для сокращений и специальных терминов — указание акта, статьи и части, где они установлены, с цитатой.
Эталонные файлы не изменяются.
"""
import re
import importlib
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from lawtext import article, load
from common import out, DATE_EN

# ------------------------------------------------------------------------------------------------ справочники
ACTS = [
    ('ACT. 001', 'bill1_pk', 'Изменения в Процессуальный кодекс'),
    ('ACT. 002', 'bill2_orders', 'Ордера, неприкосновенность, Сенат'),
    ('ACT. 003', 'bill3_agencies', 'Законы о государственных органах'),
    ('ACT. 004', 'bill4_uk_ak', 'Изменения в Уголовный и Административный кодексы'),
]

LAW_TITLE = {
    144: 'Уголовный кодекс штата San-Andreas',
    145: 'Процессуальный кодекс штата San-Andreas',
    146: 'Административный кодекс штата San-Andreas',
    152: 'Конституционный Закон «О Правительстве штата San-Andreas»',
    153: 'Конституционный закон «О Сенате штата San-Andreas»',
    155: 'Закон «О деятельности региональных правоохранительных органах»',
    156: 'Закон «О Федеральном Расследовательском Бюро»',
    157: 'Закон «О деятельности Секретной Службы Соединенных Штатов Америки в штате San-Andreas»',
    158: 'Закон «О Национальной Гвардии штата San-Andreas»',
    159: 'Закон «Об Экстренной Медицинской Службе в штате San-Andreas»',
    162: 'Закон «Об обеспечении неприкосновенности государственных служащих штата San-Andreas»',
    163: 'Закон «О чрезвычайном и военном положении»',
    171: 'Закон «О средствах массовой информации»',
    172: 'Закон «О системе ордеров штата San-Andreas»',
    173: 'Закон «Об адвокатской деятельности и адвокатуре в штате San-Andreas»',
    174: 'Конституционный Закон «О деятельности офиса Генерального прокурора штата San-Andreas»',
}
SHORT = {144: 'УК', 145: 'ПК', 146: 'АК', 152: 'КЗ о Правительстве', 153: 'КЗ о Сенате', 155: 'закон о региональных ПОО',
         156: 'закон о FIB', 157: 'закон о USSS', 158: 'закон о SANG', 159: 'закон об EMS', 162: 'закон о неприкосновенности',
         163: 'закон о ЧП/ВП', 171: 'закон о СМИ', 172: 'закон о системе ордеров', 173: 'закон об адвокатуре', 174: 'КЗ об ОГП'}

# распознавание закона по фрагменту названия (в нижнем регистре)
LAW_STEMS = [
    ('процессуальн', 145), ('пк', 145), ('уголовн', 144), ('административн', 146),
    ('о сенате', 153), ('о правительстве', 152), ('региональных правоохранительных', 155),
    ('федеральном расследовательском', 156), ('секретной службы соединен', 157), ('о национальной гвардии', 158),
    ('экстренной медицинской', 159), ('неприкосновенности', 162), ('чрезвычайном', 163), ('массовой информации', 171),
    ('системе ордеров', 172), ('адвокатской', 173), ('генерального прокурора', 174),
]


def law_by_name(name):
    n = name.lower()
    for stem, num in LAW_STEMS:
        if stem in n:
            return num
    return None


def cite(law):
    return '%s (%d.txt)' % (LAW_TITLE[law], law)


# ------------------------------------------------------------------------------------------------ реестр предложенных текстов
PROPOSED = {}     # (law, art) -> (act_no, text, is_new)
ACT_DATA = {}     # act_no -> dict(sections=[(law_num, law_name, [(old, new)])], summary=[...], module=...)


def load_bills():
    for act_no, modname, title in ACTS:
        mod = importlib.import_module(modname)
        secs = getattr(mod, 'sections', None) or [(mod.LAW, mod.rows)]
        data = []
        for law_name, rows in secs:
            law = law_by_name(law_name)
            assert law, law_name
            arts = []
            for old, new in rows:
                if new is None:
                    continue
                art = art_of(new)
                PROPOSED[(law, art)] = (act_no, new, old is None)
                arts.append((art, old, new))
            data.append((law, law_name, arts))
        ACT_DATA[act_no] = dict(sections=data, summary=mod.SUMMARY, title=title)


def art_of(text):
    m = re.match(r'Статья\.?\s*(\d+(?:\.\d+)*)', text)
    return m.group(1)


_ART_CACHE = {}


def law_articles(law):
    if law not in _ART_CACHE:
        arts = []
        for l in load(law).split('\n'):
            m = re.match(r'^Статья\.?\s*(\d+(?:\.\d+)*)', l)
            if m:
                arts.append(m.group(1))
        _ART_CACHE[law] = arts
    return _ART_CACHE[law]


def art_key(a):
    return tuple(int(x) for x in a.split('.'))


# ------------------------------------------------------------------------------------------------ извлечение частей и пунктов
PART_RE = re.compile(r'^ч\.?\s*(\d+)(?![\d.]\d)')
NUM_RE = re.compile(r'^(\d+)\.\s')
POINT_RE = re.compile(r'^([а-яё](?:\.\d+)?)\)')


def split_parts(lines):
    """-> [(part_no|None, [lines])] ; первая группа (None) — заголовок и текст до первой части."""
    use_num = not any(PART_RE.match(l) for l in lines[1:])
    groups = [(None, [lines[0]])]
    for l in lines[1:]:
        m = PART_RE.match(l) or (NUM_RE.match(l) if use_num else None)
        if m:
            groups.append((m.group(1), [l]))
        else:
            groups[-1][1].append(l)
    return groups


def extract(text, part=None, points=None):
    """Вернуть строки статьи: всю статью / часть / пункты части (с заголовком статьи для контекста)."""
    lines = text.split('\n')
    head = lines[0]
    if part is None and not points:
        return lines, True
    groups = split_parts(lines)
    if part is not None:
        sel = [g for g in groups if g[0] == str(part)]
        if not sel:
            return [head, '[часть %s в статье не найдена]' % part], False
        plines = sel[0][1]
    else:
        plines = [l for g in groups for l in g[1]][1:]
    if not points:
        return [head] + plines, True
    # пункты
    res = [head]
    if part is not None and plines:
        first = plines[0]
        res.append(first if len(first) <= 220 else first[:200] + '…')
        body = plines[1:]
    else:
        body = plines
    wanted = [p.strip('«»').lower() for p in points]
    cur = None
    found = set()
    for l in body:
        m = POINT_RE.match(l)
        if m:
            cur = m.group(1).lower()
        elif PART_RE.match(l):
            cur = None
        if cur in wanted:
            res.append(l)
            found.add(cur)
    missing = [w for w in wanted if w not in found]
    if missing:
        res.append('[пункт(ы) %s не найдены]' % ', '.join('«%s»' % m for m in missing))
    return res, not missing


# ------------------------------------------------------------------------------------------------ поиск ссылок в тексте
P_LAW = (r'(?P<law>настоящ\w+\s+(?:[Кк]одекса|[Зз]акона)|Процессуального\s+[Кк]одекса(?:\s+штата\s+(?:San-Andreas|Сан-Андреас))?'
         r'|Уголовного\s+[Кк]одекса?(?:\s+штата\s+San-Andreas)?|Административного\s+[Кк]одекса(?:\s+штата\s+San-Andreas)?'
         r'|(?:Конституционн\w+\s+)?[Зз]акон\w*\s+«[^»]+»|ПК)')
ART = r'(?P<art>\d+(?:\.\d+)*)'
ARTS = r'(?P<arts>\d+(?:\.\d+)*(?:\s*(?:[–-]|,|\s+и)\s*\d+(?:\.\d+)*)*)'
PTS = r'(?P<pts>«[^»]+»(?:\s*(?:,|и)\s*«[^»]+»)*)'
STAT = r'(?:стат\w+|ст\.)'
CH = r'(?:част\w+|ч\.)'

PATTERNS = [
    ('pt_part_art', re.compile(rf'(?:[Пп]ункт\w*|п\.)\s+{PTS}\s+{CH}\s*(?P<part>\d+)\s+{STAT}\s*{ART}(?:\s+{P_LAW})?')),
    ('pt_part_self', re.compile(rf'(?:[Пп]ункт\w*|п\.)\s+{PTS}\s+{CH}\s*(?P<part>\d+)\s+настоящей\s+статьи')),
    ('part_art', re.compile(rf'{CH}\s*(?P<part>\d+)\s+{STAT}\s*{ART}(?:\s+{P_LAW})?')),
    ('art_part', re.compile(rf'{STAT}\s*{ART}\s+(?:част\w+|ч\.?)\s*(?P<part>\d+)(?:\s+{P_LAW})?')),
    ('pts_parts_self', re.compile(rf'[Пп]ункт\w*\s+(?P<pts2>[а-я]\)(?:\s*(?:,|и)\s*[а-я]\))*)\s+(?P<pp>{CH}\s*\d+(?:\s*(?:,|и)\s*(?:{CH})?\s*\d+)*)\s+настоящей\s+статьи')),
    ('part_self', re.compile(rf'{CH}\s*(?P<parts>\d+(?:\s*(?:,|и)\s*(?:{CH})?\s*\d+)*)\s+настоящей\s+статьи')),
    ('arts', re.compile(rf'{STAT}\s*{ARTS}(?:\s+{P_LAW})?')),
    ('pt_bare', re.compile(r'(?:в\s+)?пункте\s+(?P<pt>[а-я])\)?(?=\s)')),
]


def parse_arts(s, law):
    """'6 и 7' -> ['6','7']; '17.1–17.3' -> все статьи закона в диапазоне."""
    s = s.strip()
    m = re.match(r'^(\d+(?:\.\d+)*)\s*[–-]\s*(\d+(?:\.\d+)*)$', s)
    if m:
        a, b = art_key(m.group(1)), art_key(m.group(2))
        allarts = sorted(set(law_articles(law)) | {ar for (lw, ar) in PROPOSED if lw == law}, key=art_key)
        return [x for x in allarts if a <= art_key(x) <= b] or [m.group(1), m.group(2)]
    return re.split(r'\s*(?:,|и)\s*', s)


def find_refs(text, own_law, own_art, cur_part):
    refs = []
    taken = []

    def free(s, e):
        return all(e <= a or s >= b for a, b in taken)

    for kind, rx in PATTERNS:
        for m in rx.finditer(text):
            s, e = m.span()
            if not free(s, e):
                continue
            g = m.groupdict()
            law_s = g.get('law')
            if law_s is None or law_s.startswith('настоящ'):
                law = own_law
            else:
                law = law_by_name(law_s)
                if law is None:
                    continue
            if kind == 'pt_part_art':
                items = [(law, g['art'], g['part'], re.findall(r'«([^»]+)»', g['pts']))]
            elif kind == 'pt_part_self':
                items = [(own_law, own_art, g['part'], re.findall(r'«([^»]+)»', g['pts']))]
            elif kind in ('part_art', 'art_part'):
                items = [(law, g['art'], g['part'], None)]
            elif kind == 'part_self':
                items = [(own_law, own_art, p, None) for p in re.findall(r'\d+', g['parts'])]
            elif kind == 'pts_parts_self':
                pts = re.findall(r'([а-я])\)', g['pts2'])
                items = [(own_law, own_art, p, pts) for p in re.findall(r'\d+', g['pp'])]
            elif kind == 'arts':
                items = [(law, a, None, None) for a in parse_arts(g['arts'], law)]
            else:  # pt_bare
                if cur_part is None:
                    continue
                items = [(own_law, own_art, cur_part, [g['pt']])]
            items = [it for it in items if not (it[0] == own_law and it[1] == own_art and it[2] is None)]
            if not items:
                continue
            taken.append((s, e))
            refs.append(dict(span=(s, e), text=text[s:e], items=items))
    refs.sort(key=lambda r: r['span'][0])
    return refs


# ------------------------------------------------------------------------------------------------ термины и обозначения
def _q(law, art, part=None, points=None, grep=None):
    return dict(law=law, art=art, part=part, points=points, grep=grep)


TERMS = [
    # Единый реестр
    (r'Едином реестре|Единый реестр', 'Единый реестр штата Сан-Андреас', 'официальный государственный информационный ресурс штата San-Andreas, на котором размещаются актуальные списки должностных лиц, членов Сената, лицензий, пропусков на ЗОТ и кадровых назначений (раздел официального портала «Единый реестр штата San-Andreas»)', None),
    # ордера
    (r'\bAR\b|Access to Raid', 'AR (Access to Raid)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 1 п. «а» — действующая редакция; ACT. 002 уточняет формулировку', _q(172, '6', '1', ['а'])),
    (r'\bAC\b|Access to Cordon', 'AC (Access to Cordon)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 2 п. «а» — действующая редакция; ACT. 002 уточняет формулировку', _q(172, '6', '2', ['а'])),
    (r'\bDT\b|Detention and Transfer', 'DT (Detention and Transfer)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 3 п. «а»', _q(172, '6', '3', ['а'])),
    (r'\bAS\b|Arrest and Search', 'AS (Arrest and Search)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 4 п. «а»', _q(172, '6', '4', ['а'])),
    (r'\bAI\b|Access to Investigation', 'AI (Access to Investigation)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 5 п. «а»', _q(172, '6', '5', ['а'])),
    (r'\bML\b|Martial Law', 'ML (Martial Law)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 6 п. «а»', _q(172, '6', '6', ['а'])),
    (r'\bSE\b|State of Emergency', 'SE (State of Emergency)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 7 п. «а»', _q(172, '6', '7', ['а'])),
    (r'\bSI\b|Suspension of Immunity', 'SI (Suspension of Immunity)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 8 п. «а»', _q(172, '6', '8', ['а'])),
    (r'\bRI\b|Removal of Immunity', 'RI (Removal of Immunity)', 'установлено законом «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 9 п. «а»', _q(172, '6', '9', ['а'])),
    # подразделения
    (r'\bMP\b|Military Police', 'MP (Military Police)', 'подразделение определено законом «О Национальной Гвардии штата San-Andreas» (158.txt), ст. 23 п. «д»', _q(158, '23', None, ['д'])),
    (r'\bIB\b|Intelligence Branch', 'IB (Intelligence Branch)', 'подразделение определено законом «О Федеральном Расследовательском Бюро» (156.txt), ст. 21 ч. 1', _q(156, '21', '1')),
    (r'\bCID\b|Criminal Investigative Division', 'CID (Criminal Investigative Division)', 'подразделение определено законом «О Федеральном Расследовательском Бюро» (156.txt), ст. 19 (заголовок статьи)', _q(156, '19', None)),
    (r'Detective Bureau|\bDB\b', 'Detective Bureau (DB)', 'отдельного определения в базе нет; подразделение упоминается в ПК (145.txt), ст. 18 ч. 1 п. «з» (действующая редакция)', _q(145, '18', '1', ['з'])),
    # процессуальные понятия
    (r'боло-розыск', 'боло-розыск', 'установлен Процессуальным кодексом (145.txt), ст. 39', _q(145, '39', None)),
    (r'[Фф]едеральн\w+ розыск', 'федеральный розыск', 'установлен Процессуальным кодексом (145.txt), ст. 38', _q(145, '38', None)),
    (r'субъект\w* задержания', 'субъект задержания', 'понятие раскрыто в ПК (145.txt), ст. 18 ч. 1 (перечень участников задержания)', _q(145, '18', '1')),
    (r'[Пп]ервичн\w+ обыск', 'первичный обыск', 'понятие раскрыто в ПК (145.txt), ст. 17 ч. 2 п. «в» и Примечание к статье (что изымается при первичном обыске)', _q(145, '17', '2', ['в'], grep=r'^Примечание: Первичный обыск')),
    (r'[Пп]ривод\b', 'привод', 'осуществляется по ордеру DT — закон «О системе ордеров штата San-Andreas» (172.txt), ст. 6 ч. 3 п. «а»', _q(172, '6', '3', ['а'])),
    (r'статус\w* неприкосновенности|обладающ\w+ статусом неприкосновенности', 'статус неприкосновенности', 'круг лиц определён законом «Об обеспечении неприкосновенности государственных служащих штата San-Andreas» (162.txt), ст. 2 ч. 1', _q(162, '2', '1')),
    (r'объект\w* защитной миссии', 'объект защитной миссии', 'определён законом «О деятельности Секретной Службы Соединенных Штатов Америки в штате San-Andreas» (157.txt), ст. 4 ч. 2', _q(157, '4', '2')),
    (r'[Сс]тарш\w+ прокурор', 'старший прокурор', 'должность установлена Конституционным Законом «О деятельности офиса Генерального прокурора штата San-Andreas» (174.txt), ст. 15 ч. 1–2; право авторизации ордеров — ст. 19 ч. 14', _q(174, '15', '1')),
    (r'Глав\w+ Департамента Национальной Безопасности', 'Глава Департамента Национальной Безопасности', 'должность следует из Конституционного Закона «О Правительстве штата San-Andreas» (152.txt), ст. 23 ч. 1 (USSS и SANG подчиняются Главе Департамента национальной безопасности); единое написание — ст. 2.1 ч. 2 ПК (ACT. 001)', _q(152, '23', '1')),
    (r'Глав\w+ Департамента Здравоохранения', 'Глава Департамента Здравоохранения', 'должность установлена Конституционным Законом «О Правительстве штата San-Andreas» (152.txt), ст. 31 ч. 2', _q(152, '31', '2')),
    (r'Глав\w+ Коллегии [Аа]двокатов', 'Глава Коллегии адвокатов', 'должность установлена законом «Об адвокатской деятельности и адвокатуре в штате San-Andreas» (173.txt), ст. 1 ч. 5', _q(173, '1', '5')),
    (r'Спикер\w*', 'Спикер Сената', 'установлен Конституционным законом «О Сенате штата San-Andreas» (153.txt), ст. 2', _q(153, '2', None)),
    (r'официальн\w+ интернет-портал', 'официальный интернет-портал', 'упоминается в Конституционном законе «О Сенате штата San-Andreas» (153.txt), ст. 28', _q(153, '28', None)),
    (r'Верховн\w+ [Сс]уд', 'Верховный суд', 'отдельного закона о суде в базе нет; Верховный суд упоминается в законе «О системе ордеров» (172.txt), ст. 7 ч. 4 и в КЗ об ОГП (174.txt), ст. 13', _q(174, '13', None)),
    (r'[Кк]омитет\w* Сената по этике', 'комитет Сената по этике', 'ВНИМАНИЕ: в Конституционном законе «О Сенате» (153.txt) такой комитет не упоминается; формулировка сохранена из действующей редакции ст. 6 ч. 5 закона «О системе ордеров» (172.txt)', None),
    (r'чрезвычайн\w+ (?:или|и) военн\w+ положени|чрезвычайного положения|военного положения|ЧП/ВП', 'чрезвычайное / военное положение', 'режимы установлены законом «О чрезвычайном и военном положении» (163.txt), глава II п. 2.1 (ЧП) и глава III (ВП); вводятся ордерами SE и ML — закон «О системе ордеров» (172.txt), ст. 6 ч. 6–7', None),
    (r'контртеррористическ\w+ операци\w+', 'контртеррористическая операция', 'отдельного определения в базе нет; понятие используется в законе о FIB (156.txt), ст. 14 ч. 1 п. «д» (оцепление) и в законе о SANG (158.txt), ст. 7 ч. 3 п. «в» (оцепление), действующие редакции', _q(156, '14', '1', ['д'])),
    (r'устав\w* SANG|штатной структур\w+ \(уставу\)|порядков\w+ ранг', 'устав SANG / порядковый ранг', 'уставы (внутренние положения) SANG утверждаются Главой Департамента Национальной Безопасности — закон «О Национальной Гвардии штата San-Andreas» (158.txt), ст. 3 ч. 1; иерархия званий — ст. 1 п. «е» того же закона', _q(158, '3', '1')),
    (r'Форт-Занкудо|Форта-Занкудо|Кайо-Перико|военн\w+ авианос\w+', 'охраняемые SANG объекты (Форт-Занкудо, авианосец, Кайо-Перико)', 'места дислокации установлены законом «О Национальной Гвардии штата San-Andreas» (158.txt), ст. 2 ч. 1; контрольно-пропускной режим — ст. 2 ч. 5–6', _q(158, '2', '1')),
    (r'пропускн\w+ режим', 'пропускной режим', 'контрольно-пропускной режим на территории авианосца и Форта-Занкудо — закон «О Национальной Гвардии штата San-Andreas» (158.txt), ст. 2 ч. 5 и ч. 6', _q(158, '2', '5')),
    (r'воинск\w+ преступлени\w+', 'воинские преступления', 'ВНИМАНИЕ: в УК (144.txt) отдельной главы о воинских преступлениях нет; термин опирается на функции MP по контролю военнослужащих — закон «О Национальной Гвардии штата San-Andreas» (158.txt), ст. 23 п. «д» и ст. 2 ч. 1 п. «ж» (поиск дезертиров)', _q(158, '23', None, ['д'])),
    (r'\bКПЗ\b', 'КПЗ', 'место содержания под стражей; упоминается в ПК (145.txt), ст. 25 ч. 1 п. «д» и ст. 27 п. «д»', _q(145, '27', None, ['д'])),
    (r'Федеральн\w+ Тюрьм\w+', 'Федеральная Тюрьма', 'место отбывания лишения свободы; упоминается в УК (144.txt), ст. 16.15 (санкция)', None),
    (r'ориентировк\w*', 'ориентировка', 'основание задержания/осмотра по ПК (145.txt), ст. 16 ч. 2 п. «ж» и ст. 30 п. «а»', _q(145, '16', '2', ['ж'])),
    (r'Коллеги\w+ адвокатов', 'Коллегия адвокатов', 'определена законом «Об адвокатской деятельности и адвокатуре в штате San-Andreas» (173.txt), ст. 1 ч. 1', _q(173, '1', '1')),
]
TERMS = [(re.compile(p), name, src, q) for p, name, src, q in TERMS]

ORGANS = [
    ('GOV', 'Правительство штата San-Andreas', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001) впервые — в действующих законах сокращение не использовалось; орган определён КЗ «О Правительстве штата San-Andreas» (152.txt)'),
    ('ОГП', 'Офис Генерального прокурора штата San-Andreas', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); орган определён КЗ «О деятельности офиса Генерального прокурора штата San-Andreas» (174.txt), ст. 1'),
    ('LSPD', 'Los Santos Police Department (Полицейский департамент города Лос-Сантос)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); закрепляется как «(далее — LSPD)» в ст. 1 ч. 1 закона о региональных ПОО (155.txt) в редакции ACT. 003; ранее в законах использовалось наряду с «Полицейский Департамент»'),
    ('LSCSD', 'Los Santos County Sheriff Department (Департамент шерифа округа Блэйн)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); закрепляется как «(далее — LSCSD)» в ст. 1 ч. 1 закона о региональных ПОО (155.txt) в редакции ACT. 003; ранее — «Департамент Шерифа», «Sheriff Department», «LSSD»'),
    ('FIB', 'Federal Investigation Bureau (Федеральное Расследовательское Бюро)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); закрепляется как «(далее — FIB)» в ст. 1 ч. 1 закона о FIB (156.txt) в редакции ACT. 003; в действующей редакции — «(далее - ФРБ)»'),
    ('USSS', 'United States Secret Service (Секретная Служба Соединённых Штатов Америки)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); действующая редакция закона о USSS (157.txt), ст. 1: «(далее — USSS или Секретная служба)»'),
    ('SANG', 'San Andreas National Guard (Национальная Гвардия штата San-Andreas)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); действующая редакция закона о SANG (158.txt), ст. 1: «(далее SANG)»'),
    ('EMS', 'Emergency Medical Service (Экстренная Медицинская Служба)', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); действующая редакция закона об EMS (159.txt), ст. 1 ч. 1: «далее — EMS»'),
    ('WN', 'Weazel News', 'вводится ст. 2.1 ч. 1 ПК (ACT. 001); в действующей редакции ПК (145.txt) сокращение употребляется в ст. 17.2 («на территории WN»); полное наименование — закон о СМИ (171.txt), ст. 14'),
]

# ------------------------------------------------------------------------------------------------ оформление docx
GRAY = RGBColor(0x40, 0x40, 0x40)
BLUE = RGBColor(0x1F, 0x3A, 0x93)
RED = RGBColor(0x9C, 0x00, 0x06)


def _shade(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    pPr.append(shd)
    bdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single'); left.set(qn('w:sz'), '18'); left.set(qn('w:space'), '4'); left.set(qn('w:color'), '7F7F7F')
    bdr.append(left)
    pPr.append(bdr)


def block_par(doc, fill='F2F2F2', size=10, indent=1.0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    _shade(p, fill)
    p._size = size
    return p


def add_run(p, text, bold=False, italic=False, size=None, color=None, underline=False):
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color
    r.underline = underline
    return r


class ReaderDoc:
    def __init__(self, act_no, title):
        self.act_no = act_no
        self.doc = Document()
        st = self.doc.styles['Normal']
        st.font.name = 'Times New Roman'
        st.font.size = Pt(12)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        for s in self.doc.sections:
            s.left_margin = Cm(2); s.right_margin = Cm(1.5); s.top_margin = Cm(1.5); s.bottom_margin = Cm(1.5)
        self.expanded = {}   # ключ ссылки -> (art_label, размер)
        self.title = title

    # --- заголовки
    def h(self, text, size=12, center=False, before=12, after=6):
        p = self.doc.add_paragraph()
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        add_run(p, text, bold=True, size=size)
        return p

    def para(self, text, italic=False, size=None, after=4):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        add_run(p, text, italic=italic, size=size)
        return p

    # --- текст статьи с подчёркнутыми ссылками
    def article_line(self, line, spans, bold=False):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        pos = 0
        for s, e in spans:
            if s > pos:
                add_run(p, line[pos:s], bold=bold)
            add_run(p, line[s:e], bold=bold, underline=True, color=BLUE)
            pos = e
        if pos < len(line):
            add_run(p, line[pos:], bold=bold)
        return p

    # --- справка по ссылке
    def ref_block(self, ref_text, law, art, part, points, own_law):
        key = (law, art, part, tuple(points) if points else None)
        if (law, art) in PROPOSED:
            act, text, is_new = PROPOSED[(law, art)]
            if act == self.act_no:
                edition = 'в редакции настоящего законопроекта (%s)' % act
            else:
                edition = 'в редакции %s' % act
            if is_new:
                edition = 'новая статья, ' + edition
        else:
            try:
                text = article(law, art)
            except KeyError:
                text = None
            edition = 'действующая редакция'
        where = 'ст. %s' % art
        if part is not None:
            where = 'ч. %s ' % part + where
        if points:
            where = 'п. %s ' % ', '.join('«%s»' % x for x in points) + where
        label_law = '%s (%d.txt)' % (LAW_TITLE[law], law)
        p = block_par(self.doc)
        add_run(p, '▸ Ссылка «%s» → %s — %s, %s' % (ref_text, where, label_law, edition), bold=True, italic=True, size=10, color=GRAY)
        if text is None:
            add_run(p, ' — СТАТЬЯ НЕ НАЙДЕНА В БАЗЕ', bold=True, size=10, color=RED)
            return
        lines, ok = extract(text, part, points)
        body = '\n'.join(lines)
        if key in self.expanded and len(body) > 1500:
            add_run(p, ' — полный текст приведён выше (справка к %s).' % self.expanded[key], italic=True, size=10, color=GRAY)
            return
        add_run(p, ':', bold=True, italic=True, size=10, color=GRAY)
        for l in lines:
            q = block_par(self.doc)
            if l.startswith('[') and l.endswith(']'):
                add_run(q, l, bold=True, size=10, color=RED)
            else:
                add_run(q, l, size=10)
        self.expanded.setdefault(key, self.cur_art_label)

    # --- справка по терминам
    def term_block(self, entries):
        p = block_par(self.doc, fill='EAF1FB')
        add_run(p, '▸ Обозначения и термины в этом абзаце:', bold=True, italic=True, size=10, color=GRAY)
        for name, src, q in entries:
            e = block_par(self.doc, fill='EAF1FB')
            add_run(e, '• %s — ' % name, bold=True, size=10)
            add_run(e, src, size=10)
            if q:
                law, art, part, points = q['law'], q['art'], q['part'], q['points']
                if (law, art) in PROPOSED and PROPOSED[(law, art)][0] != self.act_no:
                    text = PROPOSED[(law, art)][1]
                    ed = ' (текст — в редакции %s)' % PROPOSED[(law, art)][0]
                elif (law, art) in PROPOSED:
                    text = PROPOSED[(law, art)][1]
                    ed = ' (текст — в редакции настоящего законопроекта)'
                else:
                    try:
                        text = article(law, art)
                    except KeyError:
                        text = None
                    ed = ''
                if text is None:
                    add_run(e, ' — СТАТЬЯ НЕ НАЙДЕНА', bold=True, size=10, color=RED)
                    continue
                lines, ok = extract(text, part, points)
                if q.get('grep'):
                    lines += [l for l in text.split('\n') if re.search(q['grep'], l) and l not in lines]
                add_run(e, ed + ':', size=10)
                for l in lines[:12]:
                    q2 = block_par(self.doc, fill='EAF1FB', indent=1.5)
                    add_run(q2, l if len(l) <= 700 else l[:700] + ' […]', size=10, italic=True)
                if len(lines) > 12:
                    q2 = block_par(self.doc, fill='EAF1FB', indent=1.5)
                    add_run(q2, '[…] (приведены первые строки; полный текст — в указанной статье)', size=10, italic=True)

    def spacer(self):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)


# ------------------------------------------------------------------------------------------------ сборка документа
def build(act_no):
    data = ACT_DATA[act_no]
    rd = ReaderDoc(act_no, data['title'])
    doc = rd.doc
    rd.h('%s — %s' % (act_no, data['title']), size=14, center=True, before=0)
    rd.h('РАЗВЁРНУТАЯ РЕДАКЦИЯ ДЛЯ ЧТЕНИЯ', size=12, center=True, before=0)
    rd.para('Справочный документ к законопроекту %s от %s. Не является официальным текстом законопроекта: эталон — файл «%s — %s.docx» (формат Приложения №2). '
            'Здесь приведён только «предложенный вариант» каждой статьи, дословно совпадающий с правой колонкой эталона.' % (act_no, DATE_EN, act_no, data['title']), italic=True, size=11)
    rd.para('Как читать: ссылки на другие статьи, части и пункты в тексте подчёркнуты. Сразу после абзаца со ссылкой приведена справка ▸ с полным текстом упомянутой нормы: '
            '«действующая редакция» — текст из файла закона; «в редакции ACT. 00N» — текст, каким он станет после принятия соответствующего законопроекта '
            '(нормы, вводимые пакетом ACT. 001–004, отмечены как новые). Отдельно выделены сокращения и специальные термины с указанием акта, статьи и части, где они установлены. '
            'Если текст длинной нормы уже приводился в этом документе, повторно даётся отсылка к первой справке.', italic=True, size=11)

    rd.h('Обозначения органов государственной власти (ст. 2.1 ПК, вводится ACT. 001)', size=12)
    t = doc.add_table(rows=1, cols=3)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    for c, txt in zip(hdr, ('Обозначение', 'Полное наименование', 'Где установлено / первичный источник')):
        c.text = ''
        add_run(c.paragraphs[0], txt, bold=True, size=10)
    for abbr, full, src in ORGANS:
        cells = t.add_row().cells
        for c, txt in zip(cells, (abbr, full, src)):
            c.text = ''
            add_run(c.paragraphs[0], txt, size=10, bold=(c is cells[0]))

    rd.h('Акты, цитируемые в этом документе', size=12)
    used_laws = sorted({law for law, _, _ in data['sections']})
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = 'Table Grid'
    for c, txt in zip(t2.rows[0].cells, ('Файл базы', 'Наименование акта')):
        c.text = ''
        add_run(c.paragraphs[0], txt, bold=True, size=10)
    for law in sorted(LAW_TITLE):
        cells = t2.add_row().cells
        cells[0].text = ''; cells[1].text = ''
        add_run(cells[0].paragraphs[0], '%d.txt' % law, size=10)
        add_run(cells[1].paragraphs[0], LAW_TITLE[law] + (' — изменяется настоящим законопроектом' if law in used_laws else ''), size=10, bold=law in used_laws)

    rd.h('SECTION 1. СУТЬ ПРАВКИ.', size=12, before=14)
    for s in data['summary']:
        rd.para(s)

    n = 2
    for law, law_name, arts in data['sections']:
        rd.h('SECTION. %d. ВНЕСЕНИЕ ПРАВКИ В %s (%d.txt)' % (n, law_name, law), size=12, before=14)
        n += 1
        p = block_par(doc, fill='EAF1FB')
        add_run(p, '▸ Обозначения органов LSPD, LSCSD, FIB, USSS, SANG, EMS, WN, ОГП, GOV во всех статьях ниже — по ст. 2.1 ч. 1 ПК (вводится ACT. 001); первичные источники каждого обозначения — в таблице в начале документа.', italic=True, size=10, color=GRAY)
        for art, old, new in arts:
            lines = new.split('\n')
            rd.cur_art_label = 'ст. %s %s' % (art, SHORT[law])
            status = 'НОВАЯ СТАТЬЯ' if old is None else 'изменяемая статья (полный текст в новой редакции)'
            if len(lines[0]) <= 110:
                rd.h('%s — %s' % (lines[0], status), size=12, before=12, after=4)
            else:
                hp = rd.h(lines[0], size=12, before=12, after=0)
                sp = doc.add_paragraph()
                sp.paragraph_format.space_after = Pt(4)
                add_run(sp, '— ' + status, bold=True, italic=True, size=10, color=GRAY)
            seen_terms = set()
            cur_part = None
            for line in lines[1:]:
                if not line.strip():
                    rd.spacer()
                    continue
                m = PART_RE.match(line)
                if m:
                    cur_part = m.group(1)
                refs = find_refs(line, law, art, cur_part)
                spans = [r['span'] for r in refs]
                rd.article_line(line, spans)
                for r in refs:
                    for (tlaw, tart, tpart, tpts) in r['items']:
                        rd.ref_block(r['text'], tlaw, tart, tpart, tpts, law)
                entries = []
                for rx, name, src, q in TERMS:
                    if name in seen_terms:
                        continue
                    if rx.search(line):
                        seen_terms.add(name)
                        entries.append((name, src, q))
                if entries:
                    rd.term_block(entries)
    path = out('%s — Развёрнутая редакция для чтения.docx' % act_no)
    doc.save(path)
    return path


if __name__ == '__main__':
    import sys
    load_bills()
    which = sys.argv[1:] or [a for a, _, _ in ACTS]
    for act_no in which:
        print(build(act_no))
