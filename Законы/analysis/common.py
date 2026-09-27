# -*- coding: utf-8 -*-
import os
from lawtext import article

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Законопроекты')
os.makedirs(OUT_DIR, exist_ok=True)
DATE_EN = 'SEPTEMBER 27, 2026'


def sub(text, old, new, count=1):
    """Safe replacement: `old` must occur exactly `count` times."""
    n = text.count(old)
    if n != count:
        raise ValueError('expected %d occurrence(s), found %d for: %r' % (count, n, old[:80]))
    return text.replace(old, new)


def out(name):
    return os.path.join(OUT_DIR, name)


# unified wording of the "state protection" clause (ЦЕЛЬ 1)
def protection_clause(who, extra=''):
    return ('При исполнении %s служебных обязанностей не допускаются его привод, задержание, личный обыск и обыск '
            'его вещей, а также обыск личного и используемого им транспорта, за исключением случаев, предусмотренных '
            'частью 4 статьи 19 Процессуального кодекса штата San-Andreas: при наличии ордера или судебного акта; '
            'когда он застигнут при совершении преступления или непосредственно после его совершения; когда его '
            'действия создают непосредственную угрозу жизни или здоровью окружающих; по требованию лиц, уполномоченных '
            'на это Процессуальным кодексом штата San-Andreas%s. Задержание производится в порядке статьи 19 '
            'Процессуального кодекса штата San-Andreas.' % (who, extra))
