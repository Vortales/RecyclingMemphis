# -*- coding: utf-8 -*-
"""Utility: load cleaned law texts and extract full articles verbatim."""
import re, os, glob

SRC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')

def clean(text: str) -> str:
    text = text.replace('\r', '').replace('\u200b', '').replace('\ufeff', '')
    text = '\n'.join(l.rstrip() for l in text.split('\n'))
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text

_cache = {}

def load(num) -> str:
    num = str(num)
    if num not in _cache:
        path = os.path.join(SRC_DIR, f'{num}.txt')
        with open(path, encoding='utf-8', errors='ignore') as f:
            _cache[num] = clean(f.read())
    return _cache[num]

HEAD_RE = re.compile(r'^(Статья|Глава|ГЛАВА|Раздел|РАЗДЕЛ|ОБЩАЯ ЧАСТЬ|ОСОБЕННАЯ ЧАСТЬ)\b')

def article(num, art) -> str:
    """Return the full verbatim text of article `art` (e.g. '19', '16.12', '17.1') of law `num`.
    Text starts with the 'Статья ...' heading line and ends before the next heading."""
    lines = load(num).split('\n')
    art = str(art)
    pat = re.compile(r'^Статья\s+' + re.escape(art) + r'(?![\d.])\s*[\.\:\s]?', re.U)
    # allow "Статья 16.12 ★★" etc.  Also "Статья 5. " but not "Статья 5.1"
    start = None
    for i, l in enumerate(lines):
        if re.match(r'^Статья\s+' + re.escape(art) + r'(\s|\.(?!\d)|:|$|\*)', l):
            start = i
            break
    if start is None:
        raise KeyError(f'article {art} not found in {num}')
    end = len(lines)
    for j in range(start + 1, len(lines)):
        if HEAD_RE.match(lines[j]):
            end = j
            break
    block = lines[start:end]
    while block and not block[-1].strip():
        block.pop()
    return '\n'.join(block)

if __name__ == '__main__':
    import sys
    print(article(sys.argv[1], sys.argv[2]))
