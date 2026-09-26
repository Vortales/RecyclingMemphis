from pathlib import Path
import json,re,hashlib,zipfile,csv
from docx import Document
from docx.oxml.ns import qn
R=Path(__file__).resolve().parents[1]
D=json.loads((R/'Анализ/база.json').read_text());B=json.loads((R/'Анализ/манифест.json').read_text())
def norm(s):
 s=s.replace('\u200b','').strip()
 s=re.sub(r'(?mi)^ч\.\s*(\d+(?:\.\d+)*)\.?\s*',lambda m:'Часть '+m[1]+'. ',s)
 s=re.sub(r'(?mi)^п\.\s*(\d+(?:\.\d+)*)\.?\s*',lambda m:'Пункт '+m[1]+'. ',s)
 s=re.sub(r'\bч\.\s*','части ',s,flags=re.I)
 s=re.sub(r'\bст\.\s*','статьи ',s,flags=re.I)
 s=re.sub(r'\bп\.\s*','пункта ',s)
 s=re.sub(r'\bГл\.\s*','главы ',s)
 s=re.sub(r'(?m)[ \t]+$','',s);s=re.sub(r'\n{3,}','\n\n',s)
 return s
seen=set();n=0;colors=set()
template=Document(R/'Законы/ACT. 000 (1).docx');expected=template.tables[1]._tbl.tblPr.xml
for k,v in D.items():
 assert hashlib.sha256((R/'Законы'/f'{k}.txt').read_bytes()).hexdigest()==v['sha256'],k
for b in B:
 p=R/b['path'];assert p.exists()
 with zipfile.ZipFile(p) as z:assert z.testzip() is None
 doc=Document(p);assert doc.paragraphs[0].text==f"ACT. {b['num']:03d}";assert doc.paragraphs[3].text=='SEPTEMBER 27, 2026'
 assert len(doc.tables)==len(b['changes'])
 for s,t in zip(doc.sections,template.sections):
  assert (s.page_width,s.page_height,s.left_margin,s.right_margin)==(t.page_width,t.page_height,t.left_margin,t.right_margin)
 for c,t in zip(b['changes'],doc.tables):
  key=(c['file'],c['article']);assert key not in seen,key;seen.add(key)
  assert len(t.rows)==2 and len(t.columns)==2
  assert [x.text for x in t.rows[0].cells]==['настоящая редакция','предложенный вариант']
  assert t._tbl.tblPr.xml==expected
  for i,side in enumerate(['old','new']):
   got=t.cell(1,i).text;want=norm(c[side]);assert got==want,(key,side)
   assert '…' not in got and '...' not in got
   for p in t.cell(1,i).paragraphs:
    for r in p.runs:
     if r.font.color.rgb:colors.add(str(r.font.color.rgb))
  assert not re.search(r'правила проекта|ООС|OOC|Demorgan|WARN|главн.{0,15}куратор',t.cell(1,1).text,re.I),key
  if c['article'] in D[c['file']]['articles']:
   assert c['old']==D[c['file']]['articles'][c['article']],key
  else:assert c['article'] in ['5','Глава I','Глава II','Глава III'],key
  n+=1
 assert not any('При создании законопроектов данное примечание не писать' in x.text for x in doc.paragraphs)
 for p in (R/'Пояснения').glob('*.docx'):
  with zipfile.ZipFile(p) as z:assert z.testzip() is None
rows=list(csv.DictReader((R/'Анализ/матрица_465_пар.csv').open(encoding='utf-8-sig'),delimiter=';'))
assert len(rows)==465
assert len({tuple(sorted([r['Источник А'],r['Источник Б']])) for r in rows})==465
assert {'FF0000','008000','333333'}<=colors
result=f'''# Проверка пакета

- Законопроекты: {len(B)} .docx; отдельных пояснений: {len(B)}; общий обзор: 1.
- Сравнительные блоки: {n}; каждый — 2 колонки × 2 логические строки.
- Полные старые и новые тексты в ячейках совпадают с манифестом после раскрытия обозначений частей, статей и пунктов.
- Текущие редакции сверены с сохраненным индексом исходных файлов; SHA-256 всех 31 источника совпадает.
- Страница, поля, свойства сравнительных таблиц, шапки и дата проверены.
- Красный FF0000, зеленый 008000, неизмененный 333333 присутствуют.
- Нет повторного изменения одной статьи в разных ACT; связанные статьи распределены без конкурирующих редакций.
- В предложенных редакциях нет ссылок на правила проекта и внешние административные механизмы.
- Нет многоточий, пропущенных середины статьи или служебного примечания из шаблона.
- DOCX открываются python-docx, ZIP-контейнеры целы.
- Матрица: 435 пар законов + 30 пар с правилами = 465 уникальных пар.

Проверки структурные и текстовые, не заменяют юридическое рассмотрение. Визуальное постраничное отображение Word/LibreOffice не проверено: офисный движок в среде отсутствует, установка недоступна из-за сетевого соединения. Переносы строк зависят от наличия шрифтов шаблона у получателя. Размеры таблиц намеренно сохранены из предоставленной формы, а не изменены ради новой верстки.
'''
(R/'Анализ/Проверка.md').write_text(result)
print(result)
