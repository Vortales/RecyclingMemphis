from pathlib import Path
import re,json,hashlib
root=Path(__file__).resolve().parents[1];data={}
for p in sorted((root/'Законы').glob('*.txt')):
 if p.stem=='ТЗ':continue
 s=p.read_text(encoding='utf-8-sig').replace('\u200b','')
 starts=[m for m in re.finditer(r'(?m)^[ \t]*(Статья\s+\d+(?:\.\d+)*[.:]?[^\n]*)',s) if not re.match(r'Статья\s+\d+\s+—',m.group(1))]
 articles={}
 for i,m in enumerate(starts):
  n=re.match(r'Статья\s+(\d+(?:\.\d+)*)',m.group(1))[1]
  text=s[m.start():starts[i+1].start() if i+1<len(starts) else len(s)].strip()
  text=re.split(r'(?m)^\s*(?:Глава|ГЛАВА|Раздел|РАЗДЕЛ)\s+[IVX\d]',text)[0].strip()
  articles[n]=(articles[n]+'\n\n'+text) if n in articles else text
 data[p.stem]={'title':s.splitlines()[0], 'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'text':s,'articles':articles}
(root/'Анализ'/'база.json').write_text(json.dumps(data,ensure_ascii=False,indent=2))
print('Sources',len(data),'articles',sum(len(v['articles']) for v in data.values()))
