import base64,pathlib
P=pathlib.Path(__file__).parent
page=(P/'page.html').read_text()
libs=''.join(f'<script>{(P/f).read_text()}</script>\n' for f in ['lib/xlsx-0.18.5/dist/xlsx.full.min.js','lib/jszip-3.10.1/dist/jszip.min.js'])
tpl='const TPL={'+','.join(f'{n}:"{base64.b64encode((P/"tpl"/(n+".xlsx")).read_bytes()).decode()}"' for n in ['elec','print','eval'])+'};'
app=(P/'app.js').read_text()+'\n'+(P/'ext.js').read_text()
out=page.replace('<!--LIBS-->',libs).replace('<!--APP-->',f'<script>{tpl}\n{app}</script>')
(P/'시간표_자동편성.html').write_text(out)
