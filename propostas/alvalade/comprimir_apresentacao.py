import io, os, re, shutil, zipfile, glob
from PIL import Image

SRC = '/home/user/travelsites/propostas/alvalade/Proposta_Alvalade_Colorida.pptx'
WORK = '/tmp/build'
OUT = '/home/user/travelsites/propostas/alvalade/Proposta_Alvalade_Colorida_leve.pptx'
TARGET_TOTAL = 1_890_000          # alvo bruto; o zip final fica ~1,7 MB
MAX_W = 1400                      # largura maxima (slide 16:9 a ~96-110 dpi)

shutil.rmtree(WORK, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    z.extractall(WORK)

media = sorted(glob.glob(os.path.join(WORK, 'ppt/media/*.png')))
overhead = sum(os.path.getsize(os.path.join(WORK, n)) for n in names
               if not n.endswith('/') and not n.startswith('ppt/media/'))
budget = TARGET_TOTAL - int(overhead * 0.35) - 20_000   # xml comprime bem (~35%)
per_img = budget // len(media)
print(f'partes nao-media: {overhead/1024:.0f} KB brutos | orcamento por imagem: {per_img/1024:.0f} KB')

rename = {}
for f in media:
    im = Image.open(f)
    if im.mode in ('RGBA', 'P', 'LA'):
        bg = Image.new('RGB', im.size, (255, 255, 255))
        bg.paste(im.convert('RGBA'), mask=im.convert('RGBA').split()[-1])
        im = bg
    else:
        im = im.convert('RGB')
    if im.width > MAX_W:
        im = im.resize((MAX_W, round(im.height * MAX_W / im.width)), Image.LANCZOS)

    best = None
    lo, hi = 30, 92
    while lo <= hi:                      # procura binaria da melhor qualidade que cabe no orcamento
        q = (lo + hi) // 2
        buf = io.BytesIO()
        im.save(buf, 'JPEG', quality=q, optimize=True, progressive=True, subsampling=1)
        if buf.tell() <= per_img:
            best = (q, buf.getvalue()); lo = q + 1
        else:
            hi = q - 1
    if best is None:
        buf = io.BytesIO()
        im.save(buf, 'JPEG', quality=30, optimize=True, progressive=True, subsampling=2)
        best = (30, buf.getvalue())
    q, data = best
    new = f[:-4] + '.jpeg'
    open(new, 'wb').write(data)
    os.remove(f)
    rename[os.path.basename(f)] = os.path.basename(new)
    print(f'  {os.path.basename(f)} -> {os.path.basename(new)}  q={q}  '
          f'{im.size[0]}x{im.size[1]}  {len(data)/1024:.0f} KB')

# actualizar referencias nos .rels
for rels in glob.glob(os.path.join(WORK, 'ppt/**/_rels/*.rels'), recursive=True):
    s = open(rels, encoding='utf-8').read(); orig = s
    for old, new in rename.items():
        s = s.replace(old, new)
    if s != orig:
        open(rels, 'w', encoding='utf-8').write(s)

# [Content_Types].xml ja declara Default Extension="jpeg"; remover png se deixou de existir
ct = os.path.join(WORK, '[Content_Types].xml')
s = open(ct, encoding='utf-8').read()
if 'jpeg' not in s:
    s = s.replace('<Default Extension="png"',
                  '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png"')
if not glob.glob(os.path.join(WORK, '**/*.png'), recursive=True):
    s = re.sub(r'<Default Extension="png"[^/]*/>', '', s)
open(ct, 'w', encoding='utf-8').write(s)

# reescrever o zip preservando a ordem original das partes
order = [n for n in names if not n.endswith('/')]
order = [rename.get(os.path.basename(n), None) and os.path.dirname(n) + '/' + rename[os.path.basename(n)]
         or n for n in order]
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for n in order:
        p = os.path.join(WORK, n)
        if os.path.exists(p):
            z.write(p, n)
print(f'\n{OUT}: {os.path.getsize(OUT)/1024/1024:.2f} MB (original {os.path.getsize(SRC)/1024/1024:.2f} MB)')
