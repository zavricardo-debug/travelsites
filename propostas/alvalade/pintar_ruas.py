"""Pinta a faixa de rodagem nas fotos reais de Street View (só o asfalto é alterado)."""
import numpy as np
from PIL import Image, ImageDraw, ImageFilter
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAIZ = BASE.parents[1]

# (ficheiro original, saída, cor RGB, polígono da faixa de rodagem em pixels)
RUAS = {
    "acacio": (RAIZ / "R.acaciodepaiva.png", BASE / "fotos/rua-acacio-paiva.png", (229, 57, 53),
               [(690, 447), (760, 447), (1000, 540), (1100, 600), (1482, 640), (1482, 907),
                (240, 907), (330, 740), (560, 560), (640, 455)]),
    "marquesa": (RAIZ / "RMarquesdeAlorna.png", BASE / "fotos/rua-marquesa-de-alorna.png", (30, 110, 220),
                 [(640, 470), (700, 470), (830, 520), (830, 600), (945, 625), (1100, 760),
                  (1250, 860), (1476, 906), (150, 906), (420, 640), (560, 500)]),
    "esaguy": (RAIZ / "RjosedEsaguy.png", BASE / "fotos/rua-jose-d-esaguy.png", (46, 184, 114),
               [(700, 420), (780, 420), (1000, 500), (1250, 560), (1476, 640), (1476, 905),
                (450, 905), (680, 600), (700, 470)]),
    "duro": (RAIZ / "RJoseDuro.png", BASE / "fotos/rua-jose-duro.png", (255, 122, 51),
             [(700, 540), (800, 540), (1150, 650), (1487, 760), (1487, 910), (300, 910),
              (560, 700), (640, 560)]),
}


# Zonas a NÃO pintar (carros estacionados, contentores, motas, peões)
EXCLUIR = {
    "acacio": [
        [(500, 520), (600, 500), (725, 520), (725, 625), (620, 625), (500, 600)],   # carro prateado
        [(815, 495), (915, 500), (925, 555), (815, 560)],                           # carro branco
        [(885, 515), (1000, 510), (1072, 575), (1072, 625), (950, 620), (885, 575)],  # carro escuro
        [(1005, 555), (1100, 560), (1482, 600), (1482, 835), (1300, 835), (1200, 760), (1005, 640)],  # carro azul
        [(315, 535), (500, 535), (505, 780), (330, 780)],                           # contentor azul
        [(490, 535), (590, 535), (590, 665), (490, 665)],                           # mota
        [(255, 680), (350, 680), (350, 805), (255, 805)],                           # caixa vermelha
        [(360, 660), (460, 660), (460, 830), (360, 830)],                           # contentor verde
        [(1255, 520), (1300, 520), (1300, 575), (1255, 575)],                       # peão
    ],
    "marquesa": [
        [(550, 510), (640, 500), (700, 520), (700, 575), (600, 575), (550, 545)],   # carro preto
        [(722, 505), (772, 505), (772, 545), (722, 545)],                           # carro azul
        [(775, 485), (835, 485), (835, 540), (775, 540)],                           # carros ao fundo
        [(815, 530), (945, 545), (945, 615), (815, 615)],                           # motas
        [(470, 515), (550, 515), (550, 612), (470, 612)],                           # contentor
    ],
    "esaguy": [
        [(905, 495), (1000, 500), (1200, 530), (1476, 600), (1476, 905), (900, 905)],  # SUV escuro
        [(855, 470), (960, 490), (960, 600), (855, 600)],                           # carro verde
        [(792, 440), (870, 450), (870, 500), (792, 500)],                           # carro azul
        [(760, 415), (805, 415), (805, 450), (760, 450)],                           # carros ao fundo
    ],
    "duro": [
        [(605, 565), (780, 565), (780, 640), (605, 640)],    # carro branco
        [(725, 540), (800, 540), (800, 600), (725, 600)],    # carros ao fundo
        [(925, 550), (1040, 550), (1040, 640), (925, 640)],  # carros à direita
        [(440, 555), (570, 555), (570, 720), (440, 720)],    # contentor
        [(415, 620), (462, 620), (462, 695), (415, 695)],    # caixa vermelha
        [(990, 600), (1150, 640), (1290, 800), (1000, 800)], # estacionamento de bicicletas
    ],
}


def pintar(origem, destino, cor, poligono, excluir=()):
    img = Image.open(origem).convert("RGB")
    arr = np.asarray(img).astype(np.float32)
    w, h = img.size
    m = Image.new("L", (w, h), 0)
    ImageDraw.Draw(m).polygon(poligono, fill=255)
    for zona in excluir:
        ImageDraw.Draw(m).polygon(zona, fill=0)
    m = m.filter(ImageFilter.GaussianBlur(1.5))
    mask = np.asarray(m).astype(np.float32) / 255.0

    r, g, b = arr[..., 0], arr[..., 1], arr[..., 2]
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    sat = arr.max(-1) - arr.min(-1)
    # asfalto: cinzento, pouco saturado, escuro/médio
    asfalto = ((sat < 28) & (lum > 25) & (lum < 150)).astype(np.float32)
    asfalto = np.asarray(Image.fromarray((asfalto * 255).astype(np.uint8)).filter(
        ImageFilter.MinFilter(3)).filter(ImageFilter.GaussianBlur(1.0))).astype(np.float32) / 255.0
    a = (mask * asfalto)[..., None] * 0.9

    t = np.clip(lum / 110.0, 0, 1.2)[..., None]  # textura do asfalto preservada
    colorido = np.array(cor, dtype=np.float32)[None, None, :] * (0.55 + 0.6 * t)
    out = arr * (1 - a) + colorido * a
    Image.fromarray(np.clip(out, 0, 255).astype(np.uint8)).save(destino)


if __name__ == "__main__":
    for nome, (src, dst, cor, pol) in RUAS.items():
        pintar(src, dst, cor, pol, EXCLUIR.get(nome, ()))
        print("ok", nome)
