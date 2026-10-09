from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

W, H = Inches(13.333), Inches(7.5)
INK = RGBColor(0x1F, 0x1F, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0x5A, 0x5A, 0x6B)
PINK = RGBColor(0xF4, 0x8F, 0xB6)
BLUE = RGBColor(0x2E, 0x86, 0xDE)
YELLOW = RGBColor(0xFF, 0xC8, 0x2E)
RED = RGBColor(0xE5, 0x39, 0x35)
GREEN = RGBColor(0x2E, 0xB8, 0x72)
ORANGE = RGBColor(0xFF, 0x7A, 0x33)
CREAM = RGBColor(0xFB, 0xF8, 0xF3)

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]


def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def text(slide, x, y, w, h, content, size=20, color=INK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = content if isinstance(content, list) else [content]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(8)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
    return tb


def bullets(slide, x, y, w, h, items, size=20, color=INK):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = "•  " + item
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = "Calibri"
    return tb


def header(slide, title, kicker=None):
    rect(slide, 0, 0, W, Inches(0.12), PINK)
    if kicker:
        text(slide, Inches(0.7), Inches(0.45), Inches(12), Inches(0.4), kicker.upper(),
             size=13, color=GREY, bold=True)
    text(slide, Inches(0.7), Inches(0.8), Inches(12), Inches(1.0), title, size=34, bold=True)


def footer(slide, n=None):
    n = len(prs.slides._sldIdLst)
    text(slide, Inches(0.7), Inches(7.0), Inches(9), Inches(0.3),
         "Proposta · Junta de Freguesia de Alvalade · Alvalade Colorida", size=10, color=GREY)
    text(slide, Inches(11.9), Inches(7.0), Inches(0.8), Inches(0.3), str(n), size=10,
         color=GREY, align=PP_ALIGN.RIGHT)


# 1. Capa
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, INK)
stripes = [PINK, BLUE, YELLOW, GREEN, ORANGE]
sw = W / len(stripes)
for i, c in enumerate(stripes):
    rect(s, int(sw * i), Inches(5.6), int(sw) + 1, Inches(0.35), c)
text(s, Inches(0.9), Inches(1.4), Inches(11.5), Inches(0.5), "PROPOSTA À JUNTA DE FREGUESIA DE ALVALADE",
     size=16, color=YELLOW, bold=True)
text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(2.2),
     "Alvalade Colorida", size=66, color=WHITE, bold=True)
text(s, Inches(0.9), Inches(3.5), Inches(11.5), Inches(1.4),
     "Dar cor e visibilidade às ruas Marquesa de Alorna, Acácio Paiva, José d'Esaguy e José Duro",
     size=24, color=WHITE)
text(s, Inches(0.9), Inches(6.2), Inches(11.5), Inches(0.5),
     "Comércio local · Montras abertas · Ruas que se veem da Av. da Igreja", size=16, color=PINK)

# 2. Problema
s = prs.slides.add_slide(BLANK)
header(s, "Quem passa na Av. da Igreja não sabe o que existe ali ao lado", "O problema")
bullets(s, Inches(0.7), Inches(2.0), Inches(6.8), Inches(4.5), [
    "As ruas transversais à Av. da Igreja têm comércio, mas são invisíveis para quem circula.",
    "Não há nada que convide a entrar: fachadas discretas, sem identidade comum.",
    "O comércio de bairro depende de quem já conhece a loja, não de quem passa.",
    "Resultado: lojas com pouca procura e uma rua que parece ‘de passagem’.",
], size=21)
rect(s, Inches(8.0), Inches(2.0), Inches(4.6), Inches(4.4), CREAM)
text(s, Inches(8.3), Inches(2.3), Inches(4.0), Inches(4.0),
     ["Objetivo", "Transformar quatro ruas num percurso reconhecível, com cor, montras e lojas abertas, "
      "que se note a partir da Av. da Igreja."], size=20, color=INK)
footer(s, 2)

# 3. Referência
s = prs.slides.add_slide(BLANK)
header(s, "Uma inspiração que já funciona: a Rua Cor de Rosa", "Referência")
rect(s, Inches(0.7), Inches(2.0), Inches(5.6), Inches(4.6), PINK)
text(s, Inches(1.0), Inches(2.4), Inches(5.0), Inches(3.8),
     ["Rua Nova do Carvalho, Cais do Sodré",
      "Pintada de rosa, tornou-se um dos pontos mais fotografados de Lisboa e atrai visitantes e "
      "esplanadas para uma rua antes pouco procurada."], size=22, color=INK)
bullets(s, Inches(6.8), Inches(2.1), Inches(5.9), Inches(4.5), [
    "A cor é o gancho: quem passa pergunta ‘o que é aquilo?’ e entra.",
    "Cada rua fica com uma identidade própria, sem perder ligação às outras.",
    "A cor funciona como sinal de orientação e de destino a visitar.",
    "A adaptação para Alvalade: quatro cores, uma por rua, num único percurso.",
], size=20)
footer(s, 3)

# 4. Conceito
s = prs.slides.add_slide(BLANK)
header(s, "Conceito: uma cor por rua, um percurso só", "A proposta")
ruas = [
    ("Marquesa de Alorna", BLUE, "Azul", "Moda e acessórios"),
    ("Acácio Paiva", YELLOW, "Amarelo", "Cafés, pastelarias e restauração"),
    ("José d'Esaguy", GREEN, "Verde", "Serviços e lojas de bairro"),
    ("José Duro", ORANGE, "Laranja", "Comércio criativo e artesanato"),
]
cw = Inches(2.95)
gap = Inches(0.2)
x0 = Inches(0.7)
for i, (nome, cor, nome_cor, foco) in enumerate(ruas):
    x = x0 + i * (cw + gap)
    rect(s, x, Inches(2.0), cw, Inches(1.6), cor)
    text(s, x + Inches(0.2), Inches(2.15), cw - Inches(0.4), Inches(1.3), [nome_cor, nome],
         size=20, color=WHITE if cor != YELLOW else INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, Inches(3.6), cw, Inches(2.7), CREAM)
    text(s, x + Inches(0.2), Inches(3.8), cw - Inches(0.4), Inches(2.4),
         ["Foco sugerido", foco,
          "", "Montra colorida com as lojas abertas da rua"], size=17, color=INK)
text(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.5),
     "As cores e os focos são sugestões para discutir com os comerciantes.", size=14, color=GREY)
footer(s, 4)

# ---- Diapositivos por rua: estrada pintada + montra no início da rua ----
def street_slide(nome, cor, cor_txt, img, pintura, montra_intro, lojas):
    s = prs.slides.add_slide(BLANK)
    header(s, nome, "Rua a pintar · montra à entrada")
    # Mockup da rua com a estrada pintada
    import os
    if os.path.exists(img):
        pic = s.shapes.add_picture(img, Inches(0.6), Inches(1.95), width=Inches(6.9))
        h = pic.height
    else:
        h = Inches(4.6)
        rect(s, Inches(0.6), Inches(1.95), Inches(6.9), h, CREAM)
        text(s, Inches(0.9), Inches(3.6), Inches(6.3), Inches(1.0),
             "[Foto real da rua: estrada pintada em " + cor_txt.lower() + "]", size=18,
             color=GREY, align=PP_ALIGN.CENTER)
    rect(s, Inches(0.6), Inches(1.95) + h + Inches(0.1), Inches(0.35), Inches(0.35), cor)
    text(s, Inches(1.05), Inches(1.95) + h + Inches(0.05), Inches(6.8), Inches(0.5),
         "Estrada pintada em " + cor_txt.lower() + " e montra à entrada da rua", size=12, color=GREY)
    # Painel de texto
    text(s, Inches(8.2), Inches(1.9), Inches(4.6), Inches(1.5), pintura, size=14, color=INK)
    # Montra estilo 'diretório de centro comercial'
    bx, by, bw, bh = Inches(8.2), Inches(3.5), Inches(4.6), Inches(3.3)
    rect(s, bx, by, bw, bh, INK)
    if cor_txt == "Arco-íris":
        faixas = [RGBColor(0xE5,0x39,0x35), ORANGE, YELLOW, GREEN, BLUE, RGBColor(0x8E,0x44,0xAD)]
        fw_ = int(bw / len(faixas))
        for i, c in enumerate(faixas):
            rect(s, bx + fw_ * i, by, fw_ + 1, Inches(0.55), c)
    else:
        rect(s, bx, by, bw, Inches(0.55), cor)
    text(s, bx + Inches(0.15), by + Inches(0.02), bw - Inches(0.3), Inches(0.5),
         "MAPA DA RUA · " + nome.upper(), size=12, color=cor_txt_for(cor), bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, bx + Inches(0.15), by + Inches(0.6), bw - Inches(0.3), Inches(0.4),
         montra_intro, size=11, color=PINK)
    yy = by + Inches(0.95)
    for categoria, loja in lojas:
        text(s, bx + Inches(0.15), yy, Inches(1.6), Inches(0.35), categoria.upper(), size=10,
             color=YELLOW, bold=True)
        text(s, bx + Inches(1.7), yy, bw - Inches(1.85), Inches(0.35), loja, size=13, color=WHITE)
        yy += Inches(0.36)
    text(s, bx + Inches(0.15), by + bh - Inches(0.35), bw - Inches(0.3), Inches(0.3),
         "▼ Você está na entrada da rua (Av. da Igreja)", size=10, color=PINK)
    footer(s)
    return s


def cor_txt_for(cor):
    return INK if cor in (YELLOW, PINK) else WHITE


# ---- Antes: fotografia original + comércio que não se vê da Av. da Igreja ----
def antes_slide(nome, img, comercios, nota):
    s = prs.slides.add_slide(BLANK)
    header(s, nome, "Hoje · o que não se vê da Av. da Igreja")
    pic = s.shapes.add_picture(img, Inches(0.6), Inches(1.95), width=Inches(6.6))
    text(s, Inches(0.6), Inches(1.95) + pic.height + Inches(0.08), Inches(6.6), Inches(0.4),
         "Fotografia atual (Google Street View, 2026)", size=12, color=GREY)
    text(s, Inches(7.6), Inches(1.9), Inches(5.2), Inches(0.5),
         "Comércio existente, hoje sem visibilidade", size=16, bold=True)
    bullets(s, Inches(7.6), Inches(2.45), Inches(5.2), Inches(3.4), comercios, size=15)
    rect(s, Inches(7.6), Inches(5.75), Inches(5.2), Inches(1.0), CREAM)
    text(s, Inches(7.8), Inches(5.8), Inches(4.9), Inches(0.9), nota, size=13, color=INK)
    footer(s)


antes_slide(
    "Rua Acácio Paiva", "fotos/original-acacio-paiva.png",
    ["Millennium bcp: agência bancária",
     "Cobaia (n.º 19): restaurante mediterrânico, bar e música",
     "Oakberry Açaí (n.º 3F)",
     "Tasco Force (n.º 5D)",
     "The Coffee Alvalade (n.º 14C): em abertura, a confirmar"],
    "Quem circula na Av. da Igreja não sabe que a rua tem restauração e serviços a poucos metros da avenida.")

antes_slide(
    "Rua Marquesa de Alorna", "fotos/original-marquesa-de-alorna.png",
    ["Pasta Non Basta (n.º 17B): cozinha italiana",
     "A Triunfante de Alvalade (n.º 18)",
     "O Declive (n.º 22D)",
     "Petisco de Alvalade (n.º 25)",
     "Tasca O Cantinho dos Sabores (n.º 30A)",
     "Mickael Mezdari, pâtisserie francesa (n.º 27C)"],
    "A rua é curta e quase sem sinalização; quem passa não percebe que há comércio a poucos metros.")

antes_slide(
    "Rua José d'Esaguy", "fotos/original-jose-d-esaguy.png",
    ["Bankinter: agência bancária",
     "Alberto Oculista: ótica",
     "Isco Pão e Vinho (n.º 10D): padaria",
     "Yokohama (n.º 3B): restaurante",
     "Pérola do Ceira (n.º 4E): cozinha portuguesa",
     "100 Montaditos (n.º 5)"],
    "As montras existem, mas não são vistas. A rua parece só de passagem, com carros a ocupar o espaço visual.")

antes_slide(
    "Rua José Duro", "fotos/original-jose-duro.png",
    ["BPI: agência bancária",
     "O Luís (n.º 29): cozinha portuguesa",
     "MADPIZZA (n.º 25): pizzaria",
     "Do Beco Alvalade (n.º 31A): padaria e brunch",
     "Restaurante Courenses (n.º 25C-D)",
     "Taste Invaders (n.º 22)"],
    "A padaria e os restaurantes têm procura de vizinhança, mas nenhum sinal que os anuncie a quem passa.")

street_slide(
    "Rua Marquesa de Alorna", BLUE, "Azul", "fotos/rua-marquesa-de-alorna.png",
    ["Pintura de toda a faixa de rodagem em azul, uma cor do arco-íris por rua.",
     "A montra fica no início da rua, visível a quem entra pela Av. da Igreja."],
    "Lojas abertas agora, como num centro comercial:",
    [("Restauração", "Pasta Non Basta  (n.º 17B)"),
     ("Restauração", "A Triunfante de Alvalade  (n.º 18)"),
     ("Restauração", "Petisco de Alvalade  (n.º 25)"),
     ("Café", "Tasca O Cantinho dos Sabores  (n.º 30A)"),
     ("Pastelaria", "Mickael Mezdari  (n.º 27C)")])

street_slide(
    "Rua Acácio Paiva", RED, "Vermelho", "fotos/rua-acacio-paiva.png",
    ["Pintura de toda a faixa de rodagem em vermelho, a cor mais visível de dia.",
     "Montra à entrada da rua com as lojas abertas e os horários do dia."],
    "Lojas abertas agora, como num centro comercial:",
    [("Banco", "Millennium bcp"),
     ("Restauração", "Cobaia  (n.º 19)"),
     ("Açaí", "Oakberry Açaí  (n.º 3F)"),
     ("Restauração", "Tasco Force  (n.º 5D)"),
     ("Café", "The Coffee Alvalade  (n.º 14C)")])

street_slide(
    "Rua José d'Esaguy", GREEN, "Verde", "fotos/rua-jose-d-esaguy.png",
    ["Pintura de toda a faixa de rodagem em verde.",
     "A montra fica à entrada, junto à Av. da Igreja, com as lojas e os serviços da rua."],
    "Lojas e serviços abertos agora:",
    [("Banco", "Bankinter"),
     ("Ótica", "Alberto Oculista"),
     ("Padaria", "Isco Pão e Vinho  (n.º 10D)"),
     ("Restauração", "Yokohama  (n.º 3B)"),
     ("Restauração", "Pérola do Ceira  (n.º 4E)")])

street_slide(
    "Rua José Duro", ORANGE, "Laranja", "fotos/rua-jose-duro.png",
    ["Pintura de toda a faixa de rodagem em laranja, a cor mais quente do percurso.",
     "Montra à entrada da rua, com as pastelarias, restaurantes e lojas abertas."],
    "Lojas abertas agora, como num centro comercial:",
    [("Banco", "BPI"),
     ("Restauração", "O Luís  (n.º 29)"),
     ("Pizzaria", "MADPIZZA  (n.º 25)"),
     ("Padaria", "Do Beco Alvalade  (n.º 31A)"),
     ("Restauração", "Taste Invaders  (n.º 22)")])


# ---- Lista completa de comércios por rua ----
s = prs.slides.add_slide(BLANK)
header(s, "Comércio existente em cada rua (lista completa)", "Lista a validar no local")
colunas = [
    ("Acácio Paiva", RED, ["Millennium bcp (banco)", "Cobaia (n.º 19)", "Oakberry Açaí (n.º 3F)",
                           "Tasco Force (n.º 5D)", "The Coffee Alvalade (n.º 14C, em abertura)"]),
    ("Marquesa de Alorna", BLUE, ["Pasta Non Basta (n.º 17B)", "A Triunfante de Alvalade (n.º 18)",
                                  "O Declive (n.º 22D)", "Petisco de Alvalade (n.º 25)",
                                  "Tasca O Cantinho dos Sabores (n.º 30A)", "Truta e Meia (n.º 30)",
                                  "Mickael Mezdari, pâtisserie (n.º 27C)"]),
    ("José d'Esaguy", GREEN, ["Bankinter (banco)", "Alberto Oculista (ótica)", "Isco Pão e Vinho (n.º 10D)",
                              "Yokohama (n.º 3B)", "Pérola do Ceira (n.º 4E)", "100 Montaditos (n.º 5)"]),
    ("José Duro", ORANGE, ["BPI (banco)", "O Luís (n.º 29)", "MADPIZZA (n.º 25)", "Do Beco Alvalade (n.º 31A)",
                           "Restaurante Courenses (n.º 25C-D)", "Taste Invaders (n.º 22)",
                           "Maria Granel (n.º 22B, a confirmar)", "Laboratório do Gelado (n.º 21B, a confirmar)",
                           "Snack-bar pastelaria Helsinquia (a confirmar)"]),
]
cw = Inches(2.98)
for i, (titulo, cor, itens) in enumerate(colunas):
    x = Inches(0.7) + i * (cw + Inches(0.12))
    rect(s, x, Inches(1.95), cw, Inches(0.6), cor)
    text(s, x + Inches(0.12), Inches(1.95), cw - Inches(0.24), Inches(0.6), titulo, size=15, bold=True,
         color=INK if cor in (YELLOW, PINK) else WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, Inches(2.55), cw, Inches(4.1), CREAM)
    tb = s.shapes.add_textbox(x + Inches(0.12), Inches(2.7), cw - Inches(0.24), Inches(3.9))
    tf = tb.text_frame
    tf.word_wrap = True
    for k, item in enumerate(itens):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        r = p.add_run()
        r.text = "•  " + item
        r.font.size = Pt(12)
        r.font.color.rgb = INK
        r.font.name = "Calibri"
text(s, Inches(0.7), Inches(6.7), Inches(12), Inches(0.3),
     "Fontes: Junta de Freguesia de Alvalade, OpenAlfa, TripAdvisor, Restaurant Guru e sites dos estabelecimentos.",
     size=10, color=GREY)
footer(s)

# ---- Porque a cor ajuda ----
s = prs.slides.add_slide(BLANK)
header(s, "Porque é que a cor e as montras trazem visitantes", "O argumento")
cards = [
    ("Identificável", BLUE, "Uma cor por rua torna o percurso fácil de ver e de lembrar. 'A rua vermelha' é uma indicação que toda a gente entende."),
    ("Visível a partir da Av. da Igreja", PINK, "A cor e a montra à entrada funcionam como convite: quem passa vê que há algo ali, sem precisar de conhecer as lojas."),
    ("Autêntica", GREEN, "Comércio de bairro, com restauração, pastelarias e serviços, é uma experiência que as pessoas procuram. A cor mostra essa autenticidade."),
]
cw = Inches(3.95)
for i, (titulo, cor, desc) in enumerate(cards):
    x = Inches(0.7) + i * (cw + Inches(0.2))
    rect(s, x, Inches(2.0), cw, Inches(0.8), cor)
    text(s, x + Inches(0.2), Inches(2.0), cw - Inches(0.4), Inches(0.8), titulo, size=19, bold=True,
         color=INK if cor in (PINK, YELLOW) else WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, Inches(2.8), cw, Inches(2.9), CREAM)
    text(s, x + Inches(0.2), Inches(3.0), cw - Inches(0.4), Inches(2.6), desc, size=17, color=INK)
text(s, Inches(0.7), Inches(6.0), Inches(12), Inches(0.8),
     "Lista de comércios: Junta de Freguesia de Alvalade (2021) e guias online; confirmar no local. Referência: a Rua Nova do Carvalho (Cais do Sodré) mostra o mesmo efeito. Sugerimos medir antes e depois "
     "(contagens de peões, inquérito a clientes e vendas dos comerciantes aderentes) para validar os resultados.",
     size=14, color=GREY)
footer(s)

# 5. Montras
s = prs.slides.add_slide(BLANK)
header(s, "Uma montra por rua: as lojas abertas à vista de todos", "Montras")
bullets(s, Inches(0.7), Inches(2.0), Inches(6.5), Inches(4.8), [
    "Um painel ou montra comunitária em cada rua, na zona mais visível a partir da Av. da Igreja.",
    "Mostra os comércios abertos naquele momento, com foto, nome e horário.",
    "Atualização simples (quadro, ecrã ou placa com rotação de cartões).",
    "Mapa do percurso com as quatro ruas e a localização de cada loja.",
    "Pode ser um QR code que leva a uma página online com o mesmo conteúdo.",
], size=20)
rect(s, Inches(7.6), Inches(2.0), Inches(5.0), Inches(4.6), CREAM)
text(s, Inches(7.9), Inches(2.2), Inches(4.5), Inches(4.3),
     ["Exemplo de cartão de montra", "",
      "Nome da loja", "Rua José Duro, n.º X",
      "Aberto agora · até às 19h", "",
      "[foto da montra]"], size=18, color=INK)
footer(s, 5)

# 6. Como funciona para o comércio
s = prs.slides.add_slide(BLANK)
header(s, "O que se pede aos comerciantes", "Envolvimento")
bullets(s, Inches(0.7), Inches(2.0), Inches(6.0), Inches(4.5), [
    "Participação voluntária, com regras simples e comuns a todos.",
    "Manter a montra arrumada e a loja aberta nas horas anunciadas.",
    "Fornecer foto e horário para o painel da rua.",
    "Aderir à pintura da fachada ou de um elemento de cor (porta, toldo, vaso).",
], size=20)
rect(s, Inches(7.0), Inches(2.0), Inches(5.6), Inches(4.5), CREAM)
text(s, Inches(7.3), Inches(2.2), Inches(5.0), Inches(4.2),
     ["Ganho para o comércio", "",
      "Mais passagem de pessoas que hoje nem sabem que a loja existe.",
      "", "Campanha conjunta, com custos partilhados e divulgação comum."],
     size=20, color=INK)
footer(s, 6)

# 7. Implementação
s = prs.slides.add_slide(BLANK)
header(s, "Plano faseado", "Implementação")
fases = [
    ("1 · Estudo", "Levantamento de lojas, fachadas, pavimento e tráfego em cada rua."),
    ("2 · Adesão", "Reuniões com comerciantes e moradores; escolha de cores e montras."),
    ("3 · Licenças", "Aprovação técnica e licenciamento junto da CML e serviços competentes."),
    ("4 · Execução", "Pintura, montagem dos painéis e lançamento com evento de rua."),
    ("5 · Avaliação", "Contagens, inquéritos e ajuste ao fim de 6 a 12 meses."),
]
fw = Inches(2.4)
for i, (titulo, desc) in enumerate(fases):
    x = Inches(0.7) + i * (fw + Inches(0.12))
    rect(s, x, Inches(2.1), fw, Inches(0.8), [PINK, BLUE, YELLOW, GREEN, ORANGE][i])
    text(s, x + Inches(0.15), Inches(2.1), fw - Inches(0.3), Inches(0.8), titulo, size=18,
         bold=True, color=WHITE if i not in (2,) else INK, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(0.1), Inches(3.1), fw - Inches(0.2), Inches(2.5), desc, size=17, color=INK)
text(s, Inches(0.7), Inches(5.6), Inches(12), Inches(1.0),
     "Nota técnica: a pintura de pavimento deve ser antiderrapante, durável e não confundir a sinalização "
     "rodoviária. Pode começar por pintura nos passeios ou em troços de baixo tráfego, se a CML o aprovar.",
     size=16, color=GREY)
footer(s, 7)

# 8. Pedido e orçamento
s = prs.slides.add_slide(BLANK)
header(s, "O que pedimos à Junta", "Pedido")
bullets(s, Inches(0.7), Inches(2.0), Inches(7.0), Inches(4.8), [
    "Apoio institucional e articulação com a Câmara Municipal de Lisboa (licenças e pintura).",
    "Cedência ou apoio para a montagem das montras comunitárias.",
    "Divulgação pelos canais da Junta (site, redes sociais, boletim).",
    "Apoio financeiro parcial ou comparticipação, a definir após orçamento.",
    "Um ponto de contacto na Junta para acompanhar o projeto.",
], size=20)
rect(s, Inches(8.1), Inches(2.0), Inches(4.5), Inches(4.4), CREAM)
text(s, Inches(8.4), Inches(2.2), Inches(4.0), Inches(4.2),
     ["Orçamento", "",
      "[Pintura das 4 ruas: € a preencher]",
      "[4 montras comunitárias: € a preencher]",
      "[Comunicação e evento de lançamento: € a preencher]",
      "", "Valores a validar com orçamentos reais."], size=18, color=INK)
footer(s, 8)

# 9. Métricas
s = prs.slides.add_slide(BLANK)
header(s, "Como medimos o sucesso", "Resultados esperados")
metricas = [
    ("Visibilidade", "Nº de pessoas que tiram fotografias ou param nas ruas"),
    ("Comércio", "Nº de lojas aderentes e variação de clientes/vendas (inquérito)"),
    ("Passagem", "Contagens de peões na Av. da Igreja e nas entradas das ruas"),
    ("Comunidade", "Satisfação de moradores e comerciantes (inquérito semestral)"),
]
for i, (t, d) in enumerate(metricas):
    col, row = i % 2, i // 2
    x = Inches(0.7) + col * Inches(6.2)
    y = Inches(2.0) + row * Inches(2.3)
    rect(s, x, y, Inches(5.9), Inches(2.0), CREAM)
    rect(s, x, y, Inches(0.15), Inches(2.0), [BLUE, YELLOW, GREEN, ORANGE][i])
    text(s, x + Inches(0.4), y + Inches(0.2), Inches(5.3), Inches(0.6), t, size=24, bold=True)
    text(s, x + Inches(0.4), y + Inches(0.9), Inches(5.3), Inches(1.0), d, size=18, color=GREY)
footer(s, 9)

# 10. Próximos passos
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, INK)
text(s, Inches(0.9), Inches(1.2), Inches(11.5), Inches(0.5), "PRÓXIMOS PASSOS", size=16, color=YELLOW, bold=True)
text(s, Inches(0.9), Inches(1.8), Inches(11.5), Inches(1.2), "Vamos começar por uma rua?", size=44,
     color=WHITE, bold=True)
bullets(s, Inches(0.9), Inches(3.4), Inches(11.5), Inches(3.0), [
    "Reunião com a Junta para validar a proposta e o enquadramento.",
    "Contacto com os comerciantes das quatro ruas.",
    "Piloto numa rua (sugestão: José Duro, por ser a mais pequena) antes de avançar para as restantes.",
    "Contactos: [nome] · [email] · [telefone]",
], size=22, color=WHITE)
for i, c in enumerate(stripes):
    rect(s, int(sw * i), Inches(6.9), int(sw) + 1, Inches(0.6), c)

prs.save("/home/user/travelsites/propostas/alvalade/Proposta_Alvalade_Colorida.pptx")
print("ok")
