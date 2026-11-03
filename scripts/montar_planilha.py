# -*- coding: utf-8 -*-
"""Gera catalogo_presentes.xlsx a partir de semente_jogos.py."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

from semente_jogos import JOGOS

HOJE = dt.date.today().isoformat()
FONTE = "Arial"

# ---------------------------------------------------------------- colunas
BLOCO1 = [  # curadoria (você preenche)
    ("asin", 14), ("nome", 30), ("tipo", 10), ("idade_min", 9),
    ("idade_ideal_min", 9), ("idade_ideal_max", 9), ("perfil", 34),
    ("diversao", 9), ("educativo", 9), ("criatividade", 9), ("impressionar", 9),
    ("custo_beneficio", 9), ("ocasiao", 18), ("por_que", 70), ("ativo", 7),
]
BLOCO2 = [  # Amazon Creators API (script diário sobrescreve)
    ("preco", 10), ("preco_lista", 10), ("disponivel", 10), ("url_imagem", 18),
    ("url_afiliado", 18), ("atualizado_em", 12),
]
BLOCO3 = [  # referência BGG/Ludopedia (uso interno)
    ("nome_en", 28), ("bgg_id", 9), ("ludopedia_id", 11), ("jogadores_min", 9),
    ("jogadores_max", 9), ("duracao_min", 9), ("complexidade", 10), ("mecanicas", 40),
    ("categorias", 30), ("nota_bgg", 9), ("nota_ludopedia", 9), ("editora", 20),
]
BLOCO4 = [  # controle
    ("origem", 20), ("aprovado_em", 12), ("preco_ref_estimado", 10), ("observacoes", 40),
]
COLS = BLOCO1 + BLOCO2 + BLOCO3 + BLOCO4
HEADERS = [c for c, _ in COLS]

FILL1 = PatternFill("solid", fgColor="FFF2CC")  # amarelo: você edita
FILL2 = PatternFill("solid", fgColor="E7E6E6")  # cinza: automático
FILL3 = PatternFill("solid", fgColor="DDEBF7")  # azul: referência
FILL4 = PatternFill("solid", fgColor="E2EFDA")  # verde: controle
FILL_BY_COL = {}
for c, _ in BLOCO1: FILL_BY_COL[c] = FILL1
for c, _ in BLOCO2: FILL_BY_COL[c] = FILL2
for c, _ in BLOCO3: FILL_BY_COL[c] = FILL3
for c, _ in BLOCO4: FILL_BY_COL[c] = FILL4

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def cabecalho(ws, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(name=FONTE, bold=True, size=10)
        cell.fill = FILL_BY_COL.get(h, PatternFill("solid", fgColor="D9D9D9"))
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"


def linha_jogo(j):
    """Converte um dict da semente numa linha na ordem de HEADERS."""
    d = {h: None for h in HEADERS}
    d.update(
        asin="", nome=j["nome"], tipo="jogo", idade_min=j["idade_min"],
        idade_ideal_min=j["ideal"][0], idade_ideal_max=j["ideal"][1],
        perfil=", ".join(j["perfil"]),
        diversao=j["prio"]["div"], educativo=j["prio"]["edu"],
        criatividade=j["prio"]["cri"], impressionar=j["prio"]["imp"],
        custo_beneficio=j["prio"]["cb"], ocasiao="", por_que=j["por_que"], ativo="não",
        nome_en=j["nome_en"], bgg_id=j["bgg_id"], jogadores_min=j["jog"][0],
        jogadores_max=j["jog"][1], duracao_min=j["dur"], complexidade=j["peso"],
        mecanicas=j["mecanicas"], editora=j["editora"],
        origem="semente (rascunho)", preco_ref_estimado=j["preco"],
        observacoes="Números de referência são estimativas; rodar bgg_verificar.py. "
                    "ASIN e ativo=sim após conferir na Amazon.",
    )
    return [d[h] for h in HEADERS]


def aplicar_validacoes(ws, n_linhas):
    rng = lambda col: f"{col}2:{col}{n_linhas + 200}"
    col = {h: get_column_letter(i) for i, h in enumerate(HEADERS, start=1)}
    dv_tipo = DataValidation(type="list", formula1='"jogo,brinquedo,lego"', allow_blank=True)
    dv_sn = DataValidation(type="list", formula1='"sim,não"', allow_blank=True)
    dv_0_3 = DataValidation(type="whole", operator="between", formula1="0", formula2="3", allow_blank=True)
    dv_tipo.add(rng(col["tipo"]))
    dv_sn.add(rng(col["ativo"]))
    for h in ("diversao", "educativo", "criatividade", "impressionar", "custo_beneficio"):
        dv_0_3.add(rng(col[h]))
    for dv in (dv_tipo, dv_sn, dv_0_3):
        ws.add_data_validation(dv)


def estilizar_corpo(ws, n_linhas):
    for r in range(2, n_linhas + 1):
        for ci, h in enumerate(HEADERS, start=1):
            c = ws.cell(row=r, column=ci)
            c.font = Font(name=FONTE, size=10)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=h in ("por_que", "perfil", "mecanicas", "observacoes"))


# ---------------------------------------------------------------- workbook
wb = Workbook()

# LEIA-ME
ws = wb.active
ws.title = "LEIA-ME"
leia = [
    ("Catálogo de presentes — como usar", True),
    ("", False),
    ("Abas", True),
    ("Catálogo: o que está no ar. Só linhas com ativo = sim entram no site.", False),
    ("Candidatos: sugestões trazidas pelo script semanal. Marque aprovar = sim/não; os aprovados migram para o Catálogo na próxima rodada.", False),
    ("Rejeitados: o que você recusou, para o script não oferecer de novo.", False),
    ("Taxonomia: definição de cada tag de perfil e das notas de prioridade, e as regras que sugerem tags automaticamente.", False),
    ("", False),
    ("Cores das colunas", True),
    ("Amarelo = curadoria: você preenche e é o que o site mostra (perfil, notas 0–3, por_que).", False),
    ("Cinza = Amazon: preenchido todo dia pelo script a partir da Creators API. Não edite.", False),
    ("Azul = referência BGG/Ludopedia: uso interno para decidir e etiquetar. Não publicar nota nem ranking no site.", False),
    ("Verde = controle.", False),
    ("", False),
    ("Estado desta carga", True),
    (f"Semente de {len(JOGOS)} jogos gerada em {HOJE}. Todas as linhas estão com ativo = não e origem = 'semente (rascunho)'.", False),
    ("Jogadores, duração e complexidade são estimativas: rode bgg_verificar.py (na sua máquina) para preencher pela API do BGG.", False),
    ("bgg_id vazio = id não conhecido com segurança; o script busca pelo nome_en e pede confirmação.", False),
    ("preco_ref_estimado é só uma faixa de referência (R$) usada para a nota inicial de custo_beneficio; o preço real vem da Amazon.", False),
    ("Para publicar um jogo: preencha asin, revise perfil / notas / por_que, e mude ativo para sim.", False),
    ("", False),
    ("Notas de prioridade (0–3)", True),
    ("0 = não se aplica · 1 = um pouco · 2 = bom nisso · 3 = é o ponto forte. Use 3 com parcimônia: é o que desempata o ranking.", False),
]
for r, (txt, bold) in enumerate(leia, start=1):
    c = ws.cell(row=r, column=1, value=txt)
    c.font = Font(name=FONTE, bold=bold, size=12 if r == 1 else 10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 120

# Catálogo
ws = wb.create_sheet("Catálogo")
cabecalho(ws, HEADERS, [w for _, w in COLS])
for j in JOGOS:
    ws.append(linha_jogo(j))
n = ws.max_row
estilizar_corpo(ws, n)
aplicar_validacoes(ws, n)
ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{n}"
ws["A1"].comment = Comment("ASIN do produto na Amazon.com.br (10 caracteres). Chave que liga à Creators API.", "Claude")
ws[f"{get_column_letter(HEADERS.index('perfil')+1)}1"].comment = Comment(
    "Tags separadas por vírgula, 2 a 4 por jogo. Lista válida na aba Taxonomia.", "Claude")

# Candidatos
ws = wb.create_sheet("Candidatos")
cand_cols = [("aprovar", 9), ("data_sugestao", 12)] + COLS
cabecalho(ws, [c for c, _ in cand_cols], [w for _, w in cand_cols])
ws["A1"].fill = FILL1
ws["B1"].fill = FILL4
# linha de exemplo
ex = ["", HOJE] + linha_jogo(dict(
    nome="(exemplo) Nome como está na Amazon", nome_en="Example", bgg_id=None, idade_min=6,
    ideal=(6, 9), jog=(2, 4), dur=20, peso=1.2, mecanicas="dice rolling", editora="",
    preco=120, perfil=["curioso"], prio=dict(div=2, edu=2, cri=0, imp=1, cb=2),
    por_que="Preenchido pelo script com a sugestão automática; você corrige e aprova."))
ex[cand_cols.index(("origem", 20))] = "sugestão semanal"
ws.append(ex)
for ci in range(1, len(cand_cols) + 1):
    c = ws.cell(row=2, column=ci)
    c.font = Font(name=FONTE, size=10, italic=True, color="808080")
    c.border = BORDER
dv = DataValidation(type="list", formula1='"sim,não"', allow_blank=True)
dv.add("A2:A500")
ws.add_data_validation(dv)

# Rejeitados
ws = wb.create_sheet("Rejeitados")
rej = [("asin", 14), ("nome", 30), ("bgg_id", 9), ("motivo", 40), ("data", 12)]
cabecalho(ws, [c for c, _ in rej], [w for _, w in rej])
for ci in range(1, len(rej) + 1):
    ws.cell(row=1, column=ci).fill = FILL4

# Taxonomia
ws = wb.create_sheet("Taxonomia")
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 60
ws.column_dimensions["C"].width = 60
ws.append(["Tag de perfil (pergunta 3)", "O que significa", "Regra de sugestão automática (mecânica/categoria BGG → tag)"])
tax = [
    ("criativo", "Gosta de inventar, desenhar, contar histórias", "storytelling, drawing, word game → criativo"),
    ("competitivo", "Quer vencer; aceita bem disputa direta", "take that, player elimination, real-time, 2 jogadores abstratos → competitivo"),
    ("curioso", "Gosta de descobrir, deduzir, resolver", "deduction, memory, pattern recognition, puzzle, trivia → curioso"),
    ("construção", "Gosta de montar, encaixar, erguer", "tile placement, stacking, network/route building → construção"),
    ("ciência", "Interesse por números, lógica, natureza", "code breaking, grid coordinates, categoria Educational/Science → ciência"),
    ("arte", "Sensível a desenho, cor, estética", "categoria Art, jogos com ilustração como mecânica (Dixit, Azul) → arte"),
    ("esportes", "Energia, reflexo, velocidade", "real-time, dexterity, flicking → esportes"),
    ("estratégia", "Gosta de planejar e decidir", "complexidade (weight) >= 1.5 ou abstract strategy → estratégia"),
    ("personagens", "Se liga em bichos, heróis, monstros, temas", "categoria Animals, Fantasy, Movies/TV, Comic → personagens"),
    ("brincar em grupo", "Melhor com várias pessoas", "cooperative game ou jogadores_max >= 6 → brincar em grupo"),
    ("motor", "(brinquedos) Movimento, coordenação, corpo", "stacking and balancing, dexterity → motor"),
    ("faz de conta", "(brinquedos) Imaginação, cenários, bonecos", "storytelling, role playing → faz de conta"),
]
for row in tax:
    ws.append(list(row))
ws.append([])
ws.append(["Nota de prioridade (pergunta 6)", "O que significa", "Regra de sugestão automática"])
pri = [
    ("diversao", "Risada, agitação, vontade de repetir", "weight < 1.5 e (dice rolling ou real-time ou take that) → 3; weight < 1.5 → 2; senão 1"),
    ("educativo", "Treina habilidade (lógica, contagem, vocabulário, espacial)", "deduction, puzzle, abstract strategy, trivia, word game → 3; set collection, tile placement → 2"),
    ("criatividade", "Pede imaginação ou expressão", "storytelling, drawing → 3; voting/communication → 2; senão 0"),
    ("impressionar", "Caixa, peças e mesa bonitas; 'uau' ao abrir", "componentes 3D/madeira ou preço > R$200 → 3; preço 120–200 → 2; senão 0–1"),
    ("custo_beneficio", "Valor pelo que custa", "preço < R$80 → 3; 80–150 → 2; 150–250 → 1; > 250 → 0 (recalculado pelo preço real da Amazon)"),
]
for row in pri:
    ws.append(list(row))
for r in range(1, ws.max_row + 1):
    for c in range(1, 4):
        cell = ws.cell(row=r, column=c)
        cell.font = Font(name=FONTE, size=10, bold=(r in (1, len(tax) + 3)))
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if r in (1, len(tax) + 3):
            cell.fill = PatternFill("solid", fgColor="D9D9D9")

wb.save("catalogo_presentes.xlsx")
print("ok", len(JOGOS), "jogos")
