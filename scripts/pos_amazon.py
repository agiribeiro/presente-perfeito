# -*- coding: utf-8 -*-
"""
Limpeza pós-descoberta na Amazon (offline).
    python3 scripts/pos_amazon.py catalogo_presentes.xlsx

1. Casamentos errados: título da Amazon de acessório (sleeve, guia, dashboard, organizador, expansão…)
   ou preço absurdo para o tipo → desliga asin/preço/foto/link, ativo=não, confianca="errado".
   A sugestão fica em asin_sugerido para você revisar.
2. Itens sem perfil (vindos de --buscar): etiqueta por linha LEGO / termo de busca / palavras do
   título, preenche notas 0-3 e gera rascunho de por_que a partir do título.
3. Nome dos itens de busca: encurta títulos de loja (corta depois de " | ", " - " e vírgulas
   de especificação) para o cartão do quiz; o título completo fica em titulo_amazon.
"""
import sys, re, datetime as dt
from openpyxl import load_workbook

ACESSORIO = re.compile(r"sleeve|guia de estrat|strategy guide|unofficial|dashboard|organizador|insert|playmat|tapete de jogo|"
                       r"capa protetora|protetor|expans[aã]o|suplemento|sticker|adesivo|token|miniatura avulsa|livro de regras|"
                       r"caderno|poster|p[oô]ster|cart[aã]o presente|gift card|vale", re.I)
HOJE = dt.date.today().isoformat()

# (padrão no termo de busca ou título) → perfil, notas
REGRAS = [
    (r"technic",                       ["construção", "ciência", "motor"],           dict(div=2, edu=3, cri=1)),
    (r"duplo",                         ["construção", "motor", "faz de conta"],      dict(div=3, edu=2, cri=2)),
    (r"classic|creator|peças criativas|blocos de montar|jogo de montar|blocos magn", ["construção", "criativo"], dict(div=2, edu=2, cri=3)),
    (r"city|friends",                  ["construção", "faz de conta", "personagens"], dict(div=3, edu=1, cri=2)),
    (r"mem[oó]ria|alfabeto|letras|sílabas|silabas|soletr|palavra|matem[aá]tica|n[uú]meros|bandeiras|capitais|montessori|encaixe|educativo",
                                       ["curioso", "motor", "ciência"],              dict(div=2, edu=3, cri=1)),
    (r"star wars|ninjago|harry potter|minecraft|marvel|disney|super mario|sonic|patrulha canina|pok[eé]mon|barbie|frozen",
                                       ["personagens", "construção", "faz de conta"], dict(div=3, edu=1, cri=2)),
    (r"microsc[oó]pio|telesc[oó]pio|ci[eê]ncia|experi[eê]ncia|qu[ií]mica|laborat[oó]rio|robótica|robotica|circuito",
                                       ["ciência", "curioso", "criativo"],           dict(div=2, edu=3, cri=2)),
    (r"quebra.?cabe[cç]a|puzzle",      ["curioso", "arte"],                          dict(div=2, edu=3, cri=0)),
    (r"massinha|modelar|play.?doh|pintura|desenho|tinta|giz|slime|areia",
                                       ["criativo", "arte", "motor"],                dict(div=3, edu=1, cri=3)),
    (r"hot wheels|pista|carrinho|autorama|trem",
                                       ["motor", "esportes", "personagens"],         dict(div=3, edu=1, cri=1)),
    (r"boneca|boneco|pelúcia|pelucia|interativ",
                                       ["faz de conta", "personagens"],              dict(div=3, edu=1, cri=2)),
    (r"bola|patinete|skate|bicicleta|esporte|dardos|boliche",
                                       ["esportes", "motor", "brincar em grupo"],    dict(div=3, edu=0, cri=0)),
    (r"cozinha|m[eé]dico|mercado|casinha|profiss",
                                       ["faz de conta", "criativo", "brincar em grupo"], dict(div=3, edu=1, cri=3)),
]
PADRAO = {"lego": (["construção", "criativo"], dict(div=2, edu=2, cri=2)),
          "brinquedo": (["motor", "curioso"], dict(div=2, edu=2, cri=1)),
          "jogo": (["brincar em grupo", "competitivo"], dict(div=2, edu=1, cri=0))}


def notas_preco(preco):
    p = float(preco or 0)
    imp = 3 if p > 200 else (2 if p >= 120 else 1)
    cb = 3 if 0 < p < 80 else (2 if p <= 150 else (1 if p <= 250 else 0))
    return imp, cb


def etiquetar(tipo, termo, titulo):
    txt = f"{termo} {titulo}".lower()
    for pat, perfil, n in REGRAS:
        if re.search(pat, txt):
            return perfil, n
    return PADRAO.get(tipo, PADRAO["brinquedo"])


def nome_curto(t):
    partes = [x.strip() for x in re.split(r"\s*\|\s*", str(t)) if x.strip()]
    t = partes[0] if len(partes) == 1 or len(partes[0]) >= 20 else partes[1]   # pula marca curta ("Pais & Filhos | …")
    t = re.split(r"\s+[-–]\s+|,\s", t)[0]
    t = re.sub(r"\(\d+ pe[cç]as\)|\d+\s*pe[cç]as", "", t, flags=re.I).strip(" ,-")
    return t[:70]


def rascunho(tipo, titulo, idade):
    t = str(titulo or "")
    pecas = re.search(r"(\d{2,4})\s*pe[cç]as", t, re.I)
    base = {"lego": "Conjunto LEGO", "brinquedo": "Brinquedo", "jogo": "Jogo"}[tipo]
    detalhe = (f" com {pecas.group(1)} peças" if pecas else "")
    return f"{base}{detalhe}, indicado a partir de {idade} anos. (rascunho automático — reescrever)" if idade else f"{base}{detalhe}. (rascunho automático — reescrever)"


def main():
    arq = sys.argv[1] if len(sys.argv) > 1 else "catalogo_presentes.xlsx"
    wb = load_workbook(arq); ws = wb["Catálogo"]
    h = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
    g = lambda r, c: ws.cell(r, h[c]).value if c in h else None
    s = lambda r, c, v: ws.cell(r, h[c]).__setattr__("value", v) if c in h else None
    n_err = n_tag = n_nome = 0
    for r in range(2, ws.max_row + 1):
        if not g(r, "nome"): continue
        tipo = str(g(r, "tipo") or "jogo").lower()
        titulo = str(g(r, "titulo_amazon") or "")
        preco = float(g(r, "preco") or 0)
        origem = str(g(r, "origem") or "")
        de_busca = origem.startswith("busca amazon")
        # 1) acessórios e preços absurdos
        errado = bool(g(r, "asin")) and not de_busca and (ACESSORIO.search(titulo) or (tipo == "jogo" and 0 < preco < 35))
        if errado:
            for c in ("asin", "preco", "preco_lista", "url_imagem", "url_afiliado"): s(r, c, None)
            s(r, "disponivel", None); s(r, "ativo", "não"); s(r, "confianca", "errado")
            s(r, "observacoes", f"⚠ casamento errado na Amazon ({titulo[:60]}) · revisar asin_sugerido · " + str(g(r, "observacoes") or ""))
            n_err += 1
        # 2) etiquetas para quem não tem
        if not g(r, "perfil"):
            termo = origem.replace("busca amazon:", "").split("·")[0]
            perfil, n = etiquetar(tipo, termo, g(r, "nome") or titulo)
            imp, cb = notas_preco(preco)
            s(r, "perfil", ", ".join(perfil))
            for k, v in (("diversao", n["div"]), ("educativo", n["edu"]), ("criatividade", n["cri"]), ("impressionar", imp), ("custo_beneficio", cb)):
                if g(r, k) is None: s(r, k, v)
            pq = str(g(r, "por_que") or "")
            if not pq or pq.strip() == "(rascunho automático)":
                s(r, "por_que", rascunho(tipo, g(r, "nome"), g(r, "idade_min")))
            s(r, "observacoes", "tags por regra (revisar) · " + str(g(r, "observacoes") or ""))
            n_tag += 1
        # 3) nome curto para itens de busca
        if de_busca and len(str(g(r, "nome"))) > 45:
            if not g(r, "titulo_amazon"): s(r, "titulo_amazon", g(r, "nome"))
            s(r, "nome", nome_curto(g(r, "nome"))); n_nome += 1
    wb.save(arq)
    print(f"casamentos errados desligados: {n_err} · itens etiquetados: {n_tag} · nomes encurtados: {n_nome}")


if __name__ == "__main__":
    main()
