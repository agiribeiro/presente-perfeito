# -*- coding: utf-8 -*-
"""
Sincroniza o Catálogo com a Amazon.com.br pela Creators API.

Instalação (uma vez, com o .venv ativo):
    pip install python-amazon-paapi openpyxl

Credenciais (Associates Central → Ferramentas → Creators API → criar credencial; o segredo só
aparece uma vez). Defina no Terminal, nunca na planilha:
    export AMAZON_CRED_ID="..."
    export AMAZON_CRED_SECRET="..."
    export AMAZON_TAG="logicamentepresente-20"        # opcional; é o padrão

Modos:
    python3 amazon_sync.py catalogo_presentes.xlsx --descobrir
        Para cada linha do Catálogo SEM asin, busca o nome na Amazon e grava asin_sugerido,
        titulo_amazon, preco_sugerido e confianca (alta/média/baixa). Você confere.
    python3 amazon_sync.py catalogo_presentes.xlsx --aceitar
        Copia asin_sugerido → asin nas linhas com confianca = alta OU com aceitar = sim.
    python3 amazon_sync.py catalogo_presentes.xlsx --atualizar
        Para cada linha COM asin, grava preco, preco_lista, disponivel, url_imagem,
        url_afiliado e atualizado_em (colunas cinzas). É o que roda todo dia.
    python3 amazon_sync.py catalogo_presentes.xlsx --teste
        Só testa as credenciais com uma busca.

Depois de --atualizar, rode exportar_json.py para o quiz ler os preços novos.
"""
import os, sys, re, json, unicodedata, datetime as dt
from pathlib import Path
from openpyxl import load_workbook

try:
    from amazon_creatorsapi import AmazonCreatorsApi
except ImportError:
    sys.exit("Falta a biblioteca: pip install python-amazon-paapi")

CRED_ID = os.environ.get("AMAZON_CRED_ID", "").strip()
CRED_SECRET = os.environ.get("AMAZON_CRED_SECRET", "").strip()
TAG = os.environ.get("AMAZON_TAG", "logicamentepresente-20").strip()
VERSION = os.environ.get("AMAZON_API_VERSION") or ("3.1" if CRED_ID.startswith("amzn1.") else "2.1")
CACHE = Path("cache"); CACHE.mkdir(exist_ok=True)
HOJE = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
COLS_NOVAS = ["asin_sugerido", "titulo_amazon", "preco_sugerido", "confianca", "aceitar"]


def cliente():
    if not CRED_ID or not CRED_SECRET:
        sys.exit("Defina AMAZON_CRED_ID e AMAZON_CRED_SECRET no ambiente.")
    return AmazonCreatorsApi(credential_id=CRED_ID, credential_secret=CRED_SECRET, version=VERSION,
                             tag=TAG, country="BR", throttling=1.2)


def dig(obj, caminho, default=None):
    """Acesso tolerante: dig(item, 'offers_v2.listings.0.price.money.amount')."""
    cur = obj
    for p in caminho.split("."):
        if cur is None: return default
        if p.isdigit():
            try: cur = cur[int(p)]
            except (IndexError, TypeError, KeyError): return default
        else:
            cur = getattr(cur, p, None) if not isinstance(cur, dict) else cur.get(p)
    return default if cur is None else cur


def extrair(item):
    """Campos que a planilha usa, a partir de um Item da API."""
    preco = dig(item, "offers_v2.listings.0.price.money.amount")
    lista = dig(item, "offers_v2.listings.0.price.saving_basis.money.amount")
    disp = dig(item, "offers_v2.listings.0.availability.type") or dig(item, "offers_v2.listings.0.availability.message")
    img = dig(item, "images.primary.large.url") or dig(item, "images.primary.medium.url")
    return dict(
        asin=dig(item, "asin"), titulo=dig(item, "item_info.title.display_value"),
        preco=float(preco) if preco is not None else None,
        preco_lista=float(lista) if lista is not None else None,
        disponivel="sim" if (disp is None and preco is not None) or (disp and str(disp).upper().startswith(("IN_STOCK", "NOW", "AVAILABLE"))) else ("não" if preco is None else "sim"),
        url_imagem=img, url_afiliado=dig(item, "detail_page_url"),
    )


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return set(re.findall(r"[a-z0-9]{2,}", s)) - {"de", "do", "da", "the", "of", "jogo", "tabuleiro", "edicao", "edition", "board", "game"}


def confianca(nome, titulo):
    a, b = norm(nome), norm(titulo)
    if not a: return "baixa"
    cob = len(a & b) / len(a)
    return "alta" if cob >= 0.8 else ("média" if cob >= 0.5 else "baixa")


def garantir_colunas(ws):
    hdr = [c.value for c in ws[1]]
    for c in COLS_NOVAS:
        if c not in hdr:
            ws.cell(1, len(hdr) + 1).value = c; hdr.append(c)
    return {c: i + 1 for i, c in enumerate(hdr) if c}


# ------------------------------------------------------------ modos
def descobrir(ws, api):
    h = garantir_colunas(ws); n = 0
    for r in range(2, ws.max_row + 1):
        g = lambda c: ws.cell(r, h[c]).value
        if not g("nome") or g("asin") or g("asin_sugerido"): continue
        nome = str(g("nome")).strip()
        arq = CACHE / ("amz_" + re.sub(r"[^a-zA-Z0-9]+", "_", nome)[:60] + ".json")
        if arq.exists():
            cand = json.loads(arq.read_text(encoding="utf-8"))
        else:
            try:
                res = api.search_items(keywords=f"{nome} jogo de tabuleiro", item_count=3)
                cand = [extrair(it) for it in (res.items or [])]
            except Exception as e:
                print(f"  [{nome}] erro na busca: {str(e)[:120]}"); continue
            arq.write_text(json.dumps(cand, ensure_ascii=False), encoding="utf-8")
        if not cand:
            ws.cell(r, h["confianca"]).value = "nenhum"; continue
        melhor = max(cand, key=lambda c: (confianca(nome, c["titulo"]) == "alta", confianca(nome, c["titulo"]) == "média", c["preco"] or 0))
        ws.cell(r, h["asin_sugerido"]).value = melhor["asin"]
        ws.cell(r, h["titulo_amazon"]).value = melhor["titulo"]
        ws.cell(r, h["preco_sugerido"]).value = melhor["preco"]
        ws.cell(r, h["confianca"]).value = confianca(nome, melhor["titulo"])
        n += 1
        if n % 20 == 0: print(f"  {n} sugestões…")
    print(f"descobrir: {n} sugestões gravadas (asin_sugerido / titulo_amazon / confianca)")


def aceitar(ws):
    h = garantir_colunas(ws); n = 0
    for r in range(2, ws.max_row + 1):
        g = lambda c: ws.cell(r, h[c]).value
        if g("asin") or not g("asin_sugerido"): continue
        if str(g("aceitar") or "").strip().lower() == "sim" or (g("confianca") == "alta" and str(g("aceitar") or "").strip().lower() != "não"):
            ws.cell(r, h["asin"]).value = g("asin_sugerido"); n += 1
    print(f"aceitar: {n} ASINs copiados para a coluna asin")


def atualizar(ws, api):
    h = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
    linhas = {}
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, h["asin"]).value
        if a and ws.cell(r, h["nome"]).value: linhas.setdefault(str(a).strip(), []).append(r)
    asins = list(linhas); n = 0; sem = []
    for i in range(0, len(asins), 10):
        lote = asins[i:i + 10]
        try:
            itens = api.get_items(lote)
        except Exception as e:
            print(f"  lote {lote[0]}…: erro {str(e)[:120]}"); continue
        achados = {}
        for it in itens or []:
            d = extrair(it); achados[d["asin"]] = d
        for a in lote:
            d = achados.get(a)
            for r in linhas[a]:
                if not d:
                    ws.cell(r, h["disponivel"]).value = "não"; ws.cell(r, h["atualizado_em"]).value = HOJE; sem.append(a); continue
                ws.cell(r, h["preco"]).value = d["preco"]; ws.cell(r, h["preco_lista"]).value = d["preco_lista"]
                ws.cell(r, h["disponivel"]).value = d["disponivel"]; ws.cell(r, h["url_imagem"]).value = d["url_imagem"]
                ws.cell(r, h["url_afiliado"]).value = d["url_afiliado"]; ws.cell(r, h["atualizado_em"]).value = HOJE
                n += 1
    print(f"atualizar: {n} linhas com preço/imagem · {len(sem)} ASINs sem retorno (marcados disponivel=não): {sem[:10]}")


def main():
    args = sys.argv[1:]
    if not args or args[0].startswith("--"): print(__doc__); sys.exit(1)
    arq = args[0]
    if "--teste" in args:
        api = cliente(); res = api.search_items(keywords="Catan Junior", item_count=1)
        it = (res.items or [None])[0]
        print("OK, versão", VERSION, "→", extrair(it) if it else "busca vazia"); return
    wb = load_workbook(arq); ws = wb["Catálogo"]
    if "--descobrir" in args: descobrir(ws, cliente())
    if "--aceitar" in args: aceitar(ws)
    if "--atualizar" in args: atualizar(ws, cliente())
    wb.save(arq); print("salvo:", arq)


if __name__ == "__main__":
    main()
