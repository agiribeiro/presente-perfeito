# -*- coding: utf-8 -*-
"""
Confere e preenche as colunas de referência (azuis) do Catálogo com dados da
API XML do BoardGameGeek. Rode na sua máquina (precisa de internet):

    pip install requests openpyxl
    python bgg_verificar.py catalogo_presentes.xlsx            # só confere o Catálogo
    python bgg_verificar.py catalogo_presentes.xlsx --candidatos   # idem para a aba Candidatos

O que ele faz, linha a linha:
  - bgg_id preenchido  → busca /thing?id=...&stats=1 e grava jogadores, duração,
    idade mínima do BGG, complexidade (weight), mecânicas, categorias, nota.
  - bgg_id vazio       → busca /search?query=nome_en&type=boardgame; se houver
    resultado com nome exatamente igual, usa; senão lista até 5 opções em
    observacoes para você escolher.
  - Em todo caso, grava em observacoes o que mudou em relação ao que estava.
  - Nunca toca nas colunas amarelas (sua curadoria). A sugestão automática de
    tags fica disponível em sugerir_tags() e é usada só para linhas cujo perfil
    está vazio (ou seja, candidatos novos).

A API do BGG pede ~2 s entre chamadas e às vezes responde 202 ("processando");
o script espera e tenta de novo.
"""
import os, sys, time, re
import xml.etree.ElementTree as ET
import requests
from openpyxl import load_workbook

BASE = "https://boardgamegeek.com/xmlapi2"   # sem "www": com www o token é ignorado
PAUSA = 2.0
BGG_TOKEN = (os.environ.get("BGG_TOKEN") or os.environ.get("BGG_API_TOKEN") or "").strip()


def _get(url, params, tentativas=5):
    headers = {"User-Agent": "catalogo-presentes/1.0"}
    if BGG_TOKEN:
        headers["Authorization"] = f"Bearer {BGG_TOKEN}"
    for i in range(tentativas):
        r = requests.get(url, params=params, timeout=30, headers=headers)
        if r.status_code == 202:
            time.sleep(3); continue
        if r.status_code in (429, 503):
            time.sleep(5 * (i + 1)); continue
        if r.status_code == 401:
            raise RuntimeError("BGG respondeu 401: a API agora exige token. Defina BGG_TOKEN "
                               "(gere em https://boardgamegeek.com/applications → Tokens).")
        r.raise_for_status()
        time.sleep(PAUSA)
        return ET.fromstring(r.content)
    raise RuntimeError("BGG não respondeu: " + url)


def buscar(nome_en):
    root = _get(f"{BASE}/search", {"query": nome_en, "type": "boardgame"})
    itens = []
    for it in root.findall("item"):
        nome = it.find("name").get("value")
        ano = it.find("yearpublished")
        itens.append((int(it.get("id")), nome, ano.get("value") if ano is not None else "?"))
    return itens


def thing(ids):
    root = _get(f"{BASE}/thing", {"id": ",".join(map(str, ids)), "stats": "1"})
    out = {}
    for it in root.findall("item"):
        g = lambda tag: it.find(tag).get("value") if it.find(tag) is not None else None
        links = lambda t: [l.get("value") for l in it.findall("link") if l.get("type") == t]
        stats = it.find("statistics/ratings")
        out[int(it.get("id"))] = dict(
            nome=next((n.get("value") for n in it.findall("name") if n.get("type") == "primary"), None),
            jogadores_min=int(g("minplayers") or 0), jogadores_max=int(g("maxplayers") or 0),
            duracao_min=int(g("playingtime") or 0), idade_bgg=int(g("minage") or 0),
            complexidade=round(float(stats.find("averageweight").get("value")), 2),
            nota_bgg=round(float(stats.find("average").get("value")), 1),
            mecanicas="; ".join(links("boardgamemechanic")),
            categorias="; ".join(links("boardgamecategory")),
        )
    return out


# ------------------------------------------------------------ sugestão de tags
def sugerir_tags(mecanicas, categorias, peso, jogadores_max, preco=None):
    """Mapa mecânica/categoria → perfil e notas 0-3. Espelha a aba Taxonomia."""
    m = (mecanicas or "").lower(); c = (categorias or "").lower()
    has = lambda *ks: any(k in m or k in c for k in ks)
    perfil = []
    if has("storytelling", "drawing", "word game"): perfil.append("criativo")
    if has("take that", "player elimination", "real-time") or (jogadores_max == 2 and has("abstract")): perfil.append("competitivo")
    if has("deduction", "memory", "pattern recognition", "puzzle", "trivia"): perfil.append("curioso")
    if has("tile placement", "stacking", "network", "route"): perfil.append("construção")
    if has("code breaking", "educational", "science", "math"): perfil.append("ciência")
    if has("art") or "dixit" in m: perfil.append("arte")
    if has("real-time", "dexterity", "flicking", "action / dexterity"): perfil.append("esportes")
    if (peso or 0) >= 1.5 or has("abstract strategy"): perfil.append("estratégia")
    if has("animals", "fantasy", "movies", "comic", "mythology"): perfil.append("personagens")
    if has("cooperative") or (jogadores_max or 0) >= 6: perfil.append("brincar em grupo")
    if has("stacking and balancing", "dexterity"): perfil.append("motor")
    if has("role playing"): perfil.append("faz de conta")
    perfil = list(dict.fromkeys(perfil))[:4]

    p = peso or 1.0
    diversao = 3 if p < 1.5 and has("dice", "real-time", "take that") else (2 if p < 1.5 else 1)
    educativo = 3 if has("deduction", "puzzle", "abstract strategy", "trivia", "word game") else (2 if has("set collection", "tile placement") else 1)
    criatividade = 3 if has("storytelling", "drawing") else (2 if has("voting", "communication") else 0)
    if preco is None: impressionar = 1
    elif preco > 200: impressionar = 3
    elif preco >= 120: impressionar = 2
    else: impressionar = 1
    if preco is None: cb = 2
    elif preco < 80: cb = 3
    elif preco <= 150: cb = 2
    elif preco <= 250: cb = 1
    else: cb = 0
    return dict(perfil=", ".join(perfil), diversao=diversao, educativo=educativo,
                criatividade=criatividade, impressionar=impressionar, custo_beneficio=cb)


# ------------------------------------------------------------ planilha
def processar(ws):
    hdr = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
    col = lambda nome: hdr[nome]
    ids, linhas_sem_id = {}, []
    for r in range(2, ws.max_row + 1):
        nome = ws.cell(r, col("nome")).value
        if not nome or str(nome).startswith("(exemplo)"):
            continue
        bid = ws.cell(r, col("bgg_id")).value
        if bid:
            ids.setdefault(int(bid), []).append(r)
        else:
            linhas_sem_id.append(r)

    # 1) linhas sem id: tentar busca por nome_en
    for r in linhas_sem_id:
        nome_en = ws.cell(r, col("nome_en")).value or ws.cell(r, col("nome")).value
        res = buscar(nome_en)
        exatos = [x for x in res if x[1].lower() == str(nome_en).lower()]
        obs = ws.cell(r, col("observacoes")).value or ""
        if len(exatos) == 1:
            bid = exatos[0][0]
            ws.cell(r, col("bgg_id")).value = bid
            ids.setdefault(bid, []).append(r)
            obs += f" | bgg_id {bid} encontrado por nome ({exatos[0][2]}) — confirmar"
        elif res:
            opcoes = "; ".join(f"{i}={n} ({a})" for i, n, a in res[:5])
            obs += f" | BGG: escolha o id: {opcoes}"
        else:
            obs += " | BGG: nada encontrado para " + str(nome_en)
        ws.cell(r, col("observacoes")).value = obs.strip(" |")

    # 2) buscar detalhes em lotes de 20
    todos = list(ids)
    for i in range(0, len(todos), 20):
        lote = todos[i:i + 20]
        dados = thing(lote)
        for bid in lote:
            d = dados.get(bid)
            for r in ids[bid]:
                obs = ws.cell(r, col("observacoes")).value or ""
                if not d:
                    ws.cell(r, col("observacoes")).value = (obs + f" | BGG: id {bid} não retornou").strip(" |")
                    continue
                mudancas = []
                for campo in ("jogadores_min", "jogadores_max", "duracao_min", "complexidade", "mecanicas", "categorias", "nota_bgg"):
                    antigo = ws.cell(r, col(campo)).value
                    novo = d[campo]
                    if antigo not in (None, "") and str(antigo) != str(novo) and campo not in ("mecanicas", "categorias", "nota_bgg"):
                        mudancas.append(f"{campo} {antigo}→{novo}")
                    ws.cell(r, col(campo)).value = novo
                if d["idade_bgg"] and d["idade_bgg"] != ws.cell(r, col("idade_min")).value:
                    mudancas.append(f"idade BGG={d['idade_bgg']} (planilha {ws.cell(r, col('idade_min')).value})")
                nome_bgg = d["nome"]
                nome_en = ws.cell(r, col("nome_en")).value
                if nome_en and nome_bgg and nome_bgg.lower() != str(nome_en).lower():
                    mudancas.append(f"nome no BGG: {nome_bgg}")
                # tags só para quem ainda não tem perfil
                if not ws.cell(r, col("perfil")).value:
                    preco = ws.cell(r, col("preco")).value or ws.cell(r, col("preco_ref_estimado")).value
                    s = sugerir_tags(d["mecanicas"], d["categorias"], d["complexidade"], d["jogadores_max"], preco)
                    for k, v in s.items():
                        ws.cell(r, col(k)).value = v
                    mudancas.append("tags sugeridas automaticamente")
                obs = re.sub(r"Números de referência são estimativas; rodar bgg_verificar\.py\.?", "", obs)
                obs = (obs + (" | BGG conferido" + (": " + "; ".join(mudancas) if mudancas else ""))).strip(" |")
                ws.cell(r, col("observacoes")).value = obs
        print(f"lote {i // 20 + 1}: {len(lote)} jogos conferidos")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    arq = sys.argv[1]
    aba = "Candidatos" if "--candidatos" in sys.argv else "Catálogo"
    wb = load_workbook(arq)
    processar(wb[aba])
    wb.save(arq)
    print("salvo:", arq, "aba", aba)
