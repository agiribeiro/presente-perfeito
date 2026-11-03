# -*- coding: utf-8 -*-
"""
Varredura Ludopedia + BGG → aba Candidatos da planilha.

Como rodar (na sua máquina, mesma pasta de catalogo_presentes.xlsx e bgg_verificar.py):

    pip install requests openpyxl
    export LUDOPEDIA_TOKEN="seu_token_da_api"        # Windows: set LUDOPEDIA_TOKEN=...
    python varredura.py catalogo_presentes.xlsx                  # padrão: top 600 do ranking Ludopedia
    python varredura.py catalogo_presentes.xlsx --max 1500       # mais fundo
    python varredura.py catalogo_presentes.xlsx --busca "infantil" --busca "kids" --busca "junior"
    python varredura.py catalogo_presentes.xlsx --sem-bgg        # pula o enriquecimento no BGG (mais rápido)
    python varredura.py catalogo_presentes.xlsx --idade-max 99 --tempo-max 999   # sem filtro de idade/tempo
    python varredura.py catalogo_presentes.xlsx --idade-max 12 --tempo-max 90    # até 12 anos e 90 min

O que faz:
  1. Lista jogos base na Ludopedia (ranking paginado + buscas opcionais).
  2. Para cada um, pega os detalhes (idade mínima, jogadores, tempo, mecânicas, categorias,
     descrição, nome original) e FILTRA por idade mínima da caixa e tempo de partida.
     Padrão: sem filtro (idade até 99, tempo até 999 min). Use --idade-max / --tempo-max para restringir.
  3. Opcionalmente busca o jogo no BGG pelo nome original para trazer weight, nota e mecânicas em inglês.
  4. Descarta o que já está em Catálogo, Candidatos ou Rejeitados (por nome, bgg_id ou ludopedia_id).
  5. Escreve os novos na aba Candidatos com perfil e notas SUGERIDOS (sugerir_tags) e um rascunho
     de "por_que" tirado da descrição. Você só marca aprovar = sim/não.

Tudo o que a API devolve fica em cache/ (json) para re-rodar sem gastar chamadas. Se a Ludopedia
mudar o nome de algum campo, ajuste o dicionário CAMPOS abaixo — o script imprime as chaves do
primeiro jogo que receber para facilitar.
"""
import os, sys, json, time, re, datetime as dt, unicodedata
from pathlib import Path
import requests
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
from bgg_verificar import sugerir_tags, buscar as bgg_buscar, thing as bgg_thing

# ------------------------------------------------------------ parâmetros
IDADE_MAX = 99          # padrão: sem filtro de idade (ajuste com --idade-max)
TEMPO_MAX = 999         # padrão: sem filtro de tempo em minutos (ajuste com --tempo-max)
LUDO = "https://ludopedia.com.br/api/v1"
TOKEN = os.environ.get("LUDOPEDIA_TOKEN", "").strip()
PAUSA_LUDO = 0.7        # segundos entre chamadas à Ludopedia
CACHE = Path("cache"); CACHE.mkdir(exist_ok=True)
HOJE = dt.date.today().isoformat()

# nomes dos campos como a API da Ludopedia costuma devolver; ajuste se o print mostrar outros
CAMPOS = dict(
    lista="jogos", id="id_jogo", nome="nm_jogo", nome_original="nm_original", ano="ano_publicacao",
    jog_min="qt_jogadores_min", jog_max="qt_jogadores_max", tempo="vl_tempo_jogo", idade="idade_minima",
    descricao="descricao", mecanicas="mecanicas", categorias="categorias", temas="temas",
    nome_item="nm_mecanica|nm_categoria|nm_tema", tipo="tp_jogo", link="link",
)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


# ------------------------------------------------------------ Ludopedia
def ludo_get(path, params=None, chave=None):
    chave = chave or (path + "?" + "&".join(f"{k}={v}" for k, v in sorted((params or {}).items())))
    arq = CACHE / ("ludo_" + re.sub(r"[^a-zA-Z0-9]+", "_", chave) + ".json")
    if arq.exists():
        return json.loads(arq.read_text(encoding="utf-8"))
    if not TOKEN:
        sys.exit("Defina a variável de ambiente LUDOPEDIA_TOKEN com o token da API da Ludopedia.")
    r = requests.get(LUDO + path, params=params, timeout=30,
                     headers={"Authorization": f"Bearer {TOKEN}", "User-Agent": "catalogo-presentes/1.0"})
    if r.status_code == 429:
        time.sleep(10); return ludo_get(path, params, chave)
    r.raise_for_status()
    data = r.json()
    arq.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    time.sleep(PAUSA_LUDO)
    return data


def ludo_listar(maximo, buscas):
    """Ids de jogos base: ranking paginado + termos de busca."""
    ids, vistos = [], set()

    def add(lista):
        for j in lista or []:
            i = j.get(CAMPOS["id"])
            if i and i not in vistos and (j.get(CAMPOS["tipo"], "b") in ("b", None)):
                vistos.add(i); ids.append(i)

    page, rows = 1, 100
    while len(ids) < maximo:
        d = ludo_get("/jogos", dict(tp_jogo="b", ordem="rank", rows=rows, page=page))
        lista = d.get(CAMPOS["lista"]) or []
        if not lista: break
        add(lista); page += 1
        print(f"  ranking página {page-1}: {len(ids)} jogos")
    for termo in buscas:
        page = 1
        while True:
            d = ludo_get("/jogos", dict(search=termo, tp_jogo="b", rows=rows, page=page))
            lista = d.get(CAMPOS["lista"]) or []
            if not lista: break
            add(lista); page += 1
        print(f"  busca '{termo}': {len(ids)} jogos acumulados")
    return ids[:maximo] if not buscas else ids


def nomes_de(lista):
    out = []
    for it in lista or []:
        if isinstance(it, str): out.append(it); continue
        for k in CAMPOS["nome_item"].split("|"):
            if it.get(k): out.append(it[k]); break
        else:
            out.append(next((v for v in it.values() if isinstance(v, str)), ""))
    return [x for x in out if x]


def ludo_detalhe(i):
    d = ludo_get(f"/jogos/{i}")
    g = lambda k: d.get(CAMPOS[k])
    tempo = g("tempo") or d.get("vl_tempo_jogo_max") or 0
    return dict(
        ludopedia_id=i, nome=g("nome"), nome_en=g("nome_original") or g("nome"), ano=g("ano"),
        jogadores_min=g("jog_min") or 0, jogadores_max=g("jog_max") or 0, duracao_min=int(tempo or 0),
        idade_min=int(g("idade") or 0), descricao=re.sub(r"<[^>]+>", " ", str(g("descricao") or "")),
        mecanicas_pt="; ".join(nomes_de(g("mecanicas"))), categorias_pt="; ".join(nomes_de(g("categorias"))),
        temas_pt="; ".join(nomes_de(g("temas"))), link=g("link"), raw_keys=sorted(d.keys()),
    )


# ------------------------------------------------------------ BGG (enriquecimento)
def bgg_enriquecer(nome_en, ano):
    arq = CACHE / ("bgg_" + re.sub(r"[^a-zA-Z0-9]+", "_", norm(nome_en)) + ".json")
    if arq.exists():
        return json.loads(arq.read_text(encoding="utf-8"))
    res = bgg_buscar(nome_en)
    exatos = [x for x in res if norm(x[1]) == norm(nome_en)]
    if ano:
        por_ano = [x for x in exatos if str(x[2]) == str(ano)]
        exatos = por_ano or exatos
    out = None
    if exatos:
        bid = exatos[0][0]
        d = bgg_thing([bid]).get(bid)
        if d:
            d["bgg_id"] = bid; out = d
    arq.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return out


# ------------------------------------------------------------ taxonomia PT → EN (para sugerir_tags)
PT_EN = {
    "cooperativo": "cooperative game", "dedução": "deduction", "memória": "memory", "reconhecimento de padrões": "pattern recognition",
    "tempo real": "real-time", "destreza": "dexterity", "empilhamento": "stacking and balancing", "colocação de peças": "tile placement",
    "coleção de componentes": "set collection", "construção de rotas": "route building", "narrativa": "storytelling",
    "desenho": "drawing", "jogo de palavras": "word game", "rolagem de dados": "dice rolling", "infantil": "children's game",
    "animais": "animals", "fantasia": "fantasy", "quebra-cabeça": "puzzle", "abstrato": "abstract strategy", "party game": "party game",
    "educativo": "educational", "trivia": "trivia", "votação": "voting", "eliminação de jogadores": "player elimination",
}


def traduzir(txt):
    t = (txt or "").lower()
    return "; ".join(en for pt, en in PT_EN.items() if pt in t)


def rascunho_por_que(desc):
    desc = re.sub(r"\s+", " ", desc or "").strip()
    frase = re.split(r"(?<=[.!?])\s", desc)[0] if desc else ""
    if len(frase) > 170: frase = frase[:167].rsplit(" ", 1)[0] + "…"
    return (frase + " (rascunho automático)").strip()


# ------------------------------------------------------------ planilha
def existentes(wb):
    nomes, bggs, ludos = set(), set(), set()
    for aba in ("Catálogo", "Candidatos", "Rejeitados"):
        if aba not in wb.sheetnames: continue
        ws = wb[aba]; hdr = {c.value: i for i, c in enumerate(ws[1]) if c.value}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if hdr.get("nome") is not None and row[hdr["nome"]]: nomes.add(norm(row[hdr["nome"]]))
            if hdr.get("nome_en") is not None and row[hdr["nome_en"]]: nomes.add(norm(row[hdr["nome_en"]]))
            if hdr.get("bgg_id") is not None and row[hdr["bgg_id"]]: bggs.add(int(row[hdr["bgg_id"]]))
            if hdr.get("ludopedia_id") is not None and row[hdr["ludopedia_id"]]: ludos.add(int(row[hdr["ludopedia_id"]]))
    return nomes, bggs, ludos


def main():
    args = sys.argv[1:]
    if not args or args[0].startswith("--"):
        print(__doc__); sys.exit(1)
    arq = args[0]
    maximo = int(args[args.index("--max") + 1]) if "--max" in args else 600
    buscas = [args[i + 1] for i, a in enumerate(args) if a == "--busca"]
    usar_bgg = "--sem-bgg" not in args
    global IDADE_MAX, TEMPO_MAX
    if "--idade-max" in args: IDADE_MAX = int(args[args.index("--idade-max") + 1])
    if "--tempo-max" in args: TEMPO_MAX = int(args[args.index("--tempo-max") + 1])
    print(f"filtro: idade mínima até {IDADE_MAX} anos, partida até {TEMPO_MAX} min")

    wb = load_workbook(arq)
    ws = wb["Candidatos"]
    hdr = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
    nomes, bggs, ludos = existentes(wb)
    print(f"já conhecidos: {len(nomes)} nomes, {len(bggs)} bgg_ids, {len(ludos)} ludopedia_ids")

    print("1) listando na Ludopedia…")
    ids = [i for i in ludo_listar(maximo, buscas) if i not in ludos]
    print(f"   {len(ids)} jogos para examinar")

    novos, filtrados, repetidos = [], 0, 0
    for n, i in enumerate(ids, 1):
        try:
            j = ludo_detalhe(i)
        except requests.HTTPError as e:
            print(f"   [{i}] erro {e.response.status_code}; pulando"); continue
        if n == 1:
            print("   campos devolvidos pela Ludopedia:", ", ".join(j["raw_keys"]))
        if not j["nome"]: continue
        if j["idade_min"] > IDADE_MAX or (j["duracao_min"] and j["duracao_min"] > TEMPO_MAX) or j["idade_min"] == 0:
            filtrados += 1; continue
        if norm(j["nome"]) in nomes or norm(j["nome_en"]) in nomes:
            repetidos += 1; continue
        b = None
        if usar_bgg:
            try:
                b = bgg_enriquecer(j["nome_en"], j["ano"])
            except Exception as e:  # 401 sem token, 5xx, timeout…
                print(f"   BGG indisponível ({str(e)[:120]}). Seguindo só com a Ludopedia; "
                      f"rode depois bgg_verificar.py --candidatos para completar.")
                usar_bgg = False
        if b and b["bgg_id"] in bggs:
            repetidos += 1; continue
        novos.append((j, b))
        if n % 25 == 0: print(f"   {n}/{len(ids)} examinados · {len(novos)} novos · {filtrados} fora do filtro · {repetidos} repetidos")

    print(f"2) gravando {len(novos)} candidatos na aba Candidatos…")
    r = ws.max_row + 1
    for j, b in novos:
        mec_en = (b["mecanicas"] if b else "") or traduzir(j["mecanicas_pt"])
        cat_en = (b["categorias"] if b else "") or traduzir(j["categorias_pt"] + " " + j["temas_pt"])
        peso = b["complexidade"] if b else None
        s = sugerir_tags(mec_en, cat_en, peso, j["jogadores_max"], None)
        linha = {
            "aprovar": "", "data_sugestao": HOJE, "asin": "", "nome": j["nome"], "tipo": "jogo",
            "idade_min": j["idade_min"], "idade_ideal_min": j["idade_min"], "idade_ideal_max": 99 if j["idade_min"] >= 10 else j["idade_min"] + 4,
            "perfil": s["perfil"], "diversao": s["diversao"], "educativo": s["educativo"], "criatividade": s["criatividade"],
            "impressionar": s["impressionar"], "custo_beneficio": s["custo_beneficio"], "ocasiao": "",
            "por_que": rascunho_por_que(j["descricao"]), "ativo": "não",
            "nome_en": j["nome_en"], "bgg_id": b["bgg_id"] if b else None, "ludopedia_id": j["ludopedia_id"],
            "jogadores_min": (b or j)["jogadores_min"], "jogadores_max": (b or j)["jogadores_max"],
            "duracao_min": (b or j)["duracao_min"], "complexidade": peso,
            "mecanicas": mec_en or j["mecanicas_pt"], "categorias": cat_en or j["categorias_pt"],
            "nota_bgg": b["nota_bgg"] if b else None, "editora": "",
            "origem": "varredura ludopedia" + ("+bgg" if b else ""),
            "observacoes": f"Ludopedia: {j['link'] or ''} · mecânicas PT: {j['mecanicas_pt']} · tags sugeridas automaticamente; revisar idade ideal e por_que.",
        }
        for col, val in linha.items():
            if col in hdr: ws.cell(r, hdr[col]).value = val
        r += 1
    wb.save(arq)
    print(f"salvo em {arq}. Resumo: {len(novos)} novos · {filtrados} fora do filtro (idade>{IDADE_MAX} ou tempo>{TEMPO_MAX}) · {repetidos} já conhecidos")


if __name__ == "__main__":
    main()
