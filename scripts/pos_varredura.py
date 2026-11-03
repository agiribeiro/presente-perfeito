# -*- coding: utf-8 -*-
"""
Limpeza pós-varredura (offline: usa só a planilha e o cache/ da Ludopedia).

    python3 pos_varredura.py catalogo_presentes.xlsx

1. Livros de RPG (mecânica "RPG" na Ludopedia): em Candidatos → movidos para Rejeitados;
   no Catálogo → marcados "⚠ RPG" em observacoes (não movo o que você aprovou).
2. complexidade vazia → vl_complexidade da comunidade Ludopedia (escala 1-5, igual ao BGG).
3. idade_ideal_min vazia ou igual a idade_min → vl_idade_sugerida da comunidade, quando houver votos.
4. por_que só com "(rascunho automático)" → rascunho estrutural em português a partir de
   mecânicas, temas, jogadores e tempo. Continua marcado para você reescrever.
5. nota_ludopedia ← vl_rejogabilidade (0-5) como proxy, e editora fica vazia (a API não traz).
Nunca toca em perfil, notas 0-3, ativo ou asin.
"""
import sys, json, re, datetime as dt
from pathlib import Path
from openpyxl import load_workbook

CACHE = Path("cache")
HOJE = dt.date.today().isoformat()


def ludo(i):
    f = CACHE / f"ludo__jogos_{i}_.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else None


def nomes(lst, k):
    return [x.get(k) for x in (lst or []) if isinstance(x, dict) and x.get(k)]


TABULEIRO = ("Movimento em Grades", "Tabuleiro Modular", "Colocação de Peças", "Gerenciamento de Mãos",
             "Construção de Baralho", "Movimento de Área", "Grade Hexagonal", "Alocação de Trabalhadores",
             "Movimento Ponto-a-Ponto", "Pontos de Ação", "Colecionar Conjuntos", "Campanha / Batalhas",
             "Seleção de Ação", "Mapa:", "Criação de Baralho", "Legado", "Dedução", "Comunicação Limitada")
LIVRO = re.compile(r"livro|manual|rulebook|kit (inicial|introdut)|suplemento|guia d|cenário|módulo|\bRPG\b|roleplaying|edição revisada", re.I)


def eh_rpg(d):
    """Livro de RPG = tem mecânica RPG, nenhuma mecânica de tabuleiro, e (papel e caneta ou nome de livro)."""
    mec = nomes(d.get("mecanicas"), "nm_mecanica")
    if not any(m.strip().upper() == "RPG" for m in mec): return False
    if any(t in m for m in mec for t in TABULEIRO): return False          # Gloomhaven, Jornadas, Arcadia Quest…
    return "Papel e Caneta" in mec or bool(LIVRO.search(d.get("nm_jogo") or "")) or len(mec) <= 2


def rascunho(d):
    mec = [m for m in nomes(d.get("mecanicas"), "nm_mecanica") if m.upper() != "RPG"][:2]
    tem = nomes(d.get("temas"), "nm_tema")[:2]
    jmin, jmax = d.get("qt_jogadores_min") or 0, d.get("qt_jogadores_max") or 0
    jog = f"{jmin} jogador" if jmin == jmax == 1 else (f"{jmin} jogadores" if jmin == jmax else f"{jmin} a {jmax} jogadores")
    t = d.get("vl_tempo_jogo") or 0
    partes = []
    if mec: partes.append("jogo de " + " e ".join(m.lower() for m in mec))
    if tem: partes.append("com tema de " + " e ".join(x.lower() for x in tem))
    frase = (", ".join(partes) or "jogo de tabuleiro").capitalize()
    frase += f"; {jog}" + (f", cerca de {t} min" if t else "") + "."
    return frase + " (rascunho automático — reescrever)"


def processar(ws, aba, rejeitados):
    hdr = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
    g = lambda r, c: ws.cell(r, hdr[c]).value if c in hdr else None
    s = lambda r, c, v: ws.cell(r, hdr[c]).__setattr__("value", v) if c in hdr else None
    mover, n_rpg, n_cx, n_id, n_pq = [], 0, 0, 0, 0
    for r in range(2, ws.max_row + 1):
        li = g(r, "ludopedia_id")
        if not li or not g(r, "nome"): continue
        d = ludo(int(li))
        if not d: continue
        if eh_rpg(d):
            n_rpg += 1
            if aba == "Candidatos":
                rejeitados.append([g(r, "asin"), g(r, "nome"), g(r, "bgg_id"), "RPG (livro de regras), não é jogo de tabuleiro", HOJE])
                mover.append(r); continue
            else:
                obs = str(g(r, "observacoes") or "")
                if "⚠ RPG" not in obs: s(r, "observacoes", "⚠ RPG (livro de regras) · " + obs)
        com = d.get("comunidade") or {}
        if not g(r, "complexidade") and com.get("vl_complexidade"):
            s(r, "complexidade", round(float(com["vl_complexidade"]), 2)); n_cx += 1
        if com.get("vl_idade_sugerida") and (com.get("qt_votos_idade") or 0) >= 3 and g(r, "idade_ideal_min") in (None, g(r, "idade_min")):
            s(r, "idade_ideal_min", int(com["vl_idade_sugerida"])); n_id += 1
        if com.get("vl_rejogabilidade") and not g(r, "nota_ludopedia"):
            s(r, "nota_ludopedia", round(float(com["vl_rejogabilidade"]), 1))
        pq = str(g(r, "por_que") or "").strip()
        if pq in ("", "(rascunho automático)"):
            s(r, "por_que", rascunho(d)); n_pq += 1
    for r in sorted(mover, reverse=True):
        ws.delete_rows(r)
    print(f"{aba}: {n_rpg} RPG {'movidos para Rejeitados' if aba=='Candidatos' else 'marcados'} · "
          f"{n_cx} complexidades · {n_id} idades sugeridas · {n_pq} rascunhos de por_que")


def main():
    arq = sys.argv[1] if len(sys.argv) > 1 else "catalogo_presentes.xlsx"
    wb = load_workbook(arq)
    rej = []
    processar(wb["Candidatos"], "Candidatos", rej)
    processar(wb["Catálogo"], "Catálogo", rej)
    ws = wb["Rejeitados"]
    for linha in rej: ws.append(linha)
    wb.save(arq)
    print(f"salvo: {arq} · Rejeitados +{len(rej)}")


if __name__ == "__main__":
    main()
