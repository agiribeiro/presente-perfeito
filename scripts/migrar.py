# -*- coding: utf-8 -*-
"""
Migra os candidatos que são jogos de tabuleiro para o Catálogo (colunas alinhadas) e manda
o que não é jogo (livro de RPG, livro-jogo, suplemento) para Rejeitados.

    python3 migrar.py catalogo_presentes.xlsx            # migra tudo que for jogo
    python3 migrar.py catalogo_presentes.xlsx --so-sim   # migra só as linhas com aprovar = sim

Regras de "não é jogo" (usa cache/ da Ludopedia):
  - mecânica RPG sem nenhuma mecânica de tabuleiro;
  - nome com cara de livro: Módulo Básico, Livro, Manual, Rulebook, Suplemento, Guia, Kit Introdutório,
    Livro-jogo / Fighting Fantasy ("FF 0x"), Your Town, GURPS, Numenera…
  - mecânicas só de "Escolha Narrativa / Narração" sem tabuleiro (livro-jogo).
Linhas com aprovar = não vão para Rejeitados com motivo "recusado na revisão".
"""
import sys, json, re, datetime as dt
from pathlib import Path
from openpyxl import load_workbook

CACHE = Path("cache"); HOJE = dt.date.today().isoformat()
TABULEIRO = ("Movimento em Grades", "Tabuleiro Modular", "Colocação de Peças", "Gerenciamento de Mãos",
             "Construção de Baralho", "Movimento de Área", "Grade Hexagonal", "Alocação de Trabalhadores",
             "Movimento Ponto-a-Ponto", "Pontos de Ação", "Colecionar Conjuntos", "Campanha / Batalhas",
             "Seleção de Ação", "Mapa:", "Criação de Baralho", "Legado", "Comunicação Limitada",
             "Rolar e Mover", "Leilão", "Negociação", "Maioria de Área", "Draft", "Memória", "Padrões",
             "Tempo Real", "Empilhamento", "Destreza", "Pegar e Entregar", "Trocas", "Votação", "Blefe",
             "Controle de Área", "Posicionamento", "Grade", "Construção de Rotas", "Força sua sorte", "Force sua sorte",
             "Papel e Caneta")  # roll & write também é jogo quando não há RPG
LIVRO = re.compile(r"módulo básico|livro|manual|rulebook|suplemento|guia d|kit (inicial|introdut)|\bFF \d|your town|gurps|numenera|"
                   r"roleplaying|\bRPG\b|core rulebook|player.?s handbook|aventura (solo|para)|sourcebook|minha primeira aventura", re.I)
NARRATIVO = {"RPG", "Rolagem de Dados", "Narração de Histórias", "Jogadores com Diferentes Habilidades", "Escolha Narrativa",
             "Cooperativo", "Jogo em Equipe", "Cenário / Missão / Campanha", "Solo / Solitário", "Resolução com Testes de Estatísticas",
             "Papel e Caneta", "Simulação"}


def ludo(i):
    f = CACHE / f"ludo__jogos_{i}_.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else None


def motivo_nao_jogo(nome, d):
    mec = [m.get("nm_mecanica", "") for m in (d or {}).get("mecanicas", [])]
    tem_rpg = any(m.upper() == "RPG" for m in mec)
    tem_tab = any(t in m for m in mec for t in TABULEIRO if not (t == "Papel e Caneta" and tem_rpg))
    if LIVRO.search(nome or ""):
        return "livro / suplemento / livro-jogo, não é jogo de tabuleiro"
    if tem_rpg and not tem_tab:
        return "RPG (livro de regras), não é jogo de tabuleiro"
    return None


def main():
    arq = sys.argv[1]; so_sim = "--so-sim" in sys.argv
    wb = load_workbook(arq)
    cand, cat, rej = wb["Candidatos"], wb["Catálogo"], wb["Rejeitados"]
    hc = {c.value: i + 1 for i, c in enumerate(cand[1]) if c.value}
    hk = {c.value: i + 1 for i, c in enumerate(cat[1]) if c.value}
    ja = set()
    for row in cat.iter_rows(min_row=2, values_only=True):
        if row[hk["nome"] - 1]: ja.add(str(row[hk["nome"] - 1]).strip().lower())
    migrados, rejeitados, mantidos, apagar = [], [], 0, []
    for r in range(2, cand.max_row + 1):
        g = lambda c: cand.cell(r, hc[c]).value
        nome = g("nome")
        if not nome or str(nome).startswith("(exemplo)"): continue
        aprov = str(g("aprovar") or "").strip().lower()
        d = ludo(int(g("ludopedia_id"))) if g("ludopedia_id") else None
        motivo = "recusado na revisão" if aprov == "não" else motivo_nao_jogo(str(nome), d)
        if motivo:
            rej.append([g("asin"), nome, g("bgg_id"), motivo, HOJE]); rejeitados.append((nome, motivo)); apagar.append(r); continue
        if so_sim and aprov != "sim":
            mantidos += 1; continue
        if str(nome).strip().lower() in ja:
            rej.append([g("asin"), nome, g("bgg_id"), "já estava no Catálogo", HOJE]); apagar.append(r); continue
        linha = [None] * len(hk)
        for col, i in hk.items():
            if col in hc: linha[i - 1] = g(col)
        linha[hk["ativo"] - 1] = "não"
        linha[hk["aprovado_em"] - 1] = HOJE
        linha[hk["origem"] - 1] = (g("origem") or "varredura") + (" · aprovado" if aprov == "sim" else " · migrado sem revisão")
        cat.append(linha); ja.add(str(nome).strip().lower()); migrados.append(nome); apagar.append(r)
    for r in sorted(apagar, reverse=True): cand.delete_rows(r)
    wb.save(arq)
    print(f"migrados: {len(migrados)} · rejeitados: {len(rejeitados)} · mantidos em Candidatos: {mantidos}")
    print("rejeitados:", "; ".join(f"{n[:30]} ({m.split(',')[0]})" for n, m in rejeitados))
    print("migrados (amostra):", "; ".join(str(n)[:28] for n in migrados[:40]))


if __name__ == "__main__":
    main()
