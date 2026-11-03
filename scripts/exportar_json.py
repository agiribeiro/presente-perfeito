# -*- coding: utf-8 -*-
"""Exporta a aba Catálogo para catalogo.json (o que o quiz lê).
    python3 scripts/exportar_json.py catalogo_presentes.xlsx [--so-ativos] [--saida docs/catalogo.json]
"""
import sys, json
from openpyxl import load_workbook

arq = sys.argv[1] if len(sys.argv) > 1 else "catalogo_presentes.xlsx"
so_ativos = "--so-ativos" in sys.argv
saida = sys.argv[sys.argv.index("--saida") + 1] if "--saida" in sys.argv else "catalogo.json"
ws = load_workbook(arq, data_only=True)["Catálogo"]
hdr = [c.value for c in ws[1]]
out = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
    d = dict(zip(hdr, row))
    if not d["nome"]: continue
    if so_ativos and d["ativo"] != "sim": continue
    pq = str(d["por_que"] or "").replace(" (rascunho automático — reescrever)", "").replace("(rascunho automático)", "").strip()
    out.append(dict(
        id=i, nome=d["nome"], nome_en=d["nome_en"] or d["nome"], tipo=d["tipo"] or "jogo", asin=d["asin"] or "",
        idade_min=int(d["idade_min"] or 0), ideal_min=int(d["idade_ideal_min"] or d["idade_min"] or 0),
        ideal_max=int(d["idade_ideal_max"] or 99),
        jog_min=int(d["jogadores_min"] or 0), jog_max=int(d["jogadores_max"] or 0), dur=int(d["duracao_min"] or 0),
        peso=float(d["complexidade"] or 2),
        perfil=[t.strip() for t in str(d["perfil"] or "").split(",") if t.strip()],
        div=int(d["diversao"] or 0), edu=int(d["educativo"] or 0), cri=int(d["criatividade"] or 0),
        imp=int(d["impressionar"] or 0), cb=int(d["custo_beneficio"] or 0),
        preco=float(d["preco"] or d["preco_ref_estimado"] or 0), preco_lista=float(d["preco_lista"] or 0),
        disponivel=(d["disponivel"] in (None, "", "sim", True)), url_imagem=d["url_imagem"] or "", url_afiliado=d["url_afiliado"] or "",
        por_que=pq, editora=d["editora"] or "", ativo=(d["ativo"] == "sim"),
        atualizado_em=str(d["atualizado_em"] or ""),
    ))
json.dump(out, open(saida, "w", encoding="utf-8"), ensure_ascii=False)
print(f"{saida}: {len(out)} jogos ({'só ativos' if so_ativos else 'todos'})")
