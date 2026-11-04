# -*- coding: utf-8 -*-
"""Acrescenta a semente de LEGO e brinquedos ao Catálogo (pula nomes já existentes).
    python3 scripts/carregar_semente.py catalogo_presentes.xlsx
Depois: python3 scripts/amazon_sync.py catalogo_presentes.xlsx --descobrir  (acha ASIN/preço/foto)
"""
import sys, datetime as dt
from pathlib import Path
from openpyxl import load_workbook
sys.path.insert(0, str(Path(__file__).parent))
from semente_brinquedos import LEGO, BRINQUEDOS

arq = sys.argv[1] if len(sys.argv) > 1 else "catalogo_presentes.xlsx"
wb = load_workbook(arq); ws = wb["Catálogo"]
h = {c.value: i + 1 for i, c in enumerate(ws[1]) if c.value}
for col in ("busca_amazon",):
    if col not in h:
        ws.cell(1, ws.max_column + 1).value = col; h[col] = ws.max_column
ja = {str(ws.cell(r, h["nome"]).value).strip().lower() for r in range(2, ws.max_row + 1) if ws.cell(r, h["nome"]).value}
hoje = dt.date.today().isoformat(); n = 0
for tipo, lista in (("lego", LEGO), ("brinquedo", BRINQUEDOS)):
    for j in lista:
        if j["nome"].strip().lower() in ja: continue
        linha = {
            "nome": j["nome"], "tipo": tipo, "idade_min": j["idade_min"],
            "idade_ideal_min": j["ideal"][0], "idade_ideal_max": j["ideal"][1],
            "perfil": ", ".join(j["perfil"]), "diversao": j["prio"]["div"], "educativo": j["prio"]["edu"],
            "criatividade": j["prio"]["cri"], "impressionar": j["prio"]["imp"], "custo_beneficio": j["prio"]["cb"],
            "por_que": j["por_que"], "ativo": "não", "nome_en": j["nome"], "jogadores_min": 1, "jogadores_max": 1,
            "duracao_min": 0, "complexidade": 1.0, "editora": j.get("marca", "LEGO" if tipo == "lego" else ""),
            "origem": f"semente {tipo}", "aprovado_em": hoje, "preco_ref_estimado": j["preco"],
            "busca_amazon": j.get("busca", j["nome"]),
            "observacoes": "Nome genérico por linha/modelo; o amazon_sync --descobrir acha o produto atual.",
        }
        r = ws.max_row + 1
        for col, val in linha.items():
            if col in h: ws.cell(r, h[col]).value = val
        ja.add(j["nome"].strip().lower()); n += 1
wb.save(arq)
print(f"{n} itens acrescentados ao Catálogo ({len(LEGO)} LEGO + {len(BRINQUEDOS)} brinquedos na semente)")
