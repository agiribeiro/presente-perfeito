# Presente Perfeito

Recomendador de presentes (jogos de tabuleiro) monetizado por links de afiliado da Amazon.com.br.

- `docs/` — o que vai ao ar no GitHub Pages: `index.html` (quiz), `catalogo.json` (dados), `privacidade.html`.
- `catalogo_presentes.xlsx` — a fonte de verdade: aba **Catálogo** (curadoria), **Candidatos**, **Rejeitados**, **Taxonomia**.
- `scripts/` — varredura Ludopedia/BGG, limpeza, migração, sincronização com a Amazon e geração do site.
- `.github/workflows/atualizar-precos.yml` — robô diário: busca preços na Creators API e publica `docs/catalogo.json`.

## Rotina

| Quando | Comando (na raiz, com `.venv` ativo) |
|---|---|
| Editei a planilha e quero publicar | `./publicar.sh` |
| Buscar jogos novos | `python3 scripts/varredura.py catalogo_presentes.xlsx --busca infantil` → `scripts/pos_varredura.py` → `scripts/migrar.py` |
| Achar ASINs na Amazon | `python3 scripts/amazon_sync.py catalogo_presentes.xlsx --descobrir` → conferir → `--aceitar` |
| Atualizar preços à mão | `python3 scripts/amazon_sync.py catalogo_presentes.xlsx --atualizar` |
| Testar o site local | `cd docs && python3 -m http.server` → http://localhost:8000 |

Credenciais ficam em variáveis de ambiente (`LUDOPEDIA_TOKEN`, `BGG_TOKEN`, `AMAZON_CRED_ID`, `AMAZON_CRED_SECRET`, `AMAZON_TAG`) e, no GitHub, em *Settings → Secrets and variables → Actions*. Nunca na planilha nem no código.
