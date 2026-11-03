#!/bin/bash
# Publica a planilha e o site. Rode de dentro da pasta do repositório (clone fora do Google Drive):
#   ./publicar.sh                       # usa a planilha que está nesta pasta
#   ./publicar.sh "/caminho/da/pasta/Presenteperfeito"   # copia a planilha de lá antes
set -e
cd "$(dirname "$0")"
if [ -n "$1" ]; then cp "$1/catalogo_presentes.xlsx" ./catalogo_presentes.xlsx; fi
[ -d .venv ] && source .venv/bin/activate
python3 scripts/exportar_json.py catalogo_presentes.xlsx --saida docs/catalogo.json
python3 scripts/montar_site.py
git pull --rebase --autostash
git add -A
git commit -m "Catálogo: $(date +%Y-%m-%d\ %H:%M)" || echo "nada novo para publicar"
git push
echo "publicado. O GitHub Pages atualiza em 1–2 minutos."
