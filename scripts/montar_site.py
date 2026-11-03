# -*- coding: utf-8 -*-
"""Rode na raiz do repositório: python3 scripts/montar_site.py
Gera docs/index.html (lê catalogo.json por fetch) a partir de quiz_head.html + quiz_tail.html.
Também gera presente-perfeito.html (versão com catálogo embutido, usada no artifact/preview)."""
import json, re
from pathlib import Path
AQUI = Path(__file__).parent

head = (AQUI / "quiz_head.html").read_text(encoding="utf-8")
tail = (AQUI / "quiz_tail.html").read_text(encoding="utf-8")
cat = json.dumps(json.load(open("docs/catalogo.json", encoding="utf-8")), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

# --- versão embutida (artifact)
Path("presente-perfeito.html").write_text(head + cat + "\n" + tail, encoding="utf-8")

# --- versão site: sem JSON inline, com fetch
h = head.replace('<script id="catalogo" type="application/json">\n', "")
h = h.replace('<meta name="claude:viewport-fit-sides" content="cover">\n', "")
h = h.replace("Jogos de tabuleiro, protótipo v0.1", "Jogos de tabuleiro").replace("Jogos de tabuleiro, protótipo v0.2", "Jogos de tabuleiro")
h = h.replace('<p>Catálogo: <span id="ncat"></span> jogos · <span id="ncom"></span> com preço da Amazon (os demais levam à busca na Amazon).</p>',
              '<p>Catálogo: <span id="ncat"></span> jogos · <span id="ncom"></span> com preço da Amazon · preços atualizados em <span id="atualizado"></span> · <a href="privacidade.html">Política de privacidade</a></p>')
t = tail.replace('const CATALOGO = JSON.parse(document.getElementById("catalogo").textContent);', "let CATALOGO = [];")
t = t.replace('$("#ncat").textContent=CATALOGO.length; $("#ncom").textContent=CATALOGO.filter(j=>j.url_afiliado).length;', "")
t = t.replace("renderSteps(); renderQuiz();\n</script>", '''function iniciar(){
  $("#ncat").textContent=CATALOGO.length;
  $("#ncom").textContent=CATALOGO.filter(j=>j.url_afiliado).length;
  const dt=CATALOGO.map(j=>j.atualizado_em).filter(Boolean).sort().pop();
  $("#atualizado").textContent=dt?dt.slice(0,10).split("-").reverse().join("/"):"—";
  renderSteps(); renderQuiz();
}
fetch("catalogo.json",{cache:"no-cache"}).then(r=>r.json()).then(d=>{ CATALOGO=d.filter(j=>j.ativo!==false || !window.SO_ATIVOS); iniciar(); })
  .catch(e=>{ quiz.innerHTML='<p class="hint">Não consegui carregar o catálogo (catalogo.json). Se abriu o arquivo direto do computador, rode <code>python3 -m http.server</code> na pasta e acesse localhost:8000.</p>'; });
</script>''')
assert "iniciar()" in t and "let CATALOGO" in t
site = Path("docs"); site.mkdir(exist_ok=True)
pagina = ('<!doctype html>\n<html lang="pt-BR">\n<head>\n<meta charset="utf-8">\n'
          '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
          '<meta name="description" content="Responda 6 perguntas e receba 5 jogos de tabuleiro que combinam com quem vai ganhar o presente.">\n'
          '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n'
          + h.split("<style>")[0].replace("<title>", "<title>").strip() + "\n<style>" + h.split("<style>", 1)[1].split("</style>")[0] + "</style>\n</head>\n<body>\n"
          + h.split("</style>", 1)[1] + t + "\n</body>\n</html>\n")
(site / "index.html").write_text(pagina, encoding="utf-8")
print("docs/index.html e presente-perfeito.html gerados")
