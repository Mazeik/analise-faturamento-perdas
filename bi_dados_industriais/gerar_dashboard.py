"""
Dashboard do BI de Dados Industriais (DADOS 100% FICTÍCIOS).

Lê as tabelas geradas por gerar_bi.py (pasta saida/) e monta uma página HTML única, com gráficos em SVG
desenhados à mão (sem bibliotecas externas), e tira "prints" das três abas com o Chrome em modo headless.

Como rodar (depois de python gerar_bi.py):
    python gerar_dashboard.py

Saídas:
    dashboard/index.html                 -> abra no navegador (tem 3 abas clicáveis)
    graficos/dashboard_*.png             -> prints usados no README (precisam do Google Chrome instalado)
"""
import math
import shutil
import subprocess
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).parent
SAIDA, GRAFICOS, DASH = PASTA / "saida", PASTA / "graficos", PASTA / "dashboard"
DASH.mkdir(exist_ok=True)

AZUL, VERDE, VERM, AMAR, CINZA, TEXTO = "#1F5FA8", "#2E8B57", "#B3261E", "#E3A008", "#B0B7C0", "#595959"
COR_CLASSE = {"A": AZUL, "B": AMAR, "C": CINZA}

bal = pd.read_csv(SAIDA / "balanco_de_peso.csv", index_col=0)
abc = pd.read_csv(SAIDA / "curva_abc_clientes.csv")
cob = pd.read_csv(SAIDA / "cobertura_papel.csv", index_col=0)
MESES = list(bal.index)                       # meses com faturamento lançado (jan a ago)
TODOS = list(cob.index)                       # todos os meses (jan a out)

# ---------------------------------------------------------------------
# Funções de desenho (SVG simples)
# ---------------------------------------------------------------------
def br(x, casas=0):
    """Número no formato brasileiro (vírgula decimal e ponto de milhar)."""
    s = f"{x:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def passo(vmax):
    """Escolhe um passo 'redondo' (1, 2 ou 5 x 10^n) para a escala do eixo."""
    bruto = vmax / 4
    base = 10 ** math.floor(math.log10(bruto))
    return min((m * base for m in (1, 2, 5, 10) if m * base >= bruto))


class Grafico:
    """Área de desenho com eixos. Guarda as escalas para os métodos de barras e linhas."""
    def __init__(self, cats, vmin, vmax, w=560, h=250, fmt="{:.0f}"):
        self.cats, self.w, self.h = cats, w, h
        self.ml, self.mr, self.mt, self.mb = 46, 12, 10, 26
        self.iw, self.ih = w - self.ml - self.mr, h - self.mt - self.mb
        p = passo(max(vmax - min(vmin, 0), 1e-9))
        self.vmin = math.floor(min(vmin, 0) / p) * p
        self.vmax = math.ceil(vmax / p) * p
        self.p = p
        self.partes = []
        for v in self._ticks():
            y = self.y(v)
            self.partes.append(f'<line x1="{self.ml}" x2="{w - self.mr}" y1="{y:.1f}" y2="{y:.1f}" stroke="#E4E7EB"/>')
            self.partes.append(f'<text x="{self.ml - 6}" y="{y + 4:.1f}" text-anchor="end" font-size="11" fill="{TEXTO}">{fmt.format(v).replace(".", ",")}</text>')
        passo_x = self.iw / len(cats)
        for i, c in enumerate(cats):
            self.partes.append(f'<text x="{self.ml + passo_x * (i + .5):.1f}" y="{h - 8}" text-anchor="middle" font-size="11" fill="{TEXTO}">{c}</text>')

    def _ticks(self):
        v = self.vmin
        while v <= self.vmax + 1e-9:
            yield v; v += self.p

    def y(self, v): return self.mt + (1 - (v - self.vmin) / (self.vmax - self.vmin)) * self.ih
    def xc(self, i): return self.ml + self.iw / len(self.cats) * (i + .5)

    def barras(self, series, cores, rotulos=True, fmt="{:.0f}"):
        n = len(series); larg = self.iw / len(self.cats) * 0.72 / n
        for k, vals in enumerate(series):
            for i, v in enumerate(vals):
                cor = cores[k][i] if isinstance(cores[k], list) else cores[k]
                x = self.xc(i) - larg * n / 2 + larg * k
                y0, y1 = self.y(0), self.y(v)
                self.partes.append(f'<rect x="{x:.1f}" y="{min(y0, y1):.1f}" width="{larg:.1f}" height="{abs(y0 - y1):.1f}" fill="{cor}" rx="2"/>')
                if rotulos:
                    ty = y1 - 4 if v >= 0 else y1 + 12
                    self.partes.append(f'<text x="{x + larg / 2:.1f}" y="{ty:.1f}" text-anchor="middle" font-size="10" fill="#262626">{fmt.format(v).replace(".", ",")}</text>')

    def linha(self, vals, cor, rotulos=False, fmt="{:.0f}"):
        pts = [(self.xc(i), self.y(v)) for i, v in enumerate(vals)]
        self.partes.append(f'<polyline fill="none" stroke="{cor}" stroke-width="2.5" points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}"/>')
        for (x, y), v in zip(pts, vals):
            self.partes.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="{cor}"/>')
            if rotulos:
                self.partes.append(f'<text x="{x:.1f}" y="{y - 8:.1f}" text-anchor="middle" font-size="10" fill="#262626">{fmt.format(v).replace(".", ",")}</text>')

    def hline(self, v, cor, texto=""):
        y = self.y(v)
        self.partes.append(f'<line x1="{self.ml}" x2="{self.w - self.mr}" y1="{y:.1f}" y2="{y:.1f}" stroke="{cor}" stroke-dasharray="5 4"/>')
        if texto:
            self.partes.append(f'<text x="{self.w - self.mr - 2}" y="{y - 4:.1f}" text-anchor="end" font-size="10" fill="{cor}">{texto}</text>')

    def svg(self):
        return f'<svg viewBox="0 0 {self.w} {self.h}" width="100%">{"".join(self.partes)}</svg>'


def barras_horizontais(rotulos, valores, cores, w=560, h=300):
    ml, mr, mt = 84, 46, 6
    vmax = max(valores) * 1.05; alt = (h - mt - 6) / len(valores)
    out = []
    for i, (r, v, c) in enumerate(zip(rotulos, valores, cores)):
        y = mt + alt * i; larg = (w - ml - mr) * v / vmax
        out.append(f'<text x="{ml - 6}" y="{y + alt * .62:.1f}" text-anchor="end" font-size="11" fill="{TEXTO}">{r}</text>')
        out.append(f'<rect x="{ml}" y="{y + alt * .15:.1f}" width="{larg:.1f}" height="{alt * .7:.1f}" fill="{c}" rx="2"/>')
        out.append(f'<text x="{ml + larg + 4:.1f}" y="{y + alt * .62:.1f}" font-size="10" fill="#262626">{br(v)}</text>')
    return f'<svg viewBox="0 0 {w} {h}" width="100%">{"".join(out)}</svg>'


def legenda(itens):
    return '<div class="leg">' + "".join(f'<span><i style="background:{c}"></i>{t}</span>' for t, c in itens) + "</div>"


def card(rotulo, valor, sub="", cor="#262626"):
    return f'<div class="card"><div class="rot">{rotulo}</div><div class="val" style="color:{cor}">{valor}</div><div class="sub">{sub}</div></div>'


def painel(titulo, corpo):
    return f'<div class="painel"><h3>{titulo}</h3>{corpo}</div>'


# ---------------------------------------------------------------------
# Números do resumo (tudo calculado a partir das tabelas)
# ---------------------------------------------------------------------
prod, fat, per = bal["producao_t"].sum(), bal["faturamento_t"].sum(), bal["perdas_t"].sum()
perda_pct = per / fat
meses_alerta = [m for m in MESES if bal.loc[m, "alerta"] == "INVESTIGAR"]
t_por_t = fat / per
n = abc["classe"].value_counts()
parte_a = abc.loc[abc["classe"] == "A", "parte_pct"].sum()
dias_ult = cob["dias_papel_total"].iloc[-1]

# ---------------------------------------------------------------------
# Aba 1 · Visão geral
# ---------------------------------------------------------------------
frase1 = (f"De janeiro a agosto, <b>faturamos {br(fat)} t</b> (produzimos {br(prod)} t). "
          f"As perdas somam {br(per)} t, ou <b>{br(perda_pct * 100, 1)}% do faturado</b>. ")
frase1 += (f"Em <b>{', '.join(meses_alerta)}</b> a produção não fecha com o que saiu (diferença acima de 5%): vale investigar."
           if meses_alerta else "A produção fecha com o que saiu em todos os meses.")

g1 = Grafico(MESES, 0, bal[["producao_t", "faturamento_t"]].max().max() * 1.12, fmt="{:.0f}")
g1.barras([bal["producao_t"], bal["faturamento_t"]], [AZUL, VERDE], rotulos=False)
g2 = Grafico(MESES, bal["diferenca_pct"].min() * 100, bal["diferenca_pct"].max() * 100 * 1.15, fmt="{:.0f}%")
g2.barras([bal["diferenca_pct"] * 100], [[VERM if a == "INVESTIGAR" else VERDE for a in bal["alerta"]]], fmt="{:.1f}%")
g2.hline(5, TEXTO, "limite de alerta (5%)")
g3 = Grafico(MESES, 0, (bal["perda_sobre_faturado_pct"] * 100).max() * 1.25, fmt="{:.0f}%")
g3.linha(bal["perda_sobre_faturado_pct"] * 100, VERM, rotulos=True, fmt="{:.1f}%")
g4 = Grafico(MESES, 0, bal["perdas_t"].max() * 1.2, fmt="{:.0f}")
g4.barras([bal["perdas_t"]], [VERM], fmt="{:.0f}")

aba_geral = (
    f'<p class="frase">{frase1}</p><div class="cards">'
    + card("PRODUZIDO", f"{br(prod)} t", "jan a ago", AZUL)
    + card("FATURADO", f"{br(fat)} t", "jan a ago", VERDE)
    + card("PERDAS", f"{br(per)} t", f"{br(perda_pct * 100, 1)}% do faturado", VERM)
    + card("FATURADO POR TONELADA PERDIDA", f"{br(t_por_t, 1)} t", "quanto maior, melhor", "#262626")
    + card("MESES PARA INVESTIGAR", str(len(meses_alerta)), ", ".join(meses_alerta) or "nenhum", VERM if meses_alerta else VERDE)
    + '</div><div class="grade">'
    + painel("Produzido x faturado (t)", legenda([("Produzido", AZUL), ("Faturado", VERDE)]) + g1.svg())
    + painel("A diferença sem explicação passou de 5%?", g2.svg())
    + painel("Perda sobre o faturado (%)", g3.svg())
    + painel("Perdas por mês (t)", g4.svg())
    + "</div>")

# ---------------------------------------------------------------------
# Aba 2 · Clientes (curva ABC)
# ---------------------------------------------------------------------
top = abc.head(12)
maior = abc.iloc[0]
frase2 = (f"<b>{n.get('A', 0)} clientes (classe A)</b> respondem por {br(parte_a * 100, 0)}% da produção. "
          f"O maior, <b>{maior['cliente']}</b>, sozinho tem {br(maior['parte_pct'] * 100, 1)}%. "
          f"Ao todo são {len(abc)} clientes: {n.get('A', 0)} A, {n.get('B', 0)} B e {n.get('C', 0)} C.")
gx = Grafico([str(i) if i % 10 == 0 or i == 1 else "" for i in range(1, len(abc) + 1)], 0, 100, w=560, h=300, fmt="{:.0f}%")
pts = [(gx.xc(i), gx.y(v * 100)) for i, v in enumerate(abc["acumulado_pct"])]
gx.partes.append(f'<polyline fill="none" stroke="{AZUL}" stroke-width="2.5" points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}"/>')
gx.hline(80, TEXTO, "80%"); gx.hline(95, TEXTO, "95%")
linhas_top = "".join(
    f'<tr><td>{r.cliente}</td><td class="num">{br(r.producao_t, 1)}</td><td class="num">{br(r.parte_pct * 100, 1)}%</td><td class="num">{br(r.acumulado_pct * 100, 0)}%</td>'
    f'<td><span class="tag" style="background:{COR_CLASSE[r.classe]}">{r.classe}</span></td></tr>' for r in top.head(8).itertuples())
aba_abc = (
    f'<p class="frase">{frase2}</p><div class="cards">'
    + card("CLASSE A", f"{n.get('A', 0)} clientes", f"{br(parte_a * 100, 0)}% da produção", AZUL)
    + card("CLASSE B", f"{n.get('B', 0)} clientes", "até 95% acumulado", AMAR)
    + card("CLASSE C", f"{n.get('C', 0)} clientes", "o restante", TEXTO)
    + card("TOTAL DE CLIENTES", str(len(abc)), f"{br(abc['producao_t'].sum())} t produzidas", "#262626")
    + '</div><div class="grade">'
    + painel("12 maiores clientes (t produzidas)", legenda([("A", AZUL), ("B", AMAR), ("C", CINZA)])
             + barras_horizontais(list(top["cliente"]), list(top["producao_t"]), [COR_CLASSE[c] for c in top["classe"]]))
    + painel("Curva ABC: % acumulado da produção", gx.svg())
    + painel("Os 8 maiores em detalhe", f'<table><tr><th>Cliente</th><th>Toneladas</th><th>% do total</th><th>Acumulado</th><th>Classe</th></tr>{linhas_top}</table>')
    + painel("Como ler", '<p class="nota">A classe é definida pelo acumulado <b>antes</b> de somar o cliente: '
             'A até 80% da produção, B até 95% e C o restante. A base é o peso produzido por cliente (toneladas).</p>')
    + "</div>")

# ---------------------------------------------------------------------
# Aba 3 · Estoques
# ---------------------------------------------------------------------
frase3 = (f"O papel em estoque cobre cerca de <b>{br(dias_ult, 0)} dias</b> de produção (todos os locais). "
          f"Só o papel que já está dentro da fábrica cobre {br(cob['dias_papel_na_fabrica'].iloc[-1], 0)} dias. "
          f"O estoque de produto pronto fechou em {br(bal['estoque_produto_pronto_t'].iloc[-1], 0)} t em agosto.")
gd = Grafico(TODOS, 0, cob["dias_papel_total"].max() * 1.15, fmt="{:.0f}")
gd.linha(cob["dias_papel_total"], AZUL, rotulos=True, fmt="{:.0f}"); gd.linha(cob["dias_papel_na_fabrica"], VERDE, rotulos=True, fmt="{:.0f}")
ge = Grafico(MESES, 0, bal["estoque_produto_pronto_t"].max() * 1.15, fmt="{:.0f}")
ge.barras([bal["estoque_produto_pronto_t"]], [AZUL], fmt="{:.0f}")
gp = Grafico(TODOS, 0, cob["papel_total_t"].max() * 1.15, fmt="{:.0f}")
gp.barras([cob["papel_total_t"]], [AZUL], fmt="{:.0f}")
gen = Grafico(TODOS[1:], 0, cob["entradas_estimadas_t"].iloc[1:].max() * 1.15, fmt="{:.0f}")
gen.barras([cob["entradas_estimadas_t"].iloc[1:]], [VERDE], fmt="{:.0f}")
aba_estoque = (
    f'<p class="frase">{frase3}</p><div class="cards">'
    + card("PAPEL EM ESTOQUE", f"{br(cob['papel_total_t'].iloc[-1], 0)} t", "todos os locais, fim de out", AZUL)
    + card("DIAS DE PAPEL", f"{br(dias_ult, 0)} dias", "papel total ÷ produção diária", AZUL)
    + card("DIAS DE PAPEL NA FÁBRICA", f"{br(cob['dias_papel_na_fabrica'].iloc[-1], 0)} dias", "só G1 e G2", VERDE)
    + card("PRODUTO PRONTO", f"{br(bal['estoque_produto_pronto_t'].iloc[-1], 0)} t", "fim de agosto", "#262626")
    + '</div><div class="grade">'
    + painel("Quantos dias de produção o papel cobre?", legenda([("Papel total", AZUL), ("Dentro da fábrica", VERDE)]) + gd.svg())
    + painel("Estoque de produto pronto no fim do mês (t)", ge.svg())
    + painel("Estoque de papel, todos os locais (t)", gp.svg())
    + painel("Entradas de papel no mês (t) · ESTIMATIVA", gen.svg())
    + "</div>")

# ---------------------------------------------------------------------
# Página
# ---------------------------------------------------------------------
CSS = """
*{box-sizing:border-box} body{margin:0;font-family:Arial,Helvetica,sans-serif;background:#F4F6F8;color:#262626}
header{background:linear-gradient(90deg,#1F5FA8,#2E8B57);color:#fff;padding:18px 32px}
header h1{margin:0;font-size:22px} header p{margin:4px 0 0;font-size:12px;opacity:.9}
nav{display:flex;gap:6px;padding:12px 32px 0;background:#fff;border-bottom:1px solid #DDE1E6}
nav a{padding:9px 18px;font-size:13px;font-weight:bold;color:#595959;text-decoration:none;border-radius:8px 8px 0 0}
body[data-aba=geral] nav a[href='#geral'],body[data-aba=clientes] nav a[href='#clientes'],body[data-aba=estoques] nav a[href='#estoques']{background:#F4F6F8;color:#1F5FA8;border:1px solid #DDE1E6;border-bottom:none}
main{padding:16px 32px 20px} section{display:none}
body[data-aba=geral] #s-geral,body[data-aba=clientes] #s-clientes,body[data-aba=estoques] #s-estoques{display:block}
.frase{background:#fff;border-left:5px solid #1F5FA8;padding:12px 16px;font-size:15px;line-height:1.5;margin:0 0 14px;border-radius:4px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:14px}
.card{background:#fff;border-radius:8px;padding:12px 14px;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.rot{font-size:10px;font-weight:bold;color:#595959;letter-spacing:.3px}.val{font-size:26px;font-weight:bold;margin:4px 0}.sub{font-size:11px;color:#595959}
.grade{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.painel{background:#fff;border-radius:8px;padding:10px 14px 6px;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.painel h3{margin:0 0 6px;font-size:13px}
.leg{font-size:11px;color:#595959;margin-bottom:2px}.leg span{margin-right:14px}.leg i{display:inline-block;width:10px;height:10px;margin-right:5px;border-radius:2px}
table{width:100%;border-collapse:collapse;font-size:12px} th{text-align:left;color:#595959;font-size:11px;padding:4px 6px;border-bottom:1px solid #DDE1E6}
td{padding:6px;border-bottom:1px solid #EEF0F3}.num{text-align:right}.tag{color:#fff;font-weight:bold;padding:2px 9px;border-radius:10px;font-size:11px}
.nota{font-size:13px;line-height:1.5;color:#404040}
footer{padding:0 32px 14px;font-size:11px;color:#7F7F7F}
"""
JS = "document.body.dataset.aba=(location.hash||'#geral').slice(1);addEventListener('hashchange',()=>document.body.dataset.aba=(location.hash||'#geral').slice(1));"
pagina = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>Faturamento e perdas · Empresa Exemplo</title>
<style>{CSS}</style></head><body data-aba="geral">
<header><h1>Faturamento e perdas · Empresa Exemplo</h1><p>BI de Dados Industriais · só peso (toneladas) · DADOS FICTÍCIOS</p></header>
<nav><a href="#geral">Visão geral</a><a href="#clientes">Clientes (curva ABC)</a><a href="#estoques">Estoques</a></nav>
<main><section id="s-geral">{aba_geral}</section><section id="s-clientes">{aba_abc}</section><section id="s-estoques">{aba_estoque}</section></main>
<footer>Todos os números deste painel são fictícios e foram gerados por gerar_bi.py. A cobertura de papel e as entradas de papel são estimativas.</footer>
<script>{JS}</script></body></html>"""
(DASH / "index.html").write_text(pagina, encoding="utf-8")
print("dashboard/index.html gerado")

# ---------------------------------------------------------------------
# Prints (Chrome headless)
# ---------------------------------------------------------------------
chrome = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")
if chrome:
    for aba in ("geral", "clientes", "estoques"):
        saida = GRAFICOS / f"dashboard_{aba}.png"
        subprocess.run([chrome, "--headless", "--no-sandbox", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                        "--window-size=1280,1130", f"--screenshot={saida}", f"file://{DASH / 'index.html'}#{aba}"],
                       check=True, capture_output=True)
        print("print salvo:", saida.name)
else:
    print("Chrome não encontrado: abra dashboard/index.html no navegador para ver o painel.")
