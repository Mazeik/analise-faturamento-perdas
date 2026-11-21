"""
BI de Dados Industriais (versão com DADOS 100% FICTÍCIOS).

Pergunta que o BI responde, em peso (toneladas), sem valores em R$:
  1. Balanço de peso: o que foi produzido bate com o que foi faturado + perdido + variação do estoque de produto pronto?
  2. Clientes: quais poucos clientes concentram a maior parte da produção? (curva ABC)
  3. Estoque de papel: quantos dias de produção o papel em estoque cobre?

Como rodar:
    pip install pandas matplotlib numpy
    python gerar_bi.py

Saídas (na mesma pasta do script):
    saida/    -> tabelas em CSV (ordens, balanço, curva ABC, papel)
    graficos/ -> imagens PNG usadas no README

O script faz tudo em 5 passos simples: gerar dados -> balanço -> curva ABC -> papel -> gráficos.
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                      # gera imagens sem abrir janela
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # para importar dados_ficticios.py da raiz
from dados_ficticios import g1g2_2026

PASTA = Path(__file__).parent
SAIDA, GRAFICOS = PASTA / "saida", PASTA / "graficos"
SAIDA.mkdir(exist_ok=True); GRAFICOS.mkdir(exist_ok=True)

rng = np.random.default_rng(7)             # semente fixa: o resultado é sempre o mesmo
LIMITE_ALERTA = 0.05                       # balanço com diferença acima de 5% gera alerta
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out"]
MES_COM_PROBLEMA = "jun"                   # mês fictício em que o balanço "não fecha" (para demonstrar o alerta)
COR = {"G1": "#1F5FA8", "G2": "#2E8B57", "A": "#1F5FA8", "B": "#E3A008", "C": "#B0B7C0",
       "ok": "#1B7F3B", "alerta": "#B3261E", "neutro": "#595959"}

plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titlesize": 12, "axes.titleweight": "bold", "figure.dpi": 130})

# ---------------------------------------------------------------------
# PASSO 1 · Dados fictícios
# ---------------------------------------------------------------------
# Faturamento e perdas de G1+G2 (jan a ago/2026) vêm de dados_ficticios.py, o MESMO arquivo que alimenta a planilha.
# Assim o BI e a planilha mostram os mesmos números para os mesmos meses.
fat_lanc, per_lanc = g1g2_2026()
LANCADOS = MESES[:len(fat_lanc)]                         # jan a ago (set/out ainda sem faturamento lançado)
faturamento = pd.Series(fat_lanc, index=LANCADOS)
perdas = pd.Series(per_lanc, index=LANCADOS)
media_fat = faturamento.mean()

# Estoque de produto pronto no fim de cada mês (t) e variação em relação ao mês anterior
estoque_pa = pd.Series(0.6 * media_fat * (1 + np.cumsum(rng.normal(0, 0.05, len(MESES)))), index=MESES)
variacao_pa = estoque_pa.diff()
variacao_pa.iloc[0] = estoque_pa.iloc[0] - 0.6 * media_fat   # estoque fictício do fim do ano anterior

# Produção (t): o que explica o faturado + perdas + variação do estoque, mais uma "diferença sem explicação".
# Na vida real essa diferença vem de erro de lançamento, ordem sem data, estoque contado errado etc.
erro = pd.Series(rng.normal(0.004, 0.012, len(MESES)), index=MESES)
erro[MES_COM_PROBLEMA] = 0.072
producao_alvo = (faturamento + perdas + variacao_pa.loc[LANCADOS]) / (1 - erro.loc[LANCADOS])
# set/out: sem faturamento lançado, só a produção (média dos meses lançados com variação)
for mes in MESES[len(LANCADOS):]:
    producao_alvo[mes] = producao_alvo.mean() * (1 + rng.normal(0, 0.05))

# Ordens de produção: 60 clientes com pesos diferentes (poucos grandes e muitos pequenos = curva ABC).
clientes = [f"Cliente {i:02d}" for i in range(1, 61)]
peso_cliente = rng.lognormal(mean=0, sigma=1.1, size=60)
peso_cliente = peso_cliente / peso_cliente.sum()
from dados_ficticios import from_oct
parte_g1 = sum(from_oct["E"]) / (sum(from_oct["E"]) + sum(from_oct["D"]))   # G1 x G2 na mesma proporção do faturamento

linhas = []
for mes in MESES:
    n_ordens = int(rng.integers(300, 380))
    peso = rng.lognormal(mean=6.4, sigma=0.6, size=n_ordens)
    peso = peso / peso.sum() * producao_alvo[mes] * 1000     # ajusta as ordens para somar a produção do mês
    linhas.append(pd.DataFrame({
        "mes": mes,
        "galpao": rng.choice(["G1", "G2"], size=n_ordens, p=[parte_g1, 1 - parte_g1]),
        "cliente": rng.choice(clientes, size=n_ordens, p=peso_cliente),
        "tipo_papel": rng.choice(["grosso", "fino", "kraft", "capa"], size=n_ordens, p=[0.55, 0.30, 0.12, 0.03]),
        "peso_kg": peso,
    }))
ordens = pd.concat(linhas, ignore_index=True)

# Produção do mês (t) = soma do peso das ordens
producao = ordens.groupby("mes")["peso_kg"].sum().reindex(MESES) / 1000

# Estoque de papel no fim do mês, por local (t), proporcional ao tamanho da produção. Dois locais abastecem o G1.
escala = producao.mean() / 245
papel = (pd.DataFrame({
    "G1": 40 + rng.normal(0, 5, len(MESES)),
    "G2": 18 + rng.normal(0, 3, len(MESES)),
    "Galpão de estoque": 150 + rng.normal(0, 20, len(MESES)),
    "Filial": 60 + rng.normal(0, 10, len(MESES)),
}, index=MESES) * escala).round(1)

# ---------------------------------------------------------------------
# PASSO 2 · Balanço de peso
# ---------------------------------------------------------------------
# Regra: Produção = Faturamento + Perdas + Variação do estoque de produto pronto
balanco = pd.DataFrame({
    "producao_t": producao, "faturamento_t": faturamento, "perdas_t": perdas,
    "variacao_estoque_t": variacao_pa, "estoque_produto_pronto_t": estoque_pa,
}).loc[LANCADOS]
balanco["explicado_t"] = balanco[["faturamento_t", "perdas_t", "variacao_estoque_t"]].sum(axis=1)
balanco["diferenca_t"] = balanco["producao_t"] - balanco["explicado_t"]
balanco["diferenca_pct"] = balanco["diferenca_t"] / balanco["producao_t"]
balanco["alerta"] = np.where(balanco["diferenca_pct"].abs() > LIMITE_ALERTA, "INVESTIGAR", "ok")
balanco["perda_sobre_faturado_pct"] = balanco["perdas_t"] / balanco["faturamento_t"]

# ---------------------------------------------------------------------
# PASSO 3 · Curva ABC de clientes (pela produção em t)
# ---------------------------------------------------------------------
abc = (ordens.groupby("cliente")["peso_kg"].sum().div(1000).sort_values(ascending=False)
       .rename("producao_t").reset_index())
abc["parte_pct"] = abc["producao_t"] / abc["producao_t"].sum()
abc["acumulado_pct"] = abc["parte_pct"].cumsum()
# Classe pelo acumulado ANTES de somar o cliente: A até 80%, B até 95%, C o resto
antes = abc["acumulado_pct"] - abc["parte_pct"]
abc["classe"] = np.select([antes < 0.80, antes < 0.95], ["A", "B"], default="C")

# ---------------------------------------------------------------------
# PASSO 4 · Cobertura de papel (em dias de produção)
# ---------------------------------------------------------------------
consumo_diario = producao / 30                       # simplificação: consumo de papel = peso produzido
fabrica = papel[["G1", "G2"]].sum(axis=1)            # papel que já está dentro da fábrica
cobertura = pd.DataFrame({
    "papel_total_t": papel.sum(axis=1),
    "papel_na_fabrica_t": fabrica,
    "dias_papel_total": papel.sum(axis=1) / consumo_diario,
    "dias_papel_na_fabrica": fabrica / consumo_diario,
})
# Entradas de papel = consumo + variação do estoque (ESTIMATIVA: não vem de nota fiscal)
cobertura["entradas_estimadas_t"] = producao + papel.sum(axis=1).diff().fillna(0)

# ---------------------------------------------------------------------
# Conferências (o BI só vale se os números fecham)
# ---------------------------------------------------------------------
assert abs(ordens["peso_kg"].sum() / 1000 - producao.sum()) < 1e-6, "produção não fecha com as ordens"
assert abs(abc["producao_t"].sum() - producao.sum()) < 1e-6, "curva ABC não fecha com a produção"
assert abs((balanco["diferenca_pct"] - erro.loc[LANCADOS]).abs().max()) < 1e-9, "balanço não fecha"
assert abs(faturamento.sum() - sum(fat_lanc)) < 1e-9, "faturamento diferente da planilha"

# ---------------------------------------------------------------------
# Salvar tabelas
# ---------------------------------------------------------------------
ordens.round({"peso_kg": 1}).to_csv(SAIDA / "ordens_ficticias.csv", index=False)
balanco.round(4).to_csv(SAIDA / "balanco_de_peso.csv")
abc.round(4).to_csv(SAIDA / "curva_abc_clientes.csv", index=False)
cobertura.round(2).to_csv(SAIDA / "cobertura_papel.csv")

# ---------------------------------------------------------------------
# PASSO 5 · Gráficos
# ---------------------------------------------------------------------
# 5.1 Balanço: barras (produção x faturado+perdido+estoque) e diferença % com linha de alerta
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 6), gridspec_kw={"height_ratios": [3, 2]}, sharex=True)
x = np.arange(len(LANCADOS)); larg = 0.38
ax1.bar(x - larg / 2, balanco["producao_t"], larg, color=COR["G1"], label="Produzido")
ax1.bar(x + larg / 2, balanco["faturamento_t"], larg, color=COR["G2"], label="Faturado")
ax1.bar(x + larg / 2, balanco["perdas_t"], larg, bottom=balanco["faturamento_t"], color=COR["alerta"], label="Perdas")
ax1.bar(x + larg / 2, balanco["variacao_estoque_t"].clip(lower=0), larg,
        bottom=balanco["faturamento_t"] + balanco["perdas_t"], color=COR["B"], label="Aumento do estoque")
ax1.set_ylabel("toneladas"); ax1.set_title("O que foi produzido bate com o que saiu?")
ax1.set_ylim(0, balanco[["producao_t", "explicado_t"]].max().max() * 1.22)   # espaço no topo para a legenda
ax1.legend(ncol=4, frameon=False, fontsize=8, loc="upper center")
cores = [COR["alerta"] if a == "INVESTIGAR" else COR["ok"] for a in balanco["alerta"]]
ax2.bar(x, balanco["diferenca_pct"] * 100, color=cores)
ax2.axhline(LIMITE_ALERTA * 100, color=COR["neutro"], ls="--", lw=1); ax2.axhline(-LIMITE_ALERTA * 100, color=COR["neutro"], ls="--", lw=1)
ax2.set_ylabel("diferença (%)"); ax2.set_xticks(x); ax2.set_xticklabels(LANCADOS)
ax2.text(len(x) - 0.5, LIMITE_ALERTA * 100 + 0.3, "limite de alerta (5%)", ha="right", fontsize=8, color=COR["neutro"])
for xi, v in zip(x, balanco["diferenca_pct"] * 100): ax2.text(xi, v + (0.3 if v >= 0 else -0.9), f"{v:.1f}%", ha="center", fontsize=8)
fig.tight_layout(); fig.savefig(GRAFICOS / "balanco_de_peso.png"); plt.close(fig)

# 5.2 Curva ABC + 12 maiores clientes
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.4))
top = abc.head(12).iloc[::-1]
ax1.barh(top["cliente"], top["producao_t"], color=[COR[c] for c in top["classe"]])
ax1.set_title("12 maiores clientes (cor = classe)"); ax1.set_xlabel("toneladas produzidas")
ax2.plot(np.arange(1, len(abc) + 1), abc["acumulado_pct"] * 100, color=COR["G1"], lw=2)
for lim in (80, 95): ax2.axhline(lim, color=COR["neutro"], ls="--", lw=0.8)
n = abc["classe"].value_counts()
ax2.set_title(f"Curva ABC: {n.get('A', 0)} clientes A, {n.get('B', 0)} B e {n.get('C', 0)} C")
ax2.set_xlabel("clientes (do maior para o menor)"); ax2.set_ylabel("% acumulado da produção")
fig.tight_layout(); fig.savefig(GRAFICOS / "curva_abc.png"); plt.close(fig)

# 5.3 Papel: dias de cobertura + estoque de produto pronto
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4))
ax1.plot(MESES, cobertura["dias_papel_total"], marker="o", color=COR["G1"], label="Papel total (todos os locais)")
ax1.plot(MESES, cobertura["dias_papel_na_fabrica"], marker="o", color=COR["G2"], label="Só o papel dentro da fábrica")
ax1.set_title("Quantos dias de produção o papel cobre?"); ax1.set_ylabel("dias"); ax1.legend(frameon=False, fontsize=8)
ax2.bar(MESES, estoque_pa, color=COR["G1"]); ax2.set_title("Estoque de produto pronto no fim do mês"); ax2.set_ylabel("toneladas")
fig.tight_layout(); fig.savefig(GRAFICOS / "papel_e_estoque.png"); plt.close(fig)

# ---------------------------------------------------------------------
# Resumo no terminal
# ---------------------------------------------------------------------
meses_alerta = balanco.index[balanco["alerta"] == "INVESTIGAR"].tolist()
print(f"{len(ordens)} ordens fictícias | produção {producao.sum():,.0f} t em {len(MESES)} meses")
print(f"Balanço: diferença média {balanco['diferenca_pct'].abs().mean():.1%} | meses para investigar: {meses_alerta}")
print(f"Curva ABC: {n.get('A', 0)} A / {n.get('B', 0)} B / {n.get('C', 0)} C | maior cliente = {abc.loc[0, 'parte_pct']:.1%} da produção")
print(f"Cobertura de papel (total): {cobertura['dias_papel_total'].min():.0f} a {cobertura['dias_papel_total'].max():.0f} dias")
print(f"Faturamento G1+G2 jan-ago = {faturamento.sum():,.1f} t (igual à planilha)")
print("Conferências OK: produção = soma das ordens; ABC fecha com a produção; balanço fecha; faturamento = planilha.")
