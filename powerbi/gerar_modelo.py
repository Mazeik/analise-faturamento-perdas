"""
Prepara as tabelas do modelo do Power BI (esquema estrela) a partir dos MESMOS dados fictícios
da planilha e do BI. Rode depois de bi_dados_industriais/gerar_bi.py.

    python powerbi/gerar_modelo.py

Saída: powerbi/dados/*.csv (separador ";" e vírgula decimal, para o Power BI em português abrir sem ajuste).

Esquema estrela = tabelas de FATOS (os números, uma linha por evento) ligadas a tabelas de DIMENSÃO
(as "réguas" para filtrar: mês, setor/etapa, cliente). Cada fato se liga às dimensões por uma chave.
"""
import datetime as dt
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
from dados_ficticios import fat, from_oct, metas, MESES_FAT, MESES_OUT, ETAPAS   # noqa: E402

SAIDA = Path(__file__).parent / "dados"
SAIDA.mkdir(exist_ok=True)
NOME_MES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

# ---------- dCalendario: um mês por linha (jan/2025 a dez/2026) ----------
meses = [dt.date(2025 + i // 12, i % 12 + 1, 1) for i in range(24)]
dCalendario = pd.DataFrame({
    "Mes": meses,
    "Ano": [m.year for m in meses],
    "NumMes": [m.month for m in meses],
    "MesAno": [f"{NOME_MES[m.month - 1]}/{str(m.year)[2:]}" for m in meses],   # rótulo do eixo (ordenar por IdxMes)
    "IdxMes": range(1, len(meses) + 1),   # 1, 2, 3...: facilita "mês anterior" e "últimos 3 meses" na DAX
})

# ---------- dEtapa: setor e etapa (a etapa "Total" = setor sem quebra por etapa) ----------
dEtapa = pd.DataFrame([(col, setor, etapa, setor if etapa == "Total" else f"{setor.split()[-1]} {etapa}")
                       for col, (setor, etapa) in ETAPAS.items()],
                      columns=["IdEtapa", "Setor", "Etapa", "NomeCurto"])

# ---------- dSetor: os 4 setores (liga o faturamento, que é por setor, às perdas, que são por etapa) ----------
dSetor = pd.DataFrame({"Setor": ["Linha C", "Linha D", "Linha B G2", "Linha A G1"],
                       "Grupo": ["Outros", "Outros", "Matriz G1/G2", "Matriz G1/G2"]})

# ---------- fFaturamento: mês x setor (t) ----------
linhas = [(m, "Linha C", v / 1000) for m, v in zip(MESES_FAT, fat["B"])]
linhas += [(m, "Linha D", v / 1000) for m, v in zip(MESES_FAT, fat["C"])]
linhas += [(m, "Linha B G2", v / 1000) for m, v in zip(MESES_OUT, from_oct["D"])]
linhas += [(m, "Linha A G1", v / 1000) for m, v in zip(MESES_OUT, from_oct["E"])]
fFaturamento = pd.DataFrame(linhas, columns=["Mes", "Setor", "Faturamento_t"])

# ---------- fPerdas: mês x etapa (t) ----------
fPerdas = pd.DataFrame([(m, col, v / 1000) for col in ETAPAS for m, v in zip(MESES_OUT, from_oct[col])],
                       columns=["Mes", "IdEtapa", "Perda_t"])

# ---------- fMeta: meta vigente em cada mês x etapa (t) ----------
# A meta pode mudar no meio do período; aqui já resolvemos qual vale em cada mês, para a DAX ficar simples.
lin = []
for col, (setor, etapa) in ETAPAS.items():
    vig = sorted((d, v) for s, e, v, d, _ in metas if s == setor and e == etapa)
    for m in MESES_OUT:
        validas = [v for d, v in vig if d <= m]
        if validas:
            lin.append((m, col, validas[-1]))
fMeta = pd.DataFrame(lin, columns=["Mes", "IdEtapa", "Meta_t"])

# ---------- fOrdens e dCliente: ordens de produção de 2026 (vêm do BI) ----------
ordens = pd.read_csv(RAIZ / "bi_dados_industriais" / "saida" / "ordens_ficticias.csv")
ordens["Mes"] = ordens["mes"].map({n: dt.date(2026, i + 1, 1) for i, n in enumerate(NOME_MES)})
fOrdens = pd.DataFrame({"Mes": ordens["Mes"], "Cliente": ordens["cliente"], "Galpao": ordens["galpao"],
                        "TipoPapel": ordens["tipo_papel"], "Peso_t": ordens["peso_kg"] / 1000})
dCliente = pd.DataFrame({"Cliente": sorted(fOrdens["Cliente"].unique())})

# ---------- fEstoque: estoque no fim do mês (t) ----------
bal = pd.read_csv(RAIZ / "bi_dados_industriais" / "saida" / "balanco_de_peso.csv", index_col=0)
cob = pd.read_csv(RAIZ / "bi_dados_industriais" / "saida" / "cobertura_papel.csv", index_col=0)
fEstoque = pd.DataFrame({
    "Mes": [dt.date(2026, NOME_MES.index(n) + 1, 1) for n in cob.index],
    "ProdutoPronto_t": bal["estoque_produto_pronto_t"].reindex(cob.index).values,
    "PapelTotal_t": cob["papel_total_t"].values,
    "PapelFabrica_t": cob["papel_na_fabrica_t"].values,
})

# ---------- Conferências ----------
assert abs(fFaturamento.loc[fFaturamento["Setor"].str.contains("G1|G2") & (pd.to_datetime(fFaturamento["Mes"]).dt.year == 2026),
                            "Faturamento_t"].sum() - bal["faturamento_t"].sum()) < 1e-6, "faturamento G1+G2 diferente do BI"
assert len(fMeta) == len(ETAPAS) * len(MESES_OUT), "falta meta em algum mês"

for nome, df in {"dCalendario": dCalendario, "dSetor": dSetor, "dEtapa": dEtapa, "dCliente": dCliente, "fFaturamento": fFaturamento,
                 "fPerdas": fPerdas, "fMeta": fMeta, "fOrdens": fOrdens, "fEstoque": fEstoque}.items():
    df.to_csv(SAIDA / f"{nome}.csv", sep=";", decimal=",", index=False, float_format="%.4f", encoding="utf-8-sig")
    print(f"{nome:13s} {len(df):5d} linhas")
print(f"Conferência: faturamento G1+G2 jan-ago/2026 = {bal['faturamento_t'].sum():,.1f} t (igual à planilha e ao BI)")
