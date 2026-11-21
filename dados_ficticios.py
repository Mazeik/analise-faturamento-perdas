"""
Dados FICTÍCIOS compartilhados pelo projeto inteiro (semente fixa: o resultado é sempre o mesmo).

A planilha (gerador_planilha.py), o BI (bi_dados_industriais/gerar_bi.py) e o modelo do Power BI
leem daqui, então os números de todas as partes conversam entre si.

Valores em kg, como chegam nos relatórios:
  fat["B"], fat["C"]  -> faturamento de Linha C e Linha D, jan/2025 a ago/2026 (20 meses)
  from_oct[...]       -> out/2025 a ago/2026 (11 meses): D = faturamento G2, E = faturamento G1,
                         F..L = perdas (Linha C, Linha D, G2 Formação, G2 Acabamento,
                         G1 Corte, G1 Formação, G1 Acabamento)
"""
import datetime as dt
import random

rng = random.Random(2025)


def ruido(base, spread, n):
    return [round(base * (1 + rng.uniform(-spread, spread))) for _ in range(n)]


fat = {"B": ruido(230000, 0.30, 20),      # Linha C: jan/2025 a ago/2026
       "C": ruido(36000, 0.25, 20)}       # Linha D
D_ = ruido(65000, 0.30, 11)               # G2: a partir de out/2025
E_ = ruido(410000, 0.12, 11)              # G1: a partir de out/2025
# perda(base, taxa, i, tend): perda em kg = faturamento x taxa, com ruído; "tend" faz a taxa subir (+) ou cair (-) mês a mês
perda = lambda base, taxa, i=0, tend=0.0: round(base * taxa * (1 + tend * i) * (1 + rng.uniform(-0.15, 0.15)), 1)
from_oct = {"D": D_, "E": E_,
            "F": [perda(fat["B"][9 + i], 0.006, i, -0.02) for i in range(11)],   # Linha C: melhorando devagar
            "G": [perda(fat["C"][9 + i], 0.085, i) for i in range(11)],          # Linha D: estável
            "H": [perda(D_[i], 0.020, i, 0.06) for i in range(11)],              # G2 Formação: piorando (pesa pouco)
            "I": [perda(D_[i], 0.012, i) for i in range(11)],                    # G2 Acabamento: estável
            "J": [perda(E_[i], 0.004, i) for i in range(11)],                    # G1 Corte: estável
            "K": [perda(E_[i], 0.050, i, -0.03) for i in range(11)],             # G1 Formação: melhorando (pesa muito)
            "L": [perda(E_[i], 0.028, i, 0.09) for i in range(11)]}              # G1 Acabamento: piorando (pesa muito)


# Meses de cada lista (primeiro dia do mês)
MESES_FAT = [dt.date(2025 + i // 12, i % 12 + 1, 1) for i in range(20)]          # jan/2025 .. ago/2026
MESES_OUT = MESES_FAT[9:]                                                         # out/2025 .. ago/2026
ETAPAS = {"F": ("Linha C", "Total"), "G": ("Linha D", "Total"),
          "H": ("Linha B G2", "Formação"), "I": ("Linha B G2", "Acabamento"),
          "J": ("Linha A G1", "Corte"), "K": ("Linha A G1", "Formação"),
          "L": ("Linha A G1", "Acabamento")}


def g1g2_2026():
    """G1+G2 de jan a ago/2026, em toneladas: faturamento e perdas (soma das 5 etapas). Usado pelo BI."""
    i0 = MESES_OUT.index(dt.date(2026, 1, 1))
    fat_t = [(from_oct["D"][i] + from_oct["E"][i]) / 1000 for i in range(i0, len(MESES_OUT))]
    per_t = [sum(from_oct[c][i] for c in "HIJKL") / 1000 for i in range(i0, len(MESES_OUT))]
    return fat_t, per_t

# Metas de perda (t por mês) definidas pela "diretoria" fictícia: setor, etapa, meta, vigente a partir de, observação
metas = [("Linha C", "Total", 1.5, dt.date(2025, 1, 1), "Meta original (fictícia)"),
         ("Linha C", "Total", 1.7, dt.date(2026, 6, 1), "Exemplo de meta que muda a partir de jun/2026"),
         ("Linha D", "Total", 3.3, dt.date(2025, 1, 1), ""),
         ("Linha B G2", "Formação", 1.4, dt.date(2025, 1, 1), ""),
         ("Linha B G2", "Acabamento", 0.85, dt.date(2025, 1, 1), ""),
         ("Linha A G1", "Corte", 1.8, dt.date(2025, 1, 1), ""),
         ("Linha A G1", "Formação", 20.0, dt.date(2025, 1, 1), ""),
         ("Linha A G1", "Acabamento", 12.0, dt.date(2025, 1, 1), "")]
