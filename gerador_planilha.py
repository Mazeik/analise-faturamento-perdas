"""
Gerador da planilha "Faturamento e Perdas" (Python + openpyxl).

ATENÇÃO: todos os dados deste arquivo são FICTÍCIOS (gerados com semente fixa).
Servem apenas para demonstrar a estrutura, as fórmulas e os painéis.

Uso:  python gerador_planilha.py Faturamento_e_Perdas_exemplo.xlsx
A planilha gerada não depende de Python: tudo funciona com fórmulas do Excel (2019 ou mais novo).
"""
import datetime as dt
import sys
from dados_ficticios import fat, from_oct, metas   # dados fictícios compartilhados com o BI e o Power BI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont


# ---------------------------------------------------------------------
# Funções auxiliares de gráficos
# ---------------------------------------------------------------------
def _txt(sz=8, color="404040", bold=False):
    cp = CharacterProperties(sz=sz * 100, b=bold, solidFill=color, latin=DFont(typeface="Arial"))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def labels(series, fmt, pos, color="404040", bold=False, sz=8, bg=None):
    series.dLbls = DataLabelList(numFmt=fmt, dLblPos=pos, showVal=True, showLegendKey=False, showCatName=False,
                                 showSerName=False, showPercent=False, showBubbleSize=False, txPr=_txt(sz, color, bold))
    if bg:  # fundo sólido atrás do rótulo (legível sobre barras e linhas)
        series.dLbls.spPr = GraphicalProperties(solidFill=bg)


def last_label(series, fmt, pos, color, idx=11, sz=9):
    series.dLbls = DataLabelList(dLbl=[DataLabel(idx=idx, numFmt=fmt, dLblPos=pos, showVal=True, showLegendKey=False, showCatName=False,
                                                 showSerName=False, showPercent=False, showBubbleSize=False, txPr=_txt(sz, color, True))],
                                 showVal=False, showLegendKey=False, showCatName=False, showSerName=False, showPercent=False, showBubbleSize=False)


def place(ws, ch, c1, r1, c2, r2):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.position = "t"; ch.legend.overlay = False
    a = TwoCellAnchor(editAs="twoCell")
    a._from = AnchorMarker(col=CI(c1) - 1, colOff=0, row=r1 - 1, rowOff=0)
    a.to = AnchorMarker(col=CI(c2), colOff=0, row=r2, rowOff=0)
    ch.anchor = a
    ws.add_chart(ch)


OUT = sys.argv[1] if len(sys.argv) > 1 else "Faturamento_e_Perdas_exemplo.xlsx"
FONT = "Arial"
EMPRESA = "Empresa Exemplo"
R0, R1 = 6, 53                      # linhas de meses (jan/2025 .. dez/2028)
MONTHS = [dt.date(2025 + i // 12, i % 12 + 1, 1) for i in range(48)]
FIN_OFF = 21                        # linha financeira = linha do mês + 21
CALC = "'Cálculos'"

# ---------- cores ----------
C_IN = "FFF2CC"; C_GRAY = "D9D9D9"; C_HEAD = "404040"; C_SOFT = "F3F3F3"
GOOD_F, GOOD_B = "1B7F3B", "E3F2E6"
BAD_F, BAD_B = "B3261E", "FBE3E1"
AMB_F, AMB_B, AMB_BAR = "A15C00", "FCE4B6", "E3A008"   # âmbar = atenção (no limite da meta)
# Lançamentos: blocos com cores próprias (faturamento azul, perdas laranja)
FAT_H, FAT_S, FAT_IN = "1F4E79", "BDD7EE", "DDEBF7"
PER_H, PER_S, PER_IN = "843C0C", "F8CBAD", "FCE4D6"
TOLMETA = "TolMeta"
# toneladas: sem casa decimal a partir de 10 t; 1 casa abaixo disso
TFMT = '[>=10]#,##0" t";#,##0.0" t"'


def tfx(x): return f"FIXED({x},IF(ABS({x})>=10,0,1))"


SUBS = [  # nome, chave, cor, col faturamento (Lanc), etapas [(etapa, col perda Lanc)]
    ("Linha C", "LC", "D9822B", "B", [("Total", "F")]),
    ("Linha D", "LD", "1A9BA6", "C", [("Total", "G")]),
    ("Linha B G2", "G2", "2E8B57", "D", [("Formação", "H"), ("Acabamento", "I")]),
    ("Linha A G1", "G1", "1F5FA8", "E", [("Corte", "J"), ("Formação", "K"), ("Acabamento", "L")]),
]
SHORT = {"LC": "Linha C", "LD": "Linha D", "G2": "G2", "G1": "G1"}
FIN_FAT = {"LC": "B", "LD": "C", "G2": "D", "G1": "E"}
FIN_CUS = {"LC": "F", "LD": "G", "G2": "H", "G1": "I"}

thin = Side(style="thin", color="BFBFBF")


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c): return PatternFill("solid", start_color=c, end_color=c)


def box(ws, rng):
    rows = list(ws[rng])
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            c.border = Border(left=thin if j == 0 else c.border.left, right=thin if j == len(row) - 1 else c.border.right,
                              top=thin if i == 0 else c.border.top, bottom=thin if i == len(rows) - 1 else c.border.bottom)


def mtxt(ref, short=False):
    m = 'CHOOSE(MONTH(%s),"jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez")' % ref
    return f'{m}&"/"&RIGHT(YEAR({ref}),2)' if short else f'{m}&"/"&YEAR({ref})'


def status_cf(ws, rng, ref):
    for val, fc, bc in (("Melhorou", GOOD_F, GOOD_B), ("Piorou", BAD_F, BAD_B)):   # sem cinza: qualquer variação conta
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="{val}"'], font=Font(name=FONT, color=fc, bold=True), fill=fill(bc)))


def meta_cf(ws, rng, ref):
    for val, fc, bc in (("Melhorou", GOOD_F, GOOD_B), ("Piorou", BAD_F, BAD_B), ("Manteve", AMB_F, AMB_B)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{ref}="{val}"'], font=Font(name=FONT, color=fc, bold=True), fill=fill(bc)))


def acima_t(p, m):  # toneladas acima da meta (em kg quando < 0,1 t)
    x = f"({p}-{m})"
    return f'IF({x}<0.1,FIXED({x}*1000,IF({x}*1000<1,1,0))&" kg",FIXED({x},1)&" t")&" acima"'


def abaixo(p, m):  # toneladas abaixo da meta (em kg quando < 0,1 t)
    x = f"({m}-{p})"
    return f'IF({x}<0.1,FIXED({x}*1000,IF({x}*1000<1,1,0))&" kg",FIXED({x},1)&" t")&" abaixo"'


def setup_print(ws, area, landscape=True):
    ws.print_area = area
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4; ws.page_margins.bottom = 0.7; ws.page_margins.footer = 0.3
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = f"{EMPRESA} · Faturamento e Perdas · &A"
    ws.oddFooter.left.size = 8; ws.oddFooter.left.font = FONT
    ws.oddFooter.right.text = "Página &P"; ws.oddFooter.right.size = 8; ws.oddFooter.right.font = FONT


wb = Workbook()
wsI = wb.active; wsI.title = "Instruções"
wsL = wb.create_sheet("Lançamentos")
wsT = wb.create_sheet("Painel Consolidado")
wsO = wb.create_sheet("Onde focar")
panels = {s[1]: wb.create_sheet(f"Painel {s[0]}") for s in SUBS}
wsR = wb.create_sheet("Resumo")
wsM = wb.create_sheet("Metas")
wsF = wb.create_sheet("Financeiro")
wsC = wb.create_sheet("Cálculos")
P1 = wsT.title  # aba aberta ao abrir o arquivo

# =====================================================================
# METAS (valores fictícios)
# =====================================================================
ws = wsM
ws["A1"] = "Metas de perda (t por mês)"; ws["A1"].font = f(14, True)
ws["A2"] = "Metas definidas pela diretoria. Para mudar uma meta, inclua uma nova linha com a data a partir da qual ela vale (o histórico continua usando a meta antiga)."
ws["A2"].font = f(9, color="595959")
ws["A3"] = "Faixa de atenção da meta"; ws["A3"].font = f(10, True)
ws["C3"] = 0.01; ws["C3"].number_format = "0.0%"; ws["C3"].fill = fill(C_IN); ws["C3"].font = f(10)
ws["D3"] = ('="A meta é o limite máximo de perda. Âmbar (atenção) = a perda ainda está abaixo da meta, mas a menos de "&FIXED(C3*100,0)&"% do limite. Ex.: meta de 10 t → âmbar '
            'com perda entre "&FIXED(10*(1-C3),1)&" t e 10 t. Acima de 10 t, mesmo que pouco = vermelho."')
ws["D3"].font = f(9, color="404040")
hdr = ["Setor", "Etapa", "Meta (t/mês)", "Vigente a partir de", "Observação"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(5, i, h); c.font = f(10, True, "FFFFFF"); c.fill = fill(C_HEAD); c.alignment = Alignment(horizontal="center")
for r in range(6, 31):
    for cidx in range(1, 6):
        c = ws.cell(r, cidx); c.fill = fill(C_IN); c.font = f(10); c.border = Border(bottom=Side(style="hair", color="BFBFBF"))
    ws.cell(r, 3).number_format = "#,##0.0000"; ws.cell(r, 4).number_format = "dd/mm/yyyy"
for i, m in enumerate(metas):
    for j, v in enumerate(m): ws.cell(6 + i, 1 + j, v)
for col, w in zip("ABCDE", (22, 16, 14, 20, 34)): ws.column_dimensions[col].width = w
dvS = DataValidation(type="list", formula1='"Linha C,Linha D,Linha B G2,Linha A G1"', allow_blank=True,
                     error="Escolha um setor da lista.", errorTitle="Setor inválido", showErrorMessage=True)
dvE = DataValidation(type="list", formula1='"Total,Corte,Formação,Acabamento"', allow_blank=True,
                     error="Escolha uma etapa da lista.", errorTitle="Etapa inválida", showErrorMessage=True)
dvN = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True, error="Digite um número maior que zero.", errorTitle="Valor inválido", showErrorMessage=True)
dvD = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, error="Digite uma data (ex.: 01/06/2026).", errorTitle="Data inválida", showErrorMessage=True)
dvT = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.5", error="Digite um percentual entre 0% e 50%.", showErrorMessage=True)
for dv, rng in ((dvS, "A6:A30"), (dvE, "B6:B30"), (dvN, "C6:C30"), (dvD, "D6:D30"), (dvT, "C3")):
    ws.add_data_validation(dv); dv.add(rng)
ws.freeze_panes = "A6"
wb.defined_names[TOLMETA] = DefinedName(TOLMETA, attr_text="'Metas'!$C$3")
MR = lambda col: f"'Metas'!${col}$6:${col}$30"

# =====================================================================
# LANÇAMENTOS (dois blocos separados: faturamento em azul, perdas em laranja)
# =====================================================================
ws = wsL
ws["A1"] = "Lançamentos mensais (kg)"; ws["A1"].font = f(14, True)
ws["A3"] = "Legenda:"; ws["A3"].font = f(9, True)
ws["B3"] = "Faturamento"; ws["B3"].fill = fill(FAT_IN); ws["B3"].font = f(9, True, FAT_H)
ws["C3"] = "Perdas"; ws["C3"].fill = fill(PER_IN); ws["C3"].font = f(9, True, PER_H)
ws["D3"] = "Sem dado separado"; ws["D3"].fill = fill(C_GRAY); ws["D3"].font = f(9)
for a in ("B3", "C3"): ws[a].alignment = Alignment(horizontal="center")
SEP = Side(style="medium", color="404040")
ws.merge_cells("B4:E4"); ws["B4"] = "FATURAMENTO (kg)"
ws.merge_cells("F4:L4"); ws["F4"] = "PERDAS (kg)"
for a, cr in (("B4", FAT_H), ("F4", PER_H)):
    ws[a].font = f(12, True, "FFFFFF"); ws[a].alignment = Alignment(horizontal="center", vertical="center")
for col in "BCDE": ws[f"{col}4"].fill = fill(FAT_H)
for col in "FGHIJKL": ws[f"{col}4"].fill = fill(PER_H)
ws.row_dimensions[4].height = 22
heads = ["Mês", "Linha C", "Linha D", "Linha B G2", "Linha A G1", "Linha C", "Linha D",
         "G2 Formação", "G2 Acabamento", "G1 Corte", "G1 Formação", "G1 Acabamento", "Observação"]
for i, h in enumerate(heads, 1):
    c = ws.cell(5, i, h); c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    col = L(i)
    if col in "BCDE": c.font = f(10, True, FAT_H); c.fill = fill(FAT_S)
    elif col in "FGHIJKL": c.font = f(10, True, PER_H); c.fill = fill(PER_S)
    else: c.font = f(10, True); c.fill = fill("D9D9D9")
ws.row_dimensions[5].height = 30

# ---------- DADOS FICTÍCIOS: vêm de dados_ficticios.py (compartilhado com o BI e o Power BI) ----------

hair = Side(style="hair", color="BFBFBF")
for i, m in enumerate(MONTHS):
    r = R0 + i
    c = ws.cell(r, 1, m); c.number_format = "mmm/yyyy"; c.font = f(10, True); c.fill = fill("F2F2F2")
    for col in "BCDEFGHIJKLM":
        cc = ws[f"{col}{r}"]; cc.font = f(10); cc.number_format = "#,##0.00" if col != "M" else "@"
        gray = col in "DEHIJKL" and i < 9
        cc.fill = fill(C_GRAY if gray else (FAT_IN if col in "BCDE" else PER_IN if col in "FGHIJKL" else "FFFFFF"))
        cc.border = Border(bottom=hair, right=hair)
        if not gray: cc.protection = cc.protection.copy(locked=False)
# linha divisória forte entre os blocos
for r in range(4, R1 + 1):
    for col in ("B", "F", "M"):
        c = ws[f"{col}{r}"]; c.border = Border(left=SEP, right=c.border.right, top=c.border.top, bottom=c.border.bottom)
for col, vals in fat.items():
    for i, v in enumerate(vals): ws[f"{col}{R0 + i}"] = v
for col, vals in from_oct.items():
    for i, v in enumerate(vals): ws[f"{col}{R0 + 9 + i}"] = v
ws["M6"] = "DADOS FICTÍCIOS de exemplo (jan/2025 em diante)"
ws[f"M{R0+9}"] = "Início do acompanhamento separado de G1 e G2"
ws.column_dimensions["A"].width = 11
for col in "BCDEFGHIJKL": ws.column_dimensions[col].width = 14
ws.column_dimensions["M"].width = 44
dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    error="Digite apenas números (em kg), maiores ou iguais a zero.", errorTitle="Valor inválido", showErrorMessage=True,
                    prompt="Valor em kg", showInputMessage=False)
ws.add_data_validation(dv); dv.add(f"B{R0}:L{R1}")
ws.freeze_panes = "B6"
setup_print(ws, f"A1:M{R1}")
ws.page_setup.fitToHeight = 0

# =====================================================================
# FINANCEIRO (lançamentos em R$; análise no topo) — opcional, sem dados
# =====================================================================
ws = wsF
ws["B24"] = "Lançamentos financeiros (opcional · R$)"; ws["B24"].font = f(13, True)
ws["E24"] = "Preencher só se a diretoria quiser a análise financeira. Custo do material = R$ por kg de material perdido no mês."
ws["E24"].font = f(9, color="595959")
ws.merge_cells("B25:E25"); ws["B25"] = "Faturamento (R$)"
ws.merge_cells("F25:I25"); ws["F25"] = "Custo do material perdido (R$/kg)"
for a in ("B25", "F25"):
    ws[a].font = f(10, True, "FFFFFF"); ws[a].fill = fill(C_HEAD); ws[a].alignment = Alignment(horizontal="center")
for i, h in enumerate(["Mês", "Linha C", "Linha D", "Linha B G2", "Linha A G1"] + ["Linha C", "Linha D", "Linha B G2", "Linha A G1"], 1):
    c = ws.cell(26, i, h); c.font = f(10, True); c.fill = fill("D9D9D9"); c.alignment = Alignment(horizontal="center", wrap_text=True)
for i in range(48):
    r = R0 + i + FIN_OFF
    c = ws.cell(r, 1, f"='Lançamentos'!A{R0+i}"); c.number_format = "mmm/yyyy"; c.font = f(10, True); c.fill = fill("F2F2F2")
    for col in "BCDEFGHI":
        cc = ws[f"{col}{r}"]; cc.fill = fill(C_IN); cc.font = f(10)
        cc.number_format = '"R$" #,##0.00' if col in "BCDE" else '"R$" #,##0.000'
        cc.border = Border(bottom=Side(style="hair", color="BFBFBF"))
        cc.protection = cc.protection.copy(locked=False)
dvF = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                     error="Digite apenas números (R$), maiores ou iguais a zero.", errorTitle="Valor inválido", showErrorMessage=True)
ws.add_data_validation(dvF); dvF.add(f"B{R0+FIN_OFF}:I{R1+FIN_OFF}")

# =====================================================================
# CÁLCULOS (base de cálculo — uma linha por mês)
# =====================================================================
ws = wsC
ws["A1"] = "Último mês lançado"; ws["B1"] = f"=MAX(C{R0}:C{R1})"
ws["A2"] = "Mês padrão (último lançado)"; ws["B2"] = "=B1"
ws["A3"] = "Posição do mês"; ws["B3"] = f"=MATCH(B2,A{R0}:A{R1},0)"
ws["C2"] = "Texto"; ws["D2"] = f"=INDEX(D{R0}:D{R1},B3)"
ws["C3"] = "Mês anterior"; ws["D3"] = f"=IF(B3>1,INDEX(D{R0}:D{R1},B3-1),\"\")"
ws["C4"] = "Mesmo mês ano ant."; ws["D4"] = f"=IF(B3>12,INDEX(D{R0}:D{R1},B3-12),\"\")"
for a in ("B1", "B2"): ws[a].number_format = "mmm/yyyy"
for nm, ref in (("UltimoMes", "$B$1"), ("MesSel", "$B$2"), ("SelIdx", "$B$3"), ("MesTxt", "$D$2"), ("MesAntTxt", "$D$3"), ("MesAnoAntTxt", "$D$4")):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"{CALC}!{ref}")

cols = {}; nxt = [6]


def newcol(key, header):
    c = L(nxt[0]); nxt[0] += 1; cols[key] = c
    ws[f"{c}5"] = header; return c


for k, h in (("MES", "Mês"), ("ANO", "Ano"), ("LANC", "Lançado"), ("MTX", "Mês (texto)"), ("MCT", "Mês (curto)")):
    c = L(len(cols) + 1); cols[k] = c; ws[f"{c}5"] = h
nxt[0] = 6
formulas = {}  # key -> function(r) -> formula


def addf(key, header, fn):
    newcol(key, header); formulas[key] = fn


formulas["MES"] = lambda r: f"='Lançamentos'!A{r}"
formulas["ANO"] = lambda r: f"=YEAR(A{r})"
formulas["LANC"] = lambda r: f"=IF(COUNT('Lançamentos'!B{r}:L{r})>0,A{r},\"\")"
formulas["MTX"] = lambda r: "=" + mtxt(f"A{r}")
formulas["MCT"] = lambda r: "=" + mtxt(f"A{r}", True)


def st_rel(val, up_is_bad=True):
    bad, good = ('"Piorou"', '"Melhorou"') if up_is_bad else ('"Melhorou"', '"Piorou"')
    return f'IF({val}=0,"Igual",IF({val}>0,{bad},{good}))'


def st_meta(pm):
    return f'IF({pm}>1,"Piorou",IF({pm}>=1-{TOLMETA},"Manteve","Melhorou"))'


def metafx(sub, et, r):
    mx = f'_xlfn.MAXIFS({MR("D")},{MR("A")},"{sub}",{MR("B")},"{et}",{MR("D")},"<="&$A{r})'
    return f'=IF({mx}=0,"",SUMIFS({MR("C")},{MR("A")},"{sub}",{MR("B")},"{et}",{MR("D")},{mx}))'


for sub, key, color, fcol, etapas in SUBS:
    ek = []
    for et, pcol in etapas:
        k = f"{key}_{et}"; ek.append(k)
        addf(f"{k}_PT", f"{key} {et} perda (t)", lambda r, pcol=pcol: f"=IF('Lançamentos'!{pcol}{r}=\"\",\"\",'Lançamentos'!{pcol}{r}/1000)")
        addf(f"{k}_MT", f"{key} {et} meta (t)", lambda r, sub=sub, et=et: metafx(sub, et, r))
        addf(f"{k}_PM", f"{key} {et} % da meta", lambda r, k=k: f'=IF(OR({cols[k+"_PT"]}{r}="",{cols[k+"_MT"]}{r}=""),"",{cols[k+"_PT"]}{r}/{cols[k+"_MT"]}{r})')
        addf(f"{k}_SM", f"{key} {et} status meta", lambda r, k=k: f'=IF({cols[k+"_PM"]}{r}="","",{st_meta(cols[k+"_PM"]+str(r))})')
        addf(f"{k}_VT", f"{key} {et} var t vs ant", lambda r, k=k: '=""' if r == R0 else f'=IF(OR({cols[k+"_PT"]}{r}="",{cols[k+"_PT"]}{r-1}=""),"",{cols[k+"_PT"]}{r}-{cols[k+"_PT"]}{r-1})')
        addf(f"{k}_VP", f"{key} {et} var % vs ant", lambda r, k=k: '=""' if r == R0 else f'=IF({cols[k+"_VT"]}{r}="","",IF({cols[k+"_PT"]}{r-1}=0,"",{cols[k+"_VT"]}{r}/{cols[k+"_PT"]}{r-1}))')
        addf(f"{k}_SV", f"{key} {et} status vs ant", lambda r, k=k: f'=IF({cols[k+"_VP"]}{r}="","",{st_rel(cols[k+"_VP"]+str(r))})')
    addf(f"{key}_FT", f"{key} faturamento (t)", lambda r, fcol=fcol: f"=IF('Lançamentos'!{fcol}{r}=\"\",\"\",'Lançamentos'!{fcol}{r}/1000)")

    def ps(r, ek=ek):
        refs = ",".join(f"{cols[k+'_PT']}{r}" for k in ek)
        return f'=IF(COUNT({refs})=0,"",SUM({refs}))'
    addf(f"{key}_PS", f"{key} perda (t)", ps)
    addf(f"{key}_PP", f"{key} % perda", lambda r, key=key: f'=IF(OR({cols[key+"_FT"]}{r}="",{cols[key+"_PS"]}{r}=""),"",IF({cols[key+"_FT"]}{r}=0,"",{cols[key+"_PS"]}{r}/{cols[key+"_FT"]}{r}))')
    addf(f"{key}_PA", f"{key} % perda mês ant", lambda r, key=key: '=""' if r == R0 else f'={cols[key+"_PP"]}{r-1}')
    addf(f"{key}_DP", f"{key} var p.p.", lambda r, key=key: f'=IF(OR({cols[key+"_PP"]}{r}="",{cols[key+"_PA"]}{r}=""),"",{cols[key+"_PP"]}{r}-{cols[key+"_PA"]}{r})')
    addf(f"{key}_SA", f"{key} status vs ant", lambda r, key=key: f'=IF({cols[key+"_DP"]}{r}="","",IF({cols[key+"_PA"]}{r}=0,"",{st_rel(cols[key+"_PP"]+str(r)+"/"+cols[key+"_PA"]+str(r)+"-1")}))')

    def ms(r, ek=ek):
        refs = ",".join(f"{cols[k+'_MT']}{r}" for k in ek)
        return f'=IF(COUNT({refs})=0,"",SUM({refs}))'
    addf(f"{key}_MS", f"{key} meta (t)", ms)
    addf(f"{key}_PMS", f"{key} % da meta", lambda r, key=key: f'=IF(OR({cols[key+"_PS"]}{r}="",{cols[key+"_MS"]}{r}=""),"",{cols[key+"_PS"]}{r}/{cols[key+"_MS"]}{r})')
    addf(f"{key}_SMS", f"{key} status meta", lambda r, key=key: f'=IF({cols[key+"_PMS"]}{r}="","",{st_meta(cols[key+"_PMS"]+str(r))})')
    addf(f"{key}_PSOK", f"{key} perda dentro da meta", lambda r, key=key: f'=IF(OR({cols[key+"_SMS"]}{r}="Piorou",{cols[key+"_SMS"]}{r}="Manteve"),"",{cols[key+"_PS"]}{r})')
    addf(f"{key}_PSAT", f"{key} perda no limite da meta", lambda r, key=key: f'=IF({cols[key+"_SMS"]}{r}="Manteve",{cols[key+"_PS"]}{r},"")')
    addf(f"{key}_PSOV", f"{key} perda acima da meta", lambda r, key=key: f'=IF({cols[key+"_SMS"]}{r}="Piorou",{cols[key+"_PS"]}{r},"")')
    addf(f"{key}_VS", f"{key} perda var t vs ant", lambda r, key=key: '=""' if r == R0 else f'=IF(OR({cols[key+"_PS"]}{r}="",{cols[key+"_PS"]}{r-1}=""),"",{cols[key+"_PS"]}{r}-{cols[key+"_PS"]}{r-1})')
    addf(f"{key}_SVS", f"{key} perda status vs ant", lambda r, key=key: '=""' if r == R0 else f'=IF({cols[key+"_VS"]}{r}="","",IF({cols[key+"_PS"]}{r-1}=0,"",{st_rel(cols[key+"_VS"]+str(r)+"/"+cols[key+"_PS"]+str(r-1))}))')
    addf(f"{key}_FV", f"{key} fat var % vs ant", lambda r, key=key: '=""' if r == R0 else f'=IF(OR({cols[key+"_FT"]}{r}="",{cols[key+"_FT"]}{r-1}=""),"",IF({cols[key+"_FT"]}{r-1}=0,"",{cols[key+"_FT"]}{r}/{cols[key+"_FT"]}{r-1}-1))')
    addf(f"{key}_SF", f"{key} fat status", lambda r, key=key: f'=IF({cols[key+"_FV"]}{r}="","",{st_rel(cols[key+"_FV"]+str(r), up_is_bad=False)})')
    addf(f"{key}_PY", f"{key} % perda ano ant", lambda r, key=key: '=""' if r < R0 + 12 else f'={cols[key+"_PP"]}{r-12}')
    addf(f"{key}_DY", f"{key} var p.p. vs ano ant", lambda r, key=key: f'=IF(OR({cols[key+"_PP"]}{r}="",{cols[key+"_PY"]}{r}=""),"",{cols[key+"_PP"]}{r}-{cols[key+"_PY"]}{r})')
    addf(f"{key}_SY", f"{key} status vs ano ant", lambda r, key=key: f'=IF({cols[key+"_DY"]}{r}="","",IF({cols[key+"_PY"]}{r}=0,"",{st_rel(cols[key+"_PP"]+str(r)+"/"+cols[key+"_PY"]+str(r)+"-1")}))')
    addf(f"{key}_SFY", f"{key} fat status vs ano ant", lambda r, key=key: f'=IF({cols[key+"_FY"]}{r}="","",{st_rel(cols[key+"_FY"]+str(r), up_is_bad=False)})')
    addf(f"{key}_FY", f"{key} fat var % vs ano ant", lambda r, key=key: '=""' if r < R0 + 12 else f'=IF(OR({cols[key+"_FT"]}{r}="",{cols[key+"_FT"]}{r-12}=""),"",IF({cols[key+"_FT"]}{r-12}=0,"",{cols[key+"_FT"]}{r}/{cols[key+"_FT"]}{r-12}-1))')
    # financeiro
    addf(f"{key}_FR", f"{key} faturamento R$", lambda r, key=key: f"=IF(Financeiro!{FIN_FAT[key]}{r+FIN_OFF}=\"\",\"\",Financeiro!{FIN_FAT[key]}{r+FIN_OFF})")
    addf(f"{key}_CK", f"{key} custo R$/kg", lambda r, key=key: f"=IF(Financeiro!{FIN_CUS[key]}{r+FIN_OFF}=\"\",\"\",Financeiro!{FIN_CUS[key]}{r+FIN_OFF})")


def fin_block(key, member_keys=None):
    if member_keys:  # total
        def sum4(suffix, r):
            refs = ",".join(f"{cols[k+suffix]}{r}" for k in member_keys)
            return f'IF(COUNT({refs})<{len(member_keys)},"",SUM({refs}))'
        addf(f"{key}_FT", "TOTAL faturamento (t)", lambda r: "=" + sum4("_FT", r))
        addf(f"{key}_PS", "TOTAL perda (t)", lambda r: "=" + sum4("_PS", r))
        addf(f"{key}_PP", "TOTAL % perda", lambda r: f'=IF(OR({cols[key+"_FT"]}{r}="",{cols[key+"_PS"]}{r}=""),"",IF({cols[key+"_FT"]}{r}=0,"",{cols[key+"_PS"]}{r}/{cols[key+"_FT"]}{r}))')
        addf(f"{key}_PA", "TOTAL % perda mês ant", lambda r: '=""' if r == R0 else f'={cols[key+"_PP"]}{r-1}')
        addf(f"{key}_DP", "TOTAL var p.p.", lambda r: f'=IF(OR({cols[key+"_PP"]}{r}="",{cols[key+"_PA"]}{r}=""),"",{cols[key+"_PP"]}{r}-{cols[key+"_PA"]}{r})')
        addf(f"{key}_SA", "TOTAL status vs ant", lambda r: f'=IF({cols[key+"_DP"]}{r}="","",IF({cols[key+"_PA"]}{r}=0,"",{st_rel(cols[key+"_PP"]+str(r)+"/"+cols[key+"_PA"]+str(r)+"-1")}))')
        addf(f"{key}_MS", "TOTAL meta (t)", lambda r: "=" + sum4("_MS", r))
        addf(f"{key}_PMS", "TOTAL % da meta", lambda r: f'=IF(OR({cols[key+"_PS"]}{r}="",{cols[key+"_MS"]}{r}=""),"",{cols[key+"_PS"]}{r}/{cols[key+"_MS"]}{r})')
        addf(f"{key}_SMS", "TOTAL status meta", lambda r: f'=IF({cols[key+"_PMS"]}{r}="","",{st_meta(cols[key+"_PMS"]+str(r))})')
        addf(f"{key}_VS", "TOTAL perda var t", lambda r: '=""' if r == R0 else f'=IF(OR({cols[key+"_PS"]}{r}="",{cols[key+"_PS"]}{r-1}=""),"",{cols[key+"_PS"]}{r}-{cols[key+"_PS"]}{r-1})')
        addf(f"{key}_FV", "TOTAL fat var %", lambda r: '=""' if r == R0 else f'=IF(OR({cols[key+"_FT"]}{r}="",{cols[key+"_FT"]}{r-1}=""),"",IF({cols[key+"_FT"]}{r-1}=0,"",{cols[key+"_FT"]}{r}/{cols[key+"_FT"]}{r-1}-1))')
        addf(f"{key}_G1SH", "Participação G1 na perda", lambda r: f'=IF({cols[key+"_PS"]}{r}="","",{cols["G1_PS"]}{r}/{cols[key+"_PS"]}{r})')
        addf(f"{key}_FR", "TOTAL faturamento R$", lambda r: "=" + sum4("_FR", r))
        addf(f"{key}_CP", "TOTAL custo da perda R$", lambda r: "=" + sum4("_CP", r))
    else:
        addf(f"{key}_CP", f"{key} custo da perda R$", lambda r: f'=IF(OR({cols[key+"_CK"]}{r}="",{cols[key+"_PS"]}{r}=""),"",{cols[key+"_PS"]}{r}*1000*{cols[key+"_CK"]}{r})')
    addf(f"{key}_PMED", f"{key} preço médio R$/kg", lambda r: f'=IF(OR({cols[key+"_FR"]}{r}="",{cols[key+"_FT"]}{r}=""),"",IF({cols[key+"_FT"]}{r}=0,"",{cols[key+"_FR"]}{r}/({cols[key+"_FT"]}{r}*1000)))')
    addf(f"{key}_PF", f"{key} % perda financeira", lambda r: f'=IF(OR({cols[key+"_CP"]}{r}="",{cols[key+"_FR"]}{r}=""),"",IF({cols[key+"_FR"]}{r}=0,"",{cols[key+"_CP"]}{r}/{cols[key+"_FR"]}{r}))')
    addf(f"{key}_RL", f"{key} receita - custo perda", lambda r: f'=IF(OR({cols[key+"_FR"]}{r}="",{cols[key+"_CP"]}{r}=""),"",{cols[key+"_FR"]}{r}-{cols[key+"_CP"]}{r})')
    for src, dst in (("FR", "FRV"), ("PMED", "PMEDV"), ("RL", "RLV")):
        addf(f"{key}_{dst}", f"{key} {src} var % vs ant", lambda r, src=src: '=""' if r == R0 else
             f'=IF(OR({cols[key+"_"+src]}{r}="",{cols[key+"_"+src]}{r-1}=""),"",IF({cols[key+"_"+src]}{r-1}=0,"",{cols[key+"_"+src]}{r}/{cols[key+"_"+src]}{r-1}-1))')


for s in SUBS: fin_block(s[1])
fin_block("TOT", [s[1] for s in SUBS])

for r in range(R0, R1 + 1):
    for k, fn in formulas.items():
        ws[f"{cols[k]}{r}"] = fn(r)
    ws[f"A{r}"].number_format = "mmm/yyyy"; ws[f"C{r}"].number_format = "mmm/yyyy"
for c in ws[5]: c.font = f(9, True); c.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[5].height = 45
ws.freeze_panes = "F6"
for i in range(1, nxt[0]): ws.column_dimensions[L(i)].width = 12


def CR(key): return f"{CALC}!${cols[key]}${R0}:${cols[key]}${R1}"


def SEL(key): return f"INDEX({CR(key)},SelIdx)"


# --- blocos auxiliares dos gráficos: 12 meses até o mês escolhido em CADA aba ---
chart_keys = []
for s_ in SUBS:
    chart_keys += [(f"{s_[1]}_FT", "Faturamento (t)"), (f"{s_[1]}_PP", f"{SHORT[s_[1]]}"), (f"{s_[1]}_PS", "Perda (t)"), (f"{s_[1]}_MS", "Meta (t)"),
                   (f"{s_[1]}_PSOK", "Dentro da meta"), (f"{s_[1]}_PSAT", "No limite"), (f"{s_[1]}_PSOV", "Acima da meta"),
                   (f"{s_[1]}_SMS", "Situação meta")]
chart_keys += [("TOT_PP", "Total empresa"), ("TOT_SMS", "Situação meta total"), ("TOT_PS", "Perda total (t)"), ("TOT_MS", "Meta total (t)")]
gcol = {k: L(4 + j) for j, (k, _) in enumerate(chart_keys)}
GHS = {}
blk_sheets = [panels[s_[1]].title for s_ in SUBS] + [wsT.title, wsR.title]
for bi, title in enumerate(blk_sheets):
    GH = 59 + bi * 16; GHS[title] = GH
    ws[f"A{GH-1}"] = f"Janela dos gráficos da aba {title} (12 meses até o mês escolhido nela). #N/D é proposital: faz a linha pular meses sem dado."
    ws[f"A{GH-1}"].font = f(10, True)
    ws[f"A{GH}"] = "k"; ws[f"B{GH}"] = "Mês"; ws[f"C{GH}"] = "Rótulo"
    for k, t in chart_keys: ws[f"{gcol[k]}{GH}"] = t
    for i in range(12):
        r = GH + 1 + i
        ws[f"A{r}"] = i + 1
        ws[f"B{r}"] = f"=EDATE('{title}'!$Z$1,A{r}-12)"; ws[f"B{r}"].number_format = "mmm/yyyy"
        ws[f"C{r}"] = f'=IFERROR(INDEX($E${R0}:$E${R1},MATCH(B{r},$A${R0}:$A${R1},0)),"")'
        for k, _ in chart_keys:
            idx = f"INDEX(${cols[k]}${R0}:${cols[k]}${R1},MATCH($B{r},$A${R0}:$A${R1},0))"
            blank = '""' if k.endswith(("_FT", "_PS", "_PSOK", "_PSAT", "_PSOV", "_SMS")) else "NA()"
            ws[f"{gcol[k]}{r}"] = f'=IFERROR(IF({idx}="",{blank},{idx}),{blank})'


def add_selector(ws, cell, label="Escolha o mês", note="Vazio = último mês lançado."):
    """Seletor de mês próprio da aba + nomes locais (MesSel, SelIdx, MesTxt...)."""
    col = cell.rstrip("0123456789"); row = int(cell[len(col):])
    ws[f"{col}{row-1}"] = label; ws[f"{col}{row-1}"].font = f(9, True)
    c = ws[cell]; c.fill = fill(C_IN); c.number_format = "mmm/yyyy"; c.font = f(11, True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin); c.alignment = Alignment(horizontal="center")
    ws[f"{col}{row+1}"] = note; ws[f"{col}{row+1}"].font = f(8, color="7F7F7F")
    ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 16)
    dvm = DataValidation(type="list", formula1=f"{CALC}!$C${R0}:$C${R1}", allow_blank=True,
                         error="Escolha um mês da lista (ou apague para ver o último mês lançado).", errorTitle="Mês inválido", showErrorMessage=True)
    ws.add_data_validation(dvm); dvm.add(cell)
    ws["Z1"] = f'=IF({cell}="",UltimoMes,{cell})'; ws["Z1"].number_format = "mmm/yyyy"
    ws["Z2"] = f"=MATCH(Z1,{CALC}!$A${R0}:$A${R1},0)"
    ws["Z3"] = f"=INDEX({CALC}!$D${R0}:$D${R1},Z2)"
    ws["Z4"] = f'=IF(Z2>1,INDEX({CALC}!$D${R0}:$D${R1},Z2-1),"")'
    ws["Z5"] = f'=IF(Z2>12,INDEX({CALC}!$D${R0}:$D${R1},Z2-12),"")'
    ws.column_dimensions["Z"].hidden = True
    q = "'" + ws.title + "'"
    for nm, ref in (("MesSel", "$Z$1"), ("SelIdx", "$Z$2"), ("MesTxt", "$Z$3"), ("MesAntTxt", "$Z$4"), ("MesAnoAntTxt", "$Z$5")):
        ws.defined_names[nm] = DefinedName(nm, attr_text=f"{q}!{ref}")


def amber_note(ws, a, b):
    """Única nota abaixo das tabelas de meta: o que significa o âmbar."""
    ws.merge_cells(f"{a}:{b}")
    ws[a] = '="Âmbar = a perda ainda está abaixo da meta, mas a menos de "&FIXED(TolMeta*100,0)&"% do limite: merece atenção. Passou da meta, mesmo que pouco = vermelho."'
    ws[a].font = f(10, True, AMB_F); ws[a].alignment = Alignment(wrap_text=True, vertical="top")
    r1 = int(a.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")); r2 = int(b.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
    for rr_ in range(r1, r2 + 1): ws.row_dimensions[rr_].height = 18   # altura fixa: o Excel não ajusta células mescladas


# =====================================================================
# Helpers de painel
# =====================================================================
def panel_base(ws, title, color, subtitle):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions["A"].width = 2
    for i in range(2, 18): ws.column_dimensions[L(i)].width = 10.5
    ws.column_dimensions["R"].width = 3
    ws.column_dimensions["S"].width = 16
    for col in "TUVWXYZ": ws.column_dimensions[col].hidden = True
    ws.merge_cells("B2:L2"); ws["B2"] = title; ws["B2"].font = f(22, True, color)
    ws.merge_cells("B3:L3"); ws["B3"] = subtitle; ws["B3"].font = f(10, color="595959")
    ws.merge_cells("M2:Q2"); ws["M2"] = '="Mês de referência"'; ws["M2"].font = f(9, color="595959"); ws["M2"].alignment = Alignment(horizontal="right")
    ws.merge_cells("M3:Q3"); ws["M3"] = '=MesTxt'; ws["M3"].font = f(16, True); ws["M3"].alignment = Alignment(horizontal="right")
    for col in range(2, 18):
        ws.cell(3, col).border = Border(bottom=Side(style="medium", color=color))
    ws.row_dimensions[2].height = 30; ws.row_dimensions[3].height = 24; ws.row_dimensions[4].height = 8
    ws.merge_cells("B5:Q5"); ws["B5"] = "Destaques do mês"; ws["B5"].font = f(11, True)
    for r in (5, 6, 7, 8):
        for col in range(2, 18): ws.cell(r, col).fill = fill(C_SOFT)
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[9].height = 8
    for r in (6, 7, 8):
        ws.merge_cells(f"B{r}:Q{r}"); ws[f"B{r}"].font = f(10)
    ws.row_dimensions[14].height = 8
    setup_print(ws, "A1:Q32")
    add_selector(ws, "S2", note="Vazio = último mês lançado.")


hrow = {}


def helper(ws, label, formula):
    n = hrow.setdefault(ws.title, [1]); r = n[0]; n[0] += 1
    ws[f"U{r}"] = label; ws[f"T{r}"] = formula
    return f"$T${r}"


def card(ws, c1, c2, label, value, fmt, line1, st1, line2, st2=None, accent="404040", meta1=False, meta2=False):
    ws.merge_cells(f"{c1}10:{c2}10"); ws.merge_cells(f"{c1}11:{c2}11"); ws.merge_cells(f"{c1}12:{c2}12"); ws.merge_cells(f"{c1}13:{c2}13")
    ws[f"{c1}10"] = label; ws[f"{c1}10"].font = f(9, True, "595959")
    ws[f"{c1}11"] = value; ws[f"{c1}11"].font = f(22, True, accent); ws[f"{c1}11"].number_format = fmt
    ws[f"{c1}12"] = line1; ws[f"{c1}12"].font = f(10, True)
    ws[f"{c1}13"] = line2; ws[f"{c1}13"].font = f(10, bool(st2), "595959")
    for r in (10, 11, 12, 13):
        ws[f"{c1}{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    box(ws, f"{c1}10:{c2}13")
    if st1: (meta_cf if meta1 else status_cf)(ws, f"{c1}12:{c2}12", st1)
    if st2: (meta_cf if meta2 else status_cf)(ws, f"{c1}13:{c2}13", st2)


def arrow(status, val):
    return f'IF({status}="Igual","= ",IF({val}<0,"▼ ","▲ "))'


def fix(x, d): return f"FIXED({x},{d})"


ws_rows = {10: 16, 11: 36, 12: 20, 13: 20}

# =====================================================================
# PAINÉIS POR SUBSIDIÁRIA
# =====================================================================
for sub, key, color, fcol, etapas in SUBS:
    ws = panels[key]
    panel_base(ws, sub, color, "Faturamento e perdas · em toneladas")
    for r, h in ws_rows.items(): ws.row_dimensions[r].height = h
    H = {}
    for k in ("FT", "PS", "PP", "PA", "DP", "SA", "MS", "PMS", "SMS", "FV", "SF", "FY", "VS", "PY", "DY", "SY", "SFY"):
        H[k] = helper(ws, k, "=" + SEL(f"{key}_{k}"))
    yr = "YEAR(MesSel)"
    cond = f"({CALC}!$B${R0}:$B${R1}={yr})*({CALC}!$A${R0}:$A${R1}<=MesSel)*ISNUMBER({CR(key+'_PS')})*ISNUMBER({CR(key+'_FT')})"
    H["AF"] = helper(ws, "acum fat", f"=SUMPRODUCT({cond},{CR(key+'_FT')})")
    H["AP"] = helper(ws, "acum perda", f"=SUMPRODUCT({cond},{CR(key+'_PS')})")
    H["AQ"] = helper(ws, "acum %", f'=IF({H["AF"]}=0,"",{H["AP"]}/{H["AF"]})')

    # destaques
    verb = f'IF({H["SA"]}="Melhorou","caiu",IF({H["SA"]}="Piorou","subiu","se manteve"))'
    ws["B6"] = (f'=IF({H["PP"]}="","•  Sem dados lançados para "&MesTxt&".",IF({H["PA"]}="","•  % de perda: "&{fix(H["PP"]+"*100",1)}&"% (sem mês anterior para comparar).",'
                f'"•  O % de perda "&{verb}&": de "&{fix(H["PA"]+"*100",1)}&"% em "&MesAntTxt&" para "&{fix(H["PP"]+"*100",1)}&"% em "&MesTxt&"."))')
    ws["B7"] = (f'=IF({H["PMS"]}="","","•  Perda de "&{tfx(H["PS"])}&" t para uma meta de "&{tfx(H["MS"])}&" t: "&IF({H["SMS"]}="Piorou",{acima_t(H["PS"], H["MS"])}&" da meta.",IF({H["SMS"]}="Manteve","no limite da meta (atenção).","dentro da meta, "&{abaixo(H["PS"], H["MS"])}&".")))')
    if len(etapas) > 1:
        names, vals = [], []
        for et, _ in etapas:
            v = helper(ws, f"VT {et}", "=" + SEL(f"{key}_{et}_VT"))
            a = helper(ws, f"abs {et}", f'=IF({v}="","",ABS({v}))')
            n = helper(ws, f"nome {et}", f'="{et}"')
            names.append(n); vals.append((v, a))
        absrefs = ",".join(a for _, a in vals)
        mx = helper(ws, "max abs", f'=IF(COUNT({absrefs})=0,"",MAX({absrefs}))')
        pos = helper(ws, "posição maior var", "=" + "+".join(
            [f'IF(AND({mx}<>"",{vals[i][1]}={mx}' + "".join(f',{vals[j][1]}<>{mx}' for j in range(i)) + f'),{i+1},0)' for i in range(len(vals))]))
        nm = f'CHOOSE({pos},' + ",".join(n for n in names) + ")"
        vv = f'CHOOSE({pos},' + ",".join(v for v, _ in vals) + ")"
        ws["B8"] = (f'=IF({mx}="","","•  Maior variação entre as etapas: "&{nm}&", que "&IF({vv}>0,"subiu ","caiu ")&{fix("ABS(" + vv + ")",1)}&" t em relação a "&MesAntTxt&".")')
    else:
        ws["B8"] = (f'=IF({H["FT"]}="","","•  Faturamento de "&{tfx(H["FT"])}&" t"&IF({H["FV"]}="","",": "&IF({H["FV"]}>=0,"subiu ","caiu ")&{fix("ABS(" + H["FV"] + ")*100",1)}&"% em relação a "&MesAntTxt)'
                    f'&IF({H["FY"]}="","",IF({H["FV"]}="",": "," e ")&IF({H["FY"]}>=0,"subiu ","caiu ")&{fix("ABS(" + H["FY"] + ")*100",1)}&"% em relação a "&MesAnoAntTxt)&".")')

    # cartões
    card(ws, "B", "E", "% DE PERDA SOBRE O FATURAMENTO", f'=IF({H["PP"]}="","—",{H["PP"]})', "0.00%",
         f'=IF({H["DP"]}="","Sem mês anterior para comparar",{arrow(H["SA"], H["DP"])}&{fix("ABS(" + H["DP"] + ")*100",2)}&" p.p. vs "&MesAntTxt)', H["SA"],
         f'=IF({H["DY"]}="",IF({H["PA"]}="","","Mês anterior: "&{fix(H["PA"]+"*100",2)}&"%"),{arrow(H["SY"], H["DY"])}&{fix("ABS(" + H["DY"] + ")*100",2)}&" p.p. vs "&MesAnoAntTxt)', H["SY"], accent=color)
    card(ws, "F", "I", "PERDA × META", f'=IF({H["PS"]}="","—",{H["PS"]})', TFMT,
         f'=IF({H["PMS"]}="","Sem meta cadastrada",IF({H["SMS"]}="Piorou",{acima_t(H["PS"], H["MS"])}&" da meta",IF({H["SMS"]}="Manteve","Atenção · no limite da meta","Dentro da meta · "&{abaixo(H["PS"], H["MS"])}&" da meta")))', H["SMS"],
         f'=IF({H["MS"]}="","","Meta: "&{tfx(H["MS"])}&" t"&IF(OR({H["FT"]}="",{H["FT"]}=0),""," · "&{fix(H["MS"]+"/"+H["FT"]+"*100",1)}&"% do faturamento"))', meta1=True)
    card(ws, "J", "M", "FATURAMENTO", f'=IF({H["FT"]}="","—",{H["FT"]})', TFMT,
         f'=IF({H["FV"]}="","Sem mês anterior para comparar",{arrow(H["SF"], H["FV"])}&{fix("ABS(" + H["FV"] + ")*100",1)}&"% vs "&MesAntTxt)', H["SF"],
         f'=IF({H["FY"]}="","Sem dado do ano anterior",{arrow(H["SFY"], H["FY"])}&{fix("ABS(" + H["FY"] + ")*100",1)}&"% vs "&MesAnoAntTxt)', H["SFY"])
    card(ws, "N", "Q", '="ACUMULADO "&YEAR(MesSel)&" (% DE PERDA)"', f'=IF({H["AQ"]}="","—",{H["AQ"]})', "0.00%",
         f'="Faturamento: "&{tfx(H["AF"])}&" t"', None, f'="Perda: "&{tfx(H["AP"])}&" t"')

    # gráfico principal: faturamento (barras) + % perda (linha)
    ws.merge_cells("B15:I15" if len(etapas) > 1 else "B15:J15"); ws["B15"] = "Faturamento (t) e % de perda · últimos 12 meses"; ws["B15"].font = f(10, True)
    cats = Reference(wsC, min_col=3, min_row=GHS[ws.title] + 1, max_row=GHS[ws.title] + 12)
    bar = BarChart(); bar.type = "col"; bar.gapWidth = 60
    bar.add_data(Reference(wsC, min_col=CI(gcol[key + "_FT"]), min_row=GHS[ws.title], max_row=GHS[ws.title] + 12), titles_from_data=True)
    bar.set_categories(cats)
    bar.series[0].graphicalProperties.solidFill = color; bar.series[0].graphicalProperties.line.noFill = True
    labels(bar.series[0], '#,##0;;;', "inBase", "FFFFFF", True, sz=9)
    bar.y_axis.majorGridlines = None; bar.y_axis.number_format = "#,##0"
    ln = LineChart()
    ln.add_data(Reference(wsC, min_col=CI(gcol[key + "_PP"]), min_row=GHS[ws.title], max_row=GHS[ws.title] + 12), titles_from_data=True)
    s = ln.series[0]
    from openpyxl.chart.series import SeriesLabel
    s.tx = SeriesLabel(v="% de perda")
    s.graphicalProperties.line.solidFill = "262626"; s.graphicalProperties.line.width = 28000
    s.marker.symbol = "circle"; s.marker.size = 6; s.marker.graphicalProperties.solidFill = "262626"; s.smooth = False
    labels(s, '0.0%', "t", "262626", True, sz=9, bg="FFFFFF")   # rótulo com fundo branco: não se mistura com a barra
    ln.y_axis.axId = 200; ln.y_axis.number_format = "0%"; ln.y_axis.crosses = "max"; ln.y_axis.majorGridlines = None
    for ax in (bar.x_axis, bar.y_axis, ln.y_axis): ax.delete = False
    bar += ln
    bar.legend.position = "b"; bar.height = 8.6; bar.width = 17.8
    bar.display_blanks = "gap"
    place(ws, bar, "B", 16, "I" if len(etapas) > 1 else "J", 31)

    if len(etapas) > 1:
        ws.merge_cells("K15:Q15"); ws["K15"] = "Perda por etapa no mês"; ws["K15"].font = f(10, True)
        for a_, b_, t in (("K", "L", "Etapa"), ("M", "M", "Perda"), ("N", "N", "Meta"), ("O", "P", "Situação na meta"), ("Q", "Q", '="vs "&MesAntTxt')):
            if a_ != b_: ws.merge_cells(f"{a_}16:{b_}16")
            c = ws[f"{a_}16"]; c.value = t; c.font = f(9, True, "FFFFFF"); c.alignment = Alignment(horizontal="center", vertical="center")
            for cc in ws[f"{a_}16:{b_}16"][0]: cc.fill = fill(C_HEAD)
        ws.row_dimensions[16].height = 20
        rows = [(et, f"{key}_{et}", "_PT", "_MT", "_PM", "_SM", "_VT", "_SV") for et, _ in etapas] + [(f"Total {key}", key, "_PS", "_MS", "_PMS", "_SMS", "_VS", "_SVS")]
        r = 17
        for name, k, pt, mt, pm, sm, vt, sv in rows:
            hp = {x: helper(ws, f"{name}{x}", "=" + SEL(k + x)) for x in (pt, mt, pm, sm, vt, sv)}
            tot = name.startswith("Total")
            ws.merge_cells(f"K{r}:L{r}"); ws[f"K{r}"] = name; ws[f"K{r}"].font = f(10, True); ws[f"K{r}"].alignment = Alignment(vertical="center", indent=1)
            ws[f"M{r}"] = f'=IF({hp[pt]}="","—",{hp[pt]})'; ws[f"M{r}"].number_format = TFMT
            ws[f"N{r}"] = f'=IF({hp[mt]}="","—",{hp[mt]})'; ws[f"N{r}"].number_format = TFMT
            for c in ("M", "N"): ws[f"{c}{r}"].font = f(10, tot); ws[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="center")
            ws.merge_cells(f"O{r}:P{r}")
            ws[f"O{r}"] = f'=IF({hp[pm]}="","—",IF({hp[sm]}="Piorou",{acima_t(hp[pt], hp[mt])},IF({hp[sm]}="Manteve","No limite",{abaixo(hp[pt], hp[mt])})))'
            ws[f"Q{r}"] = f'=IF({hp[vt]}="","—",{arrow(hp[sv], hp[vt])}&{fix("ABS(" + hp[vt] + ")",1)}&" t")'
            for c in ("O", "Q"): ws[f"{c}{r}"].font = f(10, True); ws[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="center")
            meta_cf(ws, f"O{r}:P{r}", hp[sm]); status_cf(ws, f"Q{r}", hp[sv])
            for col in "KLMNOPQ":
                ws[f"{col}{r}"].border = Border(bottom=Side(style="medium" if tot else "hair", color="808080" if tot else "BFBFBF"), top=Side(style="medium", color="808080") if tot else None)
            ws.row_dimensions[r].height = 26
            r += 1
        amber_note(ws, f"K{r+1}", f"Q{r+3}")
    else:
        ws.merge_cells("L15:Q15"); ws["L15"] = f'="Perda × meta (t) · meta atual: "&IF({H["MS"]}="","—",FIXED({H["MS"]},1)&" t")'; ws["L15"].font = f(10, True)
        b2 = BarChart(); b2.type = "col"; b2.gapWidth = 60; b2.overlap = 100
        for sfx, colr, lblc in (("_PSOK", GOOD_F, "FFFFFF"), ("_PSAT", AMB_BAR, "262626"), ("_PSOV", BAD_F, "FFFFFF")):
            b2.add_data(Reference(wsC, min_col=CI(gcol[key + sfx]), min_row=GHS[ws.title], max_row=GHS[ws.title] + 12), titles_from_data=True)
            b2.series[-1].graphicalProperties.solidFill = colr; b2.series[-1].graphicalProperties.line.noFill = True
            labels(b2.series[-1], '#,##0.0;;;', "inEnd", lblc, True)
        b2.set_categories(cats)
        l2 = LineChart()
        l2.add_data(Reference(wsC, min_col=CI(gcol[key + "_MS"]), min_row=GHS[ws.title], max_row=GHS[ws.title] + 12), titles_from_data=True)
        l2.series[0].graphicalProperties.line.solidFill = "262626"; l2.series[0].graphicalProperties.line.dashStyle = "dash"
        l2.series[0].graphicalProperties.line.width = 19000; l2.series[0].marker.symbol = "none"
        b2.y_axis.number_format = "#,##0.0"; b2.y_axis.majorGridlines = None
        for ax in (b2.x_axis, b2.y_axis): ax.delete = False
        b2 += l2
        b2.legend.position = "b"; b2.height = 8.6; b2.width = 10.6
        b2.display_blanks = "gap"
        place(ws, b2, "L", 16, "Q", 31)

# =====================================================================
# PAINEL CONSOLIDADO
# =====================================================================
ws = wsT
panel_base(ws, "Visão consolidada", "262626", "As 4 subsidiárias lado a lado · em toneladas")
for r, h in ws_rows.items(): ws.row_dimensions[r].height = h
HS = {}
for sub, key, color, *_ in SUBS:
    HS[key] = {k: helper(ws, f"{key} {k}", "=" + SEL(f"{key}_{k}")) for k in ("PP", "PA", "DP", "SA", "PMS", "SMS", "PS", "MS")}
    HS[key]["N"] = helper(ws, f"{key} nome", f'="{sub}"')
HT = {k: helper(ws, f"TOT {k}", "=" + SEL(f"TOT_{k}")) for k in ("PP", "PA", "DP", "SA", "PS", "G1SH")}
dps = ",".join(HS[k]["DP"] for k in HS)
nm_ = helper(ws, "n melhorou", "=" + "+".join(f'({HS[k]["SA"]}="Melhorou")' for k in HS))
np_ = helper(ws, "n piorou", "=" + "+".join(f'({HS[k]["SA"]}="Piorou")' for k in HS))
nc_ = helper(ws, "n comparáveis", "=" + "+".join(f'({HS[k]["SA"]}<>"")' for k in HS))
nmeta = helper(ws, "n dentro meta", "=" + "+".join(f'({HS[k]["SMS"]}="Melhorou")+({HS[k]["SMS"]}="Manteve")' for k in HS))
nlim = helper(ws, "n no limite", "=" + "+".join(f'({HS[k]["SMS"]}="Manteve")' for k in HS))
nmc = helper(ws, "n com meta", "=" + "+".join(f'({HS[k]["SMS"]}<>"")' for k in HS))
mind = helper(ws, "menor delta", f'=IF(COUNT({dps})=0,"",MIN({dps}))')
maxd = helper(ws, "maior delta", f'=IF(COUNT({dps})=0,"",MAX({dps}))')
keys = list(HS)


def posof(target):
    return "+".join(f'IF(AND({target}<>"",{HS[keys[i]]["DP"]}={target}' + "".join(f',{HS[keys[j]]["DP"]}<>{target}' for j in range(i)) + f'),{i+1},0)' for i in range(4))


pmin = helper(ws, "pos menor", "=" + posof(mind)); pmax = helper(ws, "pos maior", "=" + posof(maxd))
ch = lambda pos, fld: f"CHOOSE({pos}," + ",".join(HS[k][fld] for k in keys) + ")"
ws["B6"] = (f'=IF({nc_}=0,"•  Sem mês anterior para comparar.","•  "&{nm_}&" de "&{nc_}&" subsidiárias reduziram o % de perda em relação a "&MesAntTxt'
            f'&IF({np_}>0," ("&{np_}&IF({np_}>1," pioraram)"," piorou)"),"")&".")')
ws["B7"] = f'=IF({nmc}=0,"","•  "&{nmeta}&" de "&{nmc}&" subsidiárias ficaram dentro da meta de perda da diretoria"&IF({nlim}>0," ("&{nlim}&" no limite, atenção)","")&".")'
ws["B8"] = (f'=IF({mind}="","",IF({mind}<0,"•  Maior melhora: "&{ch(pmin,"N")}&" (de "&{fix(ch(pmin,"PA")+"*100",1)}&"% para "&{fix(ch(pmin,"PP")+"*100",1)}&"%).",'
            f'"•  Nenhuma subsidiária reduziu o % de perda. Maior piora: "&{ch(pmax,"N")}&" (de "&{fix(ch(pmax,"PA")+"*100",1)}&"% para "&{fix(ch(pmax,"PP")+"*100",1)}&"%)."))')
spans = [("B", "D"), ("E", "G"), ("H", "J"), ("K", "M"), ("N", "Q")]
for (c1, c2), (sub, key, color, *_) in zip(spans, SUBS):
    h = HS[key]
    card(ws, c1, c2, sub.upper(), f'=IF({h["PP"]}="","—",{h["PP"]})', "0.00%",
         f'=IF({h["DP"]}="","Sem comparação",{arrow(h["SA"], h["DP"])}&{fix("ABS(" + h["DP"] + ")*100",2)}&" p.p. vs "&MesAntTxt)', h["SA"],
         f'=IF({h["PMS"]}="","",IF({h["SMS"]}="Piorou",{acima_t(h["PS"], h["MS"])}&" da meta",IF({h["SMS"]}="Manteve","Atenção · no limite da meta",{abaixo(h["PS"], h["MS"])}&" da meta")))', h["SMS"], accent=color, meta2=True)
card(ws, "N", "Q", "TOTAL DA EMPRESA", f'=IF({HT["PP"]}="","—",{HT["PP"]})', "0.00%",
     f'=IF({HT["DP"]}="","Sem comparação",{arrow(HT["SA"], HT["DP"])}&{fix("ABS(" + HT["DP"] + ")*100",2)}&" p.p. vs "&MesAntTxt)', HT["SA"],
     f'=IF({HT["G1SH"]}="","","G1 = "&{fix(HT["G1SH"]+"*100",0)}&"% de toda a perda")')

ws["B15"] = "% de perda por subsidiária · últimos 12 meses"; ws["B15"].font = f(10, True)
lc = LineChart()
for sub, key, color, *_ in SUBS:
    se = Series(Reference(wsC, min_col=CI(gcol[key + "_PP"]), min_row=GHS[ws.title] + 1, max_row=GHS[ws.title] + 12), title=SHORT[key])
    se.graphicalProperties.line.solidFill = color; se.graphicalProperties.line.width = 25000
    se.marker.symbol = "circle"; se.marker.size = 5; se.marker.graphicalProperties.solidFill = color; se.marker.graphicalProperties.line.solidFill = color; se.smooth = False
    last_label(se, '0.0%', "r", color)
    lc.series.append(se)
se = Series(Reference(wsC, min_col=CI(gcol["TOT_PP"]), min_row=GHS[ws.title] + 1, max_row=GHS[ws.title] + 12), title="Total empresa")
se.graphicalProperties.line.solidFill = "262626"; se.graphicalProperties.line.dashStyle = "dash"; se.graphicalProperties.line.width = 22000
se.marker.symbol = "none"; se.smooth = False
last_label(se, '0.0%', "r", "262626")
lc.series.append(se)
lc.set_categories(Reference(wsC, min_col=3, min_row=GHS[ws.title] + 1, max_row=GHS[ws.title] + 12))
lc.y_axis.number_format = "0%"
for ax in (lc.x_axis, lc.y_axis): ax.delete = False
lc.legend.position = "b"; lc.height = 8.6; lc.width = 17.8
lc.display_blanks = "gap"
place(ws, lc, "B", 16, "J", 31)

ws.merge_cells("L15:Q15"); ws["L15"] = "Perda por etapa × meta no mês"; ws["L15"].font = f(10, True)
for a_, b_, t in (("L", "M", "Etapa"), ("N", "N", "Perda"), ("O", "O", "Meta"), ("P", "Q", "Situação na meta")):
    if a_ != b_: ws.merge_cells(f"{a_}16:{b_}16")
    c = ws[f"{a_}16"]; c.value = t; c.font = f(9, True, "FFFFFF"); c.alignment = Alignment(horizontal="center", vertical="center")
    for cc in ws[f"{a_}16:{b_}16"][0]: cc.fill = fill(C_HEAD)
ws.row_dimensions[16].height = 20
r = 17
for sub, key, color, fcol, etapas in SUBS:
    for et, _ in etapas:
        k = f"{key}_{et}"
        pt = helper(ws, f"{k} PT", "=" + SEL(k + "_PT")); mt = helper(ws, f"{k} MT", "=" + SEL(k + "_MT"))
        pm = helper(ws, f"{k} PM", "=" + SEL(k + "_PM")); smh = helper(ws, f"{k} SM", "=" + SEL(k + "_SM"))
        ws.merge_cells(f"L{r}:M{r}"); ws[f"L{r}"] = SHORT[key] if et == "Total" else f"{SHORT[key]} {et}"
        ws[f"L{r}"].font = f(9, True); ws[f"L{r}"].alignment = Alignment(vertical="center", indent=1)
        ws[f"N{r}"] = f'=IF({pt}="","—",{pt})'; ws[f"O{r}"] = f'=IF({mt}="","—",{mt})'
        for c in ("N", "O"): ws[f"{c}{r}"].number_format = TFMT; ws[f"{c}{r}"].font = f(9); ws[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"P{r}:Q{r}")
        ws[f"P{r}"] = f'=IF({pm}="","—",IF({smh}="Piorou",{acima_t(pt, mt)},IF({smh}="Manteve","No limite",{abaixo(pt, mt)})))'
        ws[f"P{r}"].font = f(9, True); ws[f"P{r}"].alignment = Alignment(horizontal="center", vertical="center")
        meta_cf(ws, f"P{r}:Q{r}", smh)
        for col in "LMNOPQ": ws[f"{col}{r}"].border = Border(bottom=Side(style="hair", color="BFBFBF"))
        ws.row_dimensions[r].height = 21
        r += 1
amber_note(ws, f"L{r+1}", f"Q{r+4}")

# =====================================================================
# ONDE FOCAR (Pareto, prioridade, tendência e ganho potencial · 12 meses)
# =====================================================================
ws = wsO
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
ws.column_dimensions["A"].width = 2
for i in range(2, 18): ws.column_dimensions[L(i)].width = 10.5
ws.column_dimensions["J"].width = 13; ws.column_dimensions["R"].width = 3; ws.column_dimensions["S"].width = 16
for col in "TUVWXYZ": ws.column_dimensions[col].hidden = True
add_selector(ws, "S2", note="Vazio = último mês lançado.")
ws["S5"] = "Etapa \"pesa muito\" a partir de"; ws["S5"].font = f(9, True)
c = ws["S6"]; c.value = 0.10; c.number_format = "0%"; c.fill = fill(C_IN); c.font = f(11, True); c.alignment = Alignment(horizontal="center")
c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
ws["S7"] = "da perda total (12 meses)."; ws["S7"].font = f(8, color="7F7F7F")
dvG = DataValidation(type="decimal", operator="between", formula1="0.01", formula2="0.9", error="Digite um percentual entre 1% e 90%.", showErrorMessage=True)
ws.add_data_validation(dvG); dvG.add("S6")
GRANDE = "$S$6"

ETAPAS = [("Linha C", "LC_Total", "LC"), ("Linha D", "LD_Total", "LD"),
          ("G2 Formação", "G2_Formação", "G2"), ("G2 Acabamento", "G2_Acabamento", "G2"),
          ("G1 Corte", "G1_Corte", "G1"), ("G1 Formação", "G1_Formação", "G1"), ("G1 Acabamento", "G1_Acabamento", "G1")]
NE = len(ETAPAS)
# ---------- grade auxiliar oculta: 12 meses × etapa (perda, faturamento do setor, taxa, ganho) ----------
G0 = 3; G1r = G0 + 11            # linhas dos 12 meses
ws["AB2"] = "Mês"; ws["AC2"] = "Linha"
for i in range(12):
    r = G0 + i
    ws[f"AB{r}"] = f"=EDATE($Z$1,{i - 11})"; ws[f"AB{r}"].number_format = "mmm/yyyy"
    ws[f"AC{r}"] = f'=IFERROR(MATCH(AB{r},{CALC}!$A${R0}:$A${R1},0),"")'
S0 = 18                          # linhas do resumo por etapa (não ordenado)
shdr = ["Etapa", "Perda 12m", "Fat 12m", "Taxa 12m", "% do total", "N meses", "Bench (3 melhores)", "Ganho (t)", "Últ. 3m", "3m ant.", "Tendência",
        "Status", "Grande", "Prioridade", "Alerta", "Rank", "Média taxa", "Últ. acima", "Virada", "Antes", "Depois"]
SC = {h: L(CI("AB") + j) for j, h in enumerate(shdr)}
for h, col in SC.items(): ws[f"{col}{S0-1}"] = h
for e, (name, k, sk) in enumerate(ETAPAS):
    base = CI("AD") + e * 4
    cp, cf, ct, cg = (L(base + j) for j in range(4))
    ws[f"{cp}2"] = f"{name} perda"; ws[f"{cf}2"] = f"{name} fat"; ws[f"{ct}2"] = f"{name} taxa"; ws[f"{cg}2"] = f"{name} ganho"
    PT, FT = cols[f"{k}_PT"], cols[f"{sk}_FT"]
    sr = S0 + e
    for i in range(12):
        r = G0 + i
        ip = f"INDEX({CALC}!${PT}${R0}:${PT}${R1},$AC{r})"; iff = f"INDEX({CALC}!${FT}${R0}:${FT}${R1},$AC{r})"
        ok = f'AND($AC{r}<>"",ISNUMBER({ip}),ISNUMBER({iff}))'
        ws[f"{cp}{r}"] = f'=IF({ok},{ip},"")'
        ws[f"{cf}{r}"] = f'=IF({ok},{iff},"")'
        ws[f"{ct}{r}"] = f'=IF(AND(ISNUMBER({cp}{r}),ISNUMBER({cf}{r})),IF({cf}{r}=0,"",{cp}{r}/{cf}{r}),"")'
        ws[f"{cg}{r}"] = f'=IF(AND(ISNUMBER({ct}{r}),ISNUMBER(${SC["Bench (3 melhores)"]}${sr})),MAX(0,{ct}{r}-${SC["Bench (3 melhores)"]}${sr})*{cf}{r},0)'
    rng_ = lambda c_, a=G0, b=G1r: f"{c_}{a}:{c_}{b}"
    v = {}
    v["Etapa"] = name
    v["Perda 12m"] = f'=IF(COUNT({rng_(cp)})=0,"",SUM({rng_(cp)}))'
    v["Fat 12m"] = f'=IF(COUNT({rng_(cf)})=0,"",SUM({rng_(cf)}))'
    v["Taxa 12m"] = f'=IF(OR({SC["Fat 12m"]}{sr}="",{SC["Fat 12m"]}{sr}=0),"",{SC["Perda 12m"]}{sr}/{SC["Fat 12m"]}{sr})'
    v["% do total"] = f'=IF({SC["Perda 12m"]}{sr}="","",{SC["Perda 12m"]}{sr}/SUM(${SC["Perda 12m"]}${S0}:${SC["Perda 12m"]}${S0+NE-1}))'
    v["N meses"] = f"=COUNT({rng_(ct)})"
    v["Bench (3 melhores)"] = f'=IF({SC["N meses"]}{sr}<3,"",(SMALL({rng_(ct)},1)+SMALL({rng_(ct)},2)+SMALL({rng_(ct)},3))/3)'
    v["Ganho (t)"] = f'=IF({SC["Bench (3 melhores)"]}{sr}="","",SUM({rng_(cg)}))'
    l3, p3 = (G1r - 2, G1r), (G1r - 5, G1r - 3)
    v["Últ. 3m"] = f'=IF(COUNT({rng_(ct, *l3)})<3,"",SUM({rng_(cp, *l3)})/SUM({rng_(cf, *l3)}))'
    v["3m ant."] = f'=IF(COUNT({rng_(ct, *p3)})<3,"",SUM({rng_(cp, *p3)})/SUM({rng_(cf, *p3)}))'
    v["Tendência"] = f'=IF(OR({SC["Últ. 3m"]}{sr}="",{SC["3m ant."]}{sr}="",{SC["3m ant."]}{sr}=0),"",{SC["Últ. 3m"]}{sr}/{SC["3m ant."]}{sr}-1)'
    v["Status"] = f'=IF({SC["Tendência"]}{sr}="","",IF({SC["Tendência"]}{sr}>0,"Piorando",IF({SC["Tendência"]}{sr}<0,"Melhorando","Estável")))'
    v["Grande"] = f'=IF({SC["% do total"]}{sr}="","",{SC["% do total"]}{sr}>={GRANDE})'
    v["Prioridade"] = (f'=IF(OR({SC["Status"]}{sr}="",{SC["Grande"]}{sr}=""),"—",IF({SC["Grande"]}{sr},IF({SC["Status"]}{sr}="Piorando","Agir agora","Padronizar"),'
                       f'IF({SC["Status"]}{sr}="Piorando","Acompanhar","Manter")))')
    t4 = [f"{ct}{G1r - j}" for j in range(4)]   # t4[0] = mês atual, t4[3] = 3 meses antes
    v["Alerta"] = f'=IF(COUNT({t4[3]}:{t4[0]})<4,"",IF(AND({t4[0]}>{t4[1]},{t4[1]}>{t4[2]},{t4[2]}>{t4[3]}),"3 meses de alta",""))'
    v["Rank"] = (f'=IF({SC["Perda 12m"]}{sr}="",{NE}+{e},RANK({SC["Perda 12m"]}{sr},${SC["Perda 12m"]}${S0}:${SC["Perda 12m"]}${S0+NE-1})'
                 f'+COUNTIF(${SC["Perda 12m"]}${S0}:{SC["Perda 12m"]}{sr},{SC["Perda 12m"]}{sr})-1)')
    # mês da virada: mês seguinte ao último mês com % de perda acima da média da janela
    # (usado para dizer "desde fev/2026" na frase do Padronizar; exige 3 meses depois da virada)
    v["Média taxa"] = f'=IF({SC["N meses"]}{sr}<4,"",AVERAGE({rng_(ct)}))'
    v["Últ. acima"] = f'=IF({SC["Média taxa"]}{sr}="",0,SUMPRODUCT(MAX(ISNUMBER({rng_(ct)})*({rng_(ct)}>{SC["Média taxa"]}{sr})*ROW({rng_(ct)}))))'
    v["Virada"] = f'=IF(OR({SC["Últ. acima"]}{sr}=0,{SC["Últ. acima"]}{sr}>{G1r - 3}),"",INDEX($AB$1:$AB${G1r},{SC["Últ. acima"]}{sr}+1))'
    v["Antes"] = f'=IF({SC["Virada"]}{sr}="","",SUMIFS({rng_(cp)},$AB${G0}:$AB${G1r},"<"&{SC["Virada"]}{sr})/SUMIFS({rng_(cf)},$AB${G0}:$AB${G1r},"<"&{SC["Virada"]}{sr}))'
    v["Depois"] = f'=IF({SC["Virada"]}{sr}="","",SUMIFS({rng_(cp)},$AB${G0}:$AB${G1r},">="&{SC["Virada"]}{sr})/SUMIFS({rng_(cf)},$AB${G0}:$AB${G1r},">="&{SC["Virada"]}{sr}))'
    for h, val in v.items(): ws[f"{SC[h]}{sr}"] = val
    ws[f"{SC['Virada']}{sr}"].number_format = "mmm/yyyy"
for col in range(CI("AA"), max(base + 4, CI(SC["Depois"])) + 1): ws.column_dimensions[L(col)].hidden = True
SR = f"${SC['Rank']}${S0}:${SC['Rank']}${S0+NE-1}"


def srt(h, n): return f"INDEX(${SC[h]}${S0}:${SC[h]}${S0+NE-1},MATCH({n},{SR},0))"


# ---------- cabeçalho ----------
ws.merge_cells("B2:L2"); ws["B2"] = "Onde focar"; ws["B2"].font = f(22, True, "262626")
ws.merge_cells("B3:L3"); ws["B3"] = "Onde está a perda, para onde ela está indo e quanto dá para ganhar · em toneladas"; ws["B3"].font = f(10, color="595959")
ws.merge_cells("M2:Q2"); ws["M2"] = "Últimos 12 meses até"; ws["M2"].font = f(9, color="595959"); ws["M2"].alignment = Alignment(horizontal="right")
ws.merge_cells("M3:Q3"); ws["M3"] = "=MesTxt"; ws["M3"].font = f(16, True); ws["M3"].alignment = Alignment(horizontal="right")
for col in range(2, 18): ws.cell(3, col).border = Border(bottom=Side(style="medium", color="262626"))
ws.row_dimensions[2].height = 30; ws.row_dimensions[3].height = 24; ws.row_dimensions[4].height = 8

# ---------- tabela ordenada (Pareto) ----------
T0 = 13
ws.merge_cells(f"J{T0-2}:Q{T0-2}"); ws[f"J{T0-2}"] = "Prioridade por etapa · da maior para a menor perda"; ws[f"J{T0-2}"].font = f(10, True)
th = [("J", "K", "Etapa"), ("L", "L", "Perda (t)"), ("M", "M", "% do total"), ("N", "N", "% acum."), ("O", "O", "Tendência (média 3 meses)"), ("P", "P", "Prioridade"), ("Q", "Q", "Ganho pot. (t)")]
for a_, b_, t in th:
    if a_ != b_: ws.merge_cells(f"{a_}{T0-1}:{b_}{T0-1}")
    c = ws[f"{a_}{T0-1}"]; c.value = t; c.font = f(9, True, "FFFFFF"); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cc in ws[f"{a_}{T0-1}:{b_}{T0-1}"][0]: cc.fill = fill(C_HEAD)
ws.row_dimensions[T0 - 1].height = 38
for n in range(1, NE + 1):
    r = T0 + n - 1
    ws.merge_cells(f"J{r}:K{r}")
    ws[f"J{r}"] = f'={srt("Etapa", n)}&IF({srt("Alerta", n)}="",""," ⚠")'
    ws[f"J{r}"].font = f(9, True); ws[f"J{r}"].alignment = Alignment(vertical="center", indent=1)
    ws[f"L{r}"] = f'=IF({srt("Perda 12m", n)}="","—",{srt("Perda 12m", n)})'; ws[f"L{r}"].number_format = "#,##0.0"
    ws[f"M{r}"] = f'=IF({srt("% do total", n)}="","—",{srt("% do total", n)})'; ws[f"M{r}"].number_format = "0.0%"
    ws[f"N{r}"] = f'=IF(ISNUMBER(M{r}),SUM($M${T0}:M{r}),"—")'; ws[f"N{r}"].number_format = "0%"
    tv = srt("Tendência", n)
    ws[f"O{r}"] = f'=IF({tv}="","—",IF({tv}>0,"▲ ",IF({tv}<0,"▼ ","= "))&FIXED(ABS({tv})*100,0)&"%")'
    ws[f"AA{r}"] = f'={srt("Status", n)}'   # status para a cor (coluna oculta)
    ws[f"P{r}"] = f'={srt("Prioridade", n)}'
    gv = srt("Ganho (t)", n)
    ws[f"Q{r}"] = f'=IF({gv}="","—",{gv})'; ws[f"Q{r}"].number_format = "#,##0.0"
    for col in "LMNOPQ":
        ws[f"{col}{r}"].font = f(9, col in "OP"); ws[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center")
    for col in "JKLMNOPQ": ws[f"{col}{r}"].border = Border(bottom=Side(style="hair", color="BFBFBF"))
    ws.row_dimensions[r].height = 21
    ws.conditional_formatting.add(f"O{r}", FormulaRule(formula=[f'$AA${r}="Piorando"'], font=Font(name=FONT, color=BAD_F, bold=True), fill=fill(BAD_B)))
    ws.conditional_formatting.add(f"O{r}", FormulaRule(formula=[f'$AA${r}="Melhorando"'], font=Font(name=FONT, color=GOOD_F, bold=True), fill=fill(GOOD_B)))
    for val, fc, bc in (("Agir agora", "FFFFFF", BAD_F), ("Padronizar", "FFFFFF", GOOD_F), ("Acompanhar", AMB_F, AMB_B)):
        ws.conditional_formatting.add(f"P{r}", FormulaRule(formula=[f'$P${r}="{val}"'], font=Font(name=FONT, color=fc, bold=True), fill=fill(bc)))
TL = T0 + NE - 1
r = TL + 1
for col in "JKLMNOPQ": ws[f"{col}{r}"].border = Border(top=Side(style="medium", color="808080"))
ws.merge_cells(f"J{r}:K{r}"); ws[f"J{r}"] = "Total"; ws[f"J{r}"].font = f(9, True); ws[f"J{r}"].alignment = Alignment(vertical="center", indent=1)
ws[f"L{r}"] = f'=SUM(L{T0}:L{TL})'; ws[f"L{r}"].number_format = "#,##0.0"
ws[f"Q{r}"] = f'=SUM(Q{T0}:Q{TL})'; ws[f"Q{r}"].number_format = "#,##0.0"
for col in "LQ": ws[f"{col}{r}"].font = f(9, True); ws[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[r].height = 21
TOT_R = r

# ---------- o que é a tendência (logo abaixo da tabela, fonte legível) ----------
r = TOT_R + 1
P3A, P3B = mtxt("EDATE($Z$1,-2)", True), mtxt("$Z$1", True)
Q3A, Q3B = mtxt("EDATE($Z$1,-5)", True), mtxt("EDATE($Z$1,-3)", True)
ws.merge_cells(f"J{r}:Q{r+1}")
ws[f"J{r}"] = (f'="Tendência: compara a média do % de perda de "&{P3A}&" a "&{P3B}&" com a média de "&{Q3A}&" a "&{Q3B}&". '
               f'Não é a comparação com o mês anterior."')
ws[f"J{r}"].font = f(10, True, "262626"); ws[f"J{r}"].alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 16; ws.row_dimensions[r + 1].height = 16
TEND_R = r + 1

# ---------- legenda das prioridades ----------
r = TEND_R + 2
ws.merge_cells(f"J{r}:Q{r}"); ws[f"J{r}"] = f'="Como ler a prioridade (“pesa muito” = "&FIXED({GRANDE}*100,0)&"% ou mais de toda a perda)"'
ws[f"J{r}"].font = f(11, True); ws.row_dimensions[r].height = 20
leg = [("Agir agora", "pesa muito na perda total e está piorando: investigar a causa agora.", "FFFFFF", BAD_F),
       ("Padronizar", "pesa muito na perda total e está melhorando: descobrir o que deu certo e tornar padrão.", "FFFFFF", GOOD_F),
       ("Acompanhar", "pesa pouco, mas está piorando: acompanhar mês a mês.", AMB_F, AMB_B),
       ("Manter", "pesa pouco e está estável ou melhorando: seguir como está.", "404040", "FFFFFF")]
for i, (t, d, fc, bc) in enumerate(leg):
    rr_ = r + 1 + i
    c = ws[f"J{rr_}"]; c.value = t; c.font = f(10, True, fc); c.fill = fill(bc); c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells(f"K{rr_}:Q{rr_}"); ws[f"K{rr_}"] = d; ws[f"K{rr_}"].font = f(10, color="262626"); ws[f"K{rr_}"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[rr_].height = 21
LEG_END = r + len(leg)

# ---------- leitura do mês (bullets automáticos) ----------
ws.merge_cells("B5:Q5"); ws["B5"] = "Leitura dos últimos 12 meses"; ws["B5"].font = f(11, True)
for rr_ in range(5, 10):
    for col in range(2, 18): ws.cell(rr_, col).fill = fill(C_SOFT)
    ws.row_dimensions[rr_].height = 18
for rr_ in (6, 7, 8, 9): ws.merge_cells(f"B{rr_}:Q{rr_}"); ws[f"B{rr_}"].font = f(10)
ws.row_dimensions[8].height = 30; ws["B8"].alignment = Alignment(wrap_text=True, vertical="center")
n80 = f"MIN({NE},COUNTIF($N${T0}:$N${TL},\"<0.8\")+1)"
ws["T1"] = "=" + n80; ws["U1"] = "n etapas até 80%"
names = "&".join(f'IF({i}<=$T$1,IF({i}=1,"",IF({i}=$T$1," e ",", "))&SUBSTITUTE($J${T0+i-1}," ⚠",""),"")' for i in range(1, NE + 1))
ws["B6"] = f'=IF(COUNT($L${T0}:$L${TL})=0,"•  Sem dados de perda nos últimos 12 meses.","•  "&$T$1&" de {NE} etapas concentram "&FIXED(INDEX($N${T0}:$N${TL},$T$1)*100,0)&"% de toda a perda: "&{names}&".")'


def lst(prio):
    return "&".join(f'IF($P${T0+i}="{prio}",", "&SUBSTITUTE($J${T0+i}," ⚠",""),"")' for i in range(NE))


ws["T2"] = f'=MID({lst("Agir agora")},3,300)'; ws["U2"] = "lista Agir agora"
ws["T3"] = f'=MID({lst("Padronizar")},3,300)'; ws["U3"] = "lista Padronizar"
alerts = "&".join(f'IF(RIGHT($J${T0+i},1)="⚠",", "&SUBSTITUTE($J${T0+i}," ⚠",""),"")' for i in range(NE))
ws["T4"] = f'=MID({alerts},3,300)'; ws["U4"] = "lista alerta"
# 1ª etapa "Padronizar" (a de maior perda): nome, mês da virada e % antes/depois
ws["T5"] = f'=IFERROR(MATCH("Padronizar",$P${T0}:$P${TL},0),"")'; ws["U5"] = "posição 1º Padronizar"
ws["T6"] = f'=IF($T$5="","",{srt("Etapa", "$T$5")})'; ws["U6"] = "nome 1º Padronizar"
ws["T7"] = f'=IF($T$5="","",{srt("Virada", "$T$5")})'; ws["U7"] = "virada"; ws["T7"].number_format = "mmm/yyyy"
ws["T8"] = f'=IF($T$5="","",{srt("Antes", "$T$5")})'; ws["U8"] = "% antes"
ws["T9"] = f'=IF($T$5="","",{srt("Depois", "$T$5")})'; ws["U9"] = "% depois"
ws["T10"] = f'=IF(OR($T$7="",$T$7=0),"",{mtxt("$T$7")})'; ws["U10"] = "virada (texto)"
ws["B7"] = ('=IF($T$2="","•  Nenhuma etapa que pesa muito na perda está piorando.","•  Agir agora (pesa muito na perda e está piorando): "&$T$2&". '
            'A média dos últimos 3 meses está acima da média dos 3 meses anteriores.")&IF($T$4="",""," ⚠ 3 meses seguidos de alta: "&$T$4&".")')
ws["B8"] = ('=IF($T$3="","","•  Padronizar (pesa muito na perda e está melhorando): "&$T$3&". "&'
            'IF(AND($T$10<>"",ISNUMBER($T$8),ISNUMBER($T$9),$T$8>$T$9),'
            '"Na "&$T$6&", o % de perda caiu a partir de "&$T$10&": de "&FIXED($T$8*100,1)&"% do faturamento nos meses anteriores para "&FIXED($T$9*100,1)&"% de "&$T$10&" em diante. '
            'Vale a pena descobrir o que deu certo em "&$T$10&" e tornar padrão.",'
            '"Vale a pena descobrir o que deu certo e tornar padrão."))')
ws["B9"] = (f'=IF(Q{TOT_R}=0,"","•  Ganho potencial estimado: "&FIXED(Q{TOT_R},0)&" t em 12 meses (cerca de "&FIXED(Q{TOT_R}/MAX(1,MAX(${SC["N meses"]}${S0}:${SC["N meses"]}${S0+NE-1})),1)&" t/mês), '
            f'se, nos meses ruins, cada etapa tivesse ficado no nível dos seus 3 melhores meses.")')
ws.row_dimensions[10].height = 8

# ---------- gráfico de Pareto ----------
ws.merge_cells(f"B{T0-2}:I{T0-2}"); ws[f"B{T0-2}"] = "Pareto da perda · últimos 12 meses (t e % acumulado)"; ws[f"B{T0-2}"].font = f(10, True)
for n in range(1, NE + 1):   # séries do gráfico (colunas ocultas V/W/X)
    r = T0 + n - 1
    ws[f"V{r}"] = f'=SUBSTITUTE(J{r}," ⚠","")'
    ws[f"W{r}"] = f'=IF(ISNUMBER(L{r}),L{r},NA())'
    ws[f"X{r}"] = f'=IF(ISNUMBER(N{r}),N{r},NA())'
ws[f"V{T0-1}"] = "Etapa"; ws[f"W{T0-1}"] = "Perda (t)"; ws[f"X{T0-1}"] = "% acumulado"
pb = BarChart(); pb.type = "col"; pb.gapWidth = 40
pb.add_data(Reference(ws, min_col=CI("W"), min_row=T0 - 1, max_row=TL), titles_from_data=True)
pb.set_categories(Reference(ws, min_col=CI("V"), min_row=T0, max_row=TL))
pb.series[0].graphicalProperties.solidFill = "B4BFCC"   # cinza-azulado claro: o rótulo vermelho do % acumulado fica legível sobre a barra
labels(pb.series[0], '#,##0', "outEnd", "262626", True)
pb.y_axis.majorGridlines = None; pb.y_axis.number_format = "#,##0"
pl = LineChart()
pl.add_data(Reference(ws, min_col=CI("X"), min_row=T0 - 1, max_row=TL), titles_from_data=True)
s = pl.series[0]
s.graphicalProperties.line.solidFill = BAD_F; s.graphicalProperties.line.width = 25000
s.marker.symbol = "circle"; s.marker.size = 6; s.marker.graphicalProperties.solidFill = BAD_F; s.marker.graphicalProperties.line.solidFill = BAD_F; s.smooth = False
labels(s, '0%', "t", BAD_F, True)
pl.y_axis.axId = 200; pl.y_axis.number_format = "0%"; pl.y_axis.crosses = "max"; pl.y_axis.majorGridlines = None
pl.y_axis.scaling.min = 0; pl.y_axis.scaling.max = 1.05
for ax in (pb.x_axis, pb.y_axis, pl.y_axis): ax.delete = False
pb += pl
pb.height = 9.5; pb.width = 17.0
pb.display_blanks = "gap"
pb.visible_cells_only = False   # dados vêm de colunas ocultas
place(ws, pb, "B", T0 - 1, "I", LEG_END)

# ---------- notas ----------
r = LEG_END + 2
notes = ['Etapas: Linha C e Linha D como setor inteiro; G1 e G2 por etapa. % = perda da etapa ÷ faturamento do setor. ⚠ = 3 meses seguidos de alta.',
         'Ganho potencial (ESTIMATIVA): toneladas que a etapa deixaria de perder se, em cada mês acima do seu nível, repetisse a média dos seus 3 melhores meses no período. Não é meta.',
         'Esta aba aponta onde investigar; os relatórios não trazem os motivos da perda. Fonte: aba Lançamentos.']
for i, t in enumerate(notes):
    ws.merge_cells(f"B{r+i}:Q{r+i}"); ws[f"B{r+i}"] = t; ws[f"B{r+i}"].font = f(9, color="595959", italic=True)
setup_print(ws, f"A1:Q{r+len(notes)-1}")

# =====================================================================
# RESUMO (mês escolhido × mês de comparação + histórico 12 meses)
# =====================================================================
ws = wsR
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2; ws.column_dimensions["B"].width = 22
for col in "CDEFGHIJKLMNOPQRS": ws.column_dimensions[col].width = 10.5
ws.column_dimensions["N"].width = 13; ws.column_dimensions["T"].width = 3
add_selector(ws, "U2", label="Mês do resumo")
ws["U4"] = "Comparar com"; ws["U4"].font = f(9, True)
c = ws["U5"]; c.fill = fill(C_IN); c.number_format = "mmm/yyyy"; c.font = f(11, True); c.alignment = Alignment(horizontal="center")
c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
ws["U6"] = "Vazio = mês anterior."; ws["U6"].font = f(8, color="7F7F7F")
ws["U7"] = "A ordem não importa: a variação vai sempre do mais antigo para o mais recente."; ws["U7"].font = f(8, color="7F7F7F")
dvc = DataValidation(type="list", formula1=f"{CALC}!$C${R0}:$C${R1}", allow_blank=True, error="Escolha um mês da lista.", errorTitle="Mês inválido", showErrorMessage=True)
ws.add_data_validation(dvc); dvc.add("U5")
ws["Z6"] = f'=IF(U5="",IF(Z2>1,Z2-1,""),MATCH(U5,{CALC}!$A${R0}:$A${R1},0))'
ws["Z7"] = f'=IF(Z6="","—",INDEX({CALC}!$D${R0}:$D${R1},Z6))'
ws["Z8"] = '=IF(Z6="","",MIN(Z2,Z6))'
ws["Z9"] = '=IF(Z6="",Z2,MAX(Z2,Z6))'
ws["Z10"] = f'=IF(Z8="","—",INDEX({CALC}!$D${R0}:$D${R1},Z8))'
ws["Z11"] = f'=INDEX({CALC}!$D${R0}:$D${R1},Z9)'
for nm, ref in (("CompIdx", "$Z$6"), ("CompTxt", "$Z$7"), ("IdxA", "$Z$8"), ("IdxB", "$Z$9"), ("TxtA", "$Z$10"), ("TxtB", "$Z$11")):
    ws.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!{ref}")


def AT(key, idx): return f'IF({idx}="","",INDEX({CR(key)},{idx}))'


ws.merge_cells("B2:L2"); ws["B2"] = "Resumo do mês"; ws["B2"].font = f(20, True)
ws.merge_cells("M2:S2"); ws["M2"] = '=TxtA&"  →  "&TxtB'; ws["M2"].font = f(16, True); ws["M2"].alignment = Alignment(horizontal="right")
ws.merge_cells("B3:S3")
ws["B3"] = '="Toneladas. A variação é sempre do mês mais antigo para o mais recente ("&TxtA&" → "&TxtB&"). Meta, acumulado e ano anterior referem-se a "&MesTxt&"."'
ws["B3"].font = f(9, color="595959")
groups = [("C", "E", "Faturamento (t)"), ("F", "H", "Perda (t)"), ("I", "L", "% de perda"), ("M", "N", "Meta"), ("O", "Q", '="Acumulado "&YEAR(MesSel)'), ("R", "S", "Ano anterior")]
for a_, b_, t in groups:
    ws.merge_cells(f"{a_}4:{b_}4"); ws[f"{a_}4"] = t
    ws[f"{a_}4"].font = f(9, True, "FFFFFF"); ws[f"{a_}4"].alignment = Alignment(horizontal="center")
    for cc in ws[f"{a_}4:{b_}4"][0]: cc.fill = fill("262626")
heads = {"B": "Subsidiária / etapa", "C": "=TxtA", "D": "=TxtB", "E": "Variação", "F": "=TxtA", "G": "=TxtB", "H": "Variação",
         "I": "=TxtA", "J": "=TxtB", "K": "Variação (p.p.)", "L": "Status", "M": "Meta (t)", "N": "Situação",
         "O": "Faturamento", "P": "Perda", "Q": "% de perda", "R": "Faturamento", "S": "% perda (p.p.)"}
for col, h in heads.items():
    c = ws[f"{col}5"]; c.value = h; c.font = f(9, True, "FFFFFF"); c.fill = fill(C_HEAD); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[5].height = 28


def acc(key_ft, key_ps, which):
    cond = f"({CALC}!$B${R0}:$B${R1}=YEAR(MesSel))*({CALC}!$A${R0}:$A${R1}<=MesSel)*ISNUMBER({CR(key_ps)})" + (f"*ISNUMBER({CR(key_ft)})" if key_ft else "")
    return f"SUMPRODUCT({cond},{CR(key_ps if which == 'P' else key_ft)})"


def dash(x): return f'=IF({x}="","—",{x})'


def situ(pm, sm, p, m): return f'=IF({pm}="","—",IF({sm}="Piorou",{acima_t(p, m)},IF({sm}="Manteve","No limite",{abaixo(p, m)})))'


lines = []
for sub, key, color, fcol, etapas in SUBS:
    lines.append(("sub", sub, key, color))
    if len(etapas) > 1:
        for et, _ in etapas: lines.append(("et", et, f"{key}_{et}", color))
lines.append(("tot", "Total da empresa", "TOT", "262626"))
rr = 6; sub_rows = []
for kind, name, k, color in lines:
    c = ws.cell(rr, 2, name if kind != "et" else "   " + name)
    c.font = f(10, kind != "et", color if kind == "sub" else "000000")
    num = lambda col: f"ISNUMBER({col}{rr})"
    if kind in ("sub", "tot"):
        v = {"C": dash(AT(k + "_FT", "IdxA")), "D": dash(AT(k + "_FT", "IdxB")), "F": dash(AT(k + "_PS", "IdxA")), "G": dash(AT(k + "_PS", "IdxB")),
             "I": dash(AT(k + "_PP", "IdxA")), "J": dash(AT(k + "_PP", "IdxB")), "M": dash(SEL(k + "_MS")), "N": situ(SEL(k + "_PMS"), SEL(k + "_SMS"), SEL(k + "_PS"), SEL(k + "_MS"))}
        v["E"] = f'=IF(AND({num("C")},{num("D")}),IF(C{rr}=0,"—",D{rr}/C{rr}-1),"—")'
        v["H"] = f'=IF(AND({num("F")},{num("G")}),G{rr}-F{rr},"—")'
        v["K"] = f'=IF(AND({num("I")},{num("J")}),(J{rr}-I{rr})*100,"—")'
        v["L"] = f'=IF(AND({num("I")},{num("J")}),IF(I{rr}=0,"—",IF(J{rr}=I{rr},"Igual",IF(J{rr}>I{rr},"Piorou","Melhorou"))),"—")'
        if kind == "sub":
            v["O"] = "=" + acc(k + "_FT", k + "_PS", "F"); v["P"] = "=" + acc(k + "_FT", k + "_PS", "P")
            v["R"] = dash(SEL(k + "_FY")); v["S"] = f'=IF({SEL(k+"_DY")}="","—",{SEL(k+"_DY")}*100)'
            sub_rows.append(rr)
        else:
            v["O"] = "=" + "+".join(f"O{r_}" for r_ in sub_rows); v["P"] = "=" + "+".join(f"P{r_}" for r_ in sub_rows)
            v["R"] = "—"; v["S"] = "—"
        v["Q"] = f'=IF(O{rr}=0,"—",P{rr}/O{rr})'
    else:
        v = {"F": dash(AT(k + "_PT", "IdxA")), "G": dash(AT(k + "_PT", "IdxB")), "M": dash(SEL(k + "_MT")), "N": situ(SEL(k + "_PM"), SEL(k + "_SM"), SEL(k + "_PT"), SEL(k + "_MT")),
             "P": "=" + acc(None, k + "_PT", "P")}
        v["H"] = f'=IF(AND({num("F")},{num("G")}),G{rr}-F{rr},"—")'
    for col, val in v.items():
        cc = ws[f"{col}{rr}"]; cc.value = val; cc.font = f(10, kind == "tot"); cc.alignment = Alignment(horizontal="center")
    fm = {"C": "#,##0.0", "D": "#,##0.0", "E": "+0.0%;-0.0%;0.0%", "F": "#,##0.00", "G": "#,##0.00", "H": "+#,##0.00;-#,##0.00;0.00",
          "I": "0.00%", "J": "0.00%", "K": "+0.00;-0.00;0.00", "M": "#,##0.00", "O": "#,##0.0", "P": "#,##0.0", "Q": "0.00%",
          "R": "+0.0%;-0.0%;0.0%", "S": "+0.00;-0.00;0.00"}
    for col, x in fm.items(): ws[f"{col}{rr}"].number_format = x
    for col in "DGJOR": ws[f"{col}{rr}"].font = f(10, True)
    # cores
    ws.conditional_formatting.add(f"E{rr}", FormulaRule(formula=[f'AND({num("E")},E{rr}>0)'], font=Font(name=FONT, color=GOOD_F, bold=True)))
    ws.conditional_formatting.add(f"E{rr}", FormulaRule(formula=[f'AND({num("E")},E{rr}<0)'], font=Font(name=FONT, color=BAD_F, bold=True)))
    ws.conditional_formatting.add(f"H{rr}", FormulaRule(formula=[f'AND({num("H")},{num("F")},F{rr}<>0,H{rr}/F{rr}>0)'], font=Font(name=FONT, color=BAD_F, bold=True)))
    ws.conditional_formatting.add(f"H{rr}", FormulaRule(formula=[f'AND({num("H")},{num("F")},F{rr}<>0,H{rr}/F{rr}<0)'], font=Font(name=FONT, color=GOOD_F, bold=True)))
    status_cf(ws, f"K{rr}:L{rr}", f"$L${rr}")
    for val, fc, bc in (("acima", BAD_F, BAD_B), ("abaixo", GOOD_F, GOOD_B), ("limite", AMB_F, AMB_B)):
        ws.conditional_formatting.add(f"N{rr}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{val}",$N${rr}))'], font=Font(name=FONT, color=fc, bold=True), fill=fill(bc)))
    ws.conditional_formatting.add(f"R{rr}", FormulaRule(formula=[f'AND({num("R")},R{rr}>0)'], font=Font(name=FONT, color=GOOD_F, bold=True)))
    ws.conditional_formatting.add(f"R{rr}", FormulaRule(formula=[f'AND({num("R")},R{rr}<0)'], font=Font(name=FONT, color=BAD_F, bold=True)))
    if kind == "sub":
        for val, fc in (("Melhorou", GOOD_F), ("Piorou", BAD_F)):
            ws.conditional_formatting.add(f"S{rr}", FormulaRule(formula=[f'{SEL(k+"_SY")}="{val}"'], font=Font(name=FONT, color=fc, bold=True)))
        for col in range(2, 20): ws.cell(rr, col).fill = fill(C_SOFT)
    for col in range(2, 20):
        ws.cell(rr, col).border = Border(top=Side(style="medium", color="404040"), bottom=Side(style="medium", color="404040")) if kind == "tot" else Border(bottom=Side(style="hair", color="BFBFBF"))
    ws.row_dimensions[rr].height = 18
    rr += 1
LAST_R = rr - 1
# bordas que separam os blocos (Faturamento | Perda | % de perda | Meta | Acumulado | Ano anterior)
BLK = Side(style="medium", color="404040")
for a_, b_ in [("B", "B")] + [(g[0], g[1]) for g in groups]:
    for r_ in range(4, LAST_R + 1):
        for col in {a_, b_}:
            c = ws[f"{col}{r_}"]
            c.border = Border(left=BLK if col == a_ else c.border.left, right=BLK if col == b_ else c.border.right,
                              top=c.border.top if r_ > 4 or col == "B" else BLK, bottom=c.border.bottom)
for col in "BCDEFGHIJKLMNOPQRS":   # contorno de cima
    c = ws[f"{col}4"]; c.border = Border(left=c.border.left, right=c.border.right, top=BLK, bottom=c.border.bottom)

# ---------- histórico: mapa de calor da perda em t ao lado da meta (cor = situação na meta no mês) ----------
GHR = GHS[ws.title]
rr += 1
ws.merge_cells(f"B{rr}:O{rr}"); ws[f"B{rr}"] = '="Perda (t) nos 12 meses até "&MesTxt&", comparada com a meta"'; ws[f"B{rr}"].font = f(12, True); rr += 1
ws.merge_cells(f"B{rr}:O{rr}"); ws[f"B{rr}"] = "A cor de cada mês compara a perda daquele mês (em toneladas) com a meta da diretoria:"; ws[f"B{rr}"].font = f(9, color="404040"); rr += 1
for (a_, b_), (t, fc, bc) in zip((("B", "B"), ("C", "D"), ("E", "F")),
                                 (("Abaixo da meta", GOOD_F, GOOD_B), ("No limite (atenção)", AMB_F, AMB_B), ("Acima da meta", BAD_F, BAD_B))):
    if a_ != b_: ws.merge_cells(f"{a_}{rr}:{b_}{rr}")
    c = ws[f"{a_}{rr}"]; c.value = t; c.font = f(9, True, fc); c.alignment = Alignment(horizontal="center", vertical="center")
    for cc in ws[f"{a_}{rr}:{b_}{rr}"][0]: cc.fill = fill(bc)
ws.row_dimensions[rr].height = 18
rr += 2
ws.cell(rr, 2, "Subsidiária").font = f(9, True, "FFFFFF"); ws.cell(rr, 2).fill = fill(C_HEAD)
c = ws.cell(rr, 3, "Meta (t/mês)"); c.font = f(9, True, "FFFFFF"); c.fill = fill("262626"); c.alignment = Alignment(horizontal="center")
for i in range(12):
    c = ws.cell(rr, 4 + i, f"={CALC}!C{GHR+1+i}"); c.font = f(9, True, "FFFFFF"); c.fill = fill(C_HEAD); c.alignment = Alignment(horizontal="center")
ws.row_dimensions[rr].height = 18
rr += 1
HIST_HELP = CI("AB")   # colunas ocultas com a situação na meta de cada mês (usadas só para a cor)
for name, k in [(s_[0], s_[1]) for s_ in SUBS] + [("Total da empresa", "TOT")]:
    ws.cell(rr, 2, name).font = f(10, True)
    ws.cell(rr, 2).border = Border(bottom=Side(style="hair", color="BFBFBF"))
    m1, m2 = f"{CALC}!{gcol[k + '_MS']}{GHR+1}", f"{CALC}!{gcol[k + '_MS']}{GHR+12}"
    c = ws.cell(rr, 3, f'=IFERROR(IF(ABS({m1}-{m2})>0.0001,FIXED({m1},1)&" → "&FIXED({m2},1),{m2}),IFERROR({m2},"—"))')
    c.number_format = "#,##0.0"; c.font = f(10, True); c.alignment = Alignment(horizontal="center", vertical="center"); c.fill = fill("EDEDED")
    c.border = Border(right=Side(style="medium", color="404040"), bottom=Side(style="thin", color="FFFFFF"))
    for i in range(12):
        src = f"{CALC}!{gcol[k + '_PS']}{GHR+1+i}"
        c = ws.cell(rr, 4 + i, f'=IFERROR(IF({src}="","—",{src}),"—")'); c.number_format = '[>=10]#,##0;#,##0.0'; c.font = f(10, k == "TOT"); c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"), bottom=Side(style="thin", color="FFFFFF"))
        hc = L(HIST_HELP + i)
        ws[f"{hc}{rr}"] = f"={CALC}!{gcol[k + '_SMS']}{GHR+1+i}"
        cell = f"{L(4 + i)}{rr}"
        for val, fc, bc in (("Melhorou", GOOD_F, GOOD_B), ("Manteve", AMB_F, AMB_B), ("Piorou", BAD_F, BAD_B)):
            ws.conditional_formatting.add(cell, FormulaRule(formula=[f'${hc}${rr}="{val}"'], font=Font(name=FONT, color=fc, bold=True), fill=fill(bc)))
    if k == "TOT":
        for col in range(2, 16): ws.cell(rr, col).border = Border(top=Side(style="medium", color="404040"), right=ws.cell(rr, col).border.right)
    ws.row_dimensions[rr].height = 20
    rr += 1
ws.merge_cells(f"B{rr}:O{rr}")
ws[f"B{rr}"] = "Meta com seta (ex.: 2,6 → 3,0): a meta mudou dentro desses 12 meses; cada mês é comparado com a meta que valia nele. O % de perda de cada mês está nos painéis."
ws[f"B{rr}"].font = f(9, color="404040"); rr += 1
for col in range(CI("AA"), HIST_HELP + 12): ws.column_dimensions[L(col)].hidden = True
setup_print(ws, f"A1:S{rr-1}")

# =====================================================================
# FINANCEIRO — análise do mês (parte impressa)
# =====================================================================
ws = wsF
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 11; ws.column_dimensions["B"].width = 20
for col in "CDEFGHIJKLMN": ws.column_dimensions[col].width = 13
add_selector(ws, "P2", label="Mês da análise")
ws["P4"] = "Comparar com"; ws["P4"].font = f(9, True)
c = ws["P5"]; c.fill = fill(C_IN); c.number_format = "mmm/yyyy"; c.font = f(11, True); c.alignment = Alignment(horizontal="center")
c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
ws["P6"] = "Vazio = mês anterior."; ws["P6"].font = f(8, color="7F7F7F")
ws["P7"] = "A ordem não importa: a variação vai sempre do mais antigo para o mais recente."; ws["P7"].font = f(8, color="7F7F7F")
dvc2 = DataValidation(type="list", formula1=f"{CALC}!$C${R0}:$C${R1}", allow_blank=True, error="Escolha um mês da lista.", errorTitle="Mês inválido", showErrorMessage=True)
ws.add_data_validation(dvc2); dvc2.add("P5")
ws["Z6"] = f'=IF(P5="",IF(Z2>1,Z2-1,""),MATCH(P5,{CALC}!$A${R0}:$A${R1},0))'
ws["Z7"] = f'=IF(Z6="","—",INDEX({CALC}!$D${R0}:$D${R1},Z6))'
ws["Z8"] = '=IF(Z6="","",MIN(Z2,Z6))'
ws["Z9"] = '=IF(Z6="",Z2,MAX(Z2,Z6))'
ws["Z10"] = f'=IF(Z8="","—",INDEX({CALC}!$D${R0}:$D${R1},Z8))'
ws["Z11"] = f'=INDEX({CALC}!$D${R0}:$D${R1},Z9)'
for nm, ref in (("CompIdx", "$Z$6"), ("CompTxt", "$Z$7"), ("IdxA", "$Z$8"), ("IdxB", "$Z$9"), ("TxtA", "$Z$10"), ("TxtB", "$Z$11")):
    ws.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!{ref}")


def VB(key): return f"INDEX({CR(key)},IdxB)"   # valor do mês mais recente


def VAR(key):  # variação % do mês mais antigo (A) para o mais recente (B)
    a, b = f"INDEX({CR(key)},IdxA)", VB(key)
    return f'=IF(IdxA="","—",IF(OR({a}="",{b}=""),"—",IF({a}=0,"—",{b}/{a}-1)))'


ws.merge_cells("B2:J2"); ws["B2"] = "Análise financeira"; ws["B2"].font = f(20, True)
ws.merge_cells("K2:N2"); ws["K2"] = '=TxtA&"  →  "&TxtB'; ws["K2"].font = f(16, True); ws["K2"].alignment = Alignment(horizontal="right")
ws.merge_cells("B3:N3"); ws["B3"] = '="Faturamos mais, e isso virou resultado? Valores de "&TxtB&"; variações de "&TxtA&" para "&TxtB&" (sempre do mês mais antigo para o mais recente)."'
ws["B3"].font = f(9, color="595959")
anyfin = "+".join(f'COUNT({VB(s[1]+"_FR")},{VB(s[1]+"_CK")})' for s in SUBS)
ws.merge_cells("B4:N4")
ws["B4"] = f'=IF({anyfin}=0,"Nenhum valor financeiro lançado para "&TxtB&". Preencha os campos amarelos a partir da linha 27 para ver a análise.","")'
ws["B4"].font = f(10, True, "9C5700")
fh = ["Subsidiária", "Faturamento (t)", '="Fat. (t) vs "&TxtA', "Faturamento (R$)", '="Fat. (R$) vs "&TxtA', "Preço médio (R$/kg)", '="Preço médio vs "&TxtA',
      "Perda (t)", "Custo material (R$/kg)", "Custo da perda (R$)", "% perda financeira", "Receita − custo da perda (R$)", '="vs "&TxtA']
for i, h in enumerate(fh):
    c = ws.cell(6, 2 + i, h); c.font = f(9, True, "FFFFFF"); c.fill = fill(C_HEAD); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[6].height = 32
for j, (name, key) in enumerate([(s[0], s[1]) for s in SUBS] + [("Total da empresa", "TOT")]):
    r = 7 + j
    ws.cell(r, 2, name).font = f(10, True)
    flds = ["FT", "~FT", "FR", "~FR", "PMED", "~PMED", "PS", None, "CP", "PF", "RL", "~RL"]   # "~X" = variação de X (A → B)
    fmts = ["#,##0.0", "+0.0%;-0.0%;0.0%", '"R$" #,##0', "+0.0%;-0.0%;0.0%", '"R$" #,##0.000', "+0.0%;-0.0%;0.0%", "#,##0.00", '"R$" #,##0.000', '"R$" #,##0', "0.00%", '"R$" #,##0', "+0.0%;-0.0%;0.0%"]
    for i, (fld, fm) in enumerate(zip(flds, fmts)):
        if fld is None:
            v = dash(VB(key + "_CK")) if key != "TOT" else f'=IF(OR({VB("TOT_CP")}="",{VB("TOT_PS")}=""),"—",{VB("TOT_CP")}/({VB("TOT_PS")}*1000))'
        elif fld.startswith("~"):
            v = VAR(f"{key}_{fld[1:]}")
        else:
            v = dash(VB(f"{key}_{fld}"))
        c = ws.cell(r, 3 + i, v); c.number_format = fm; c.font = f(10, key == "TOT"); c.alignment = Alignment(horizontal="center")
        c.border = Border(bottom=Side(style="hair", color="BFBFBF"))
    if key == "TOT":
        for col in range(2, 15): ws.cell(r, col).border = Border(top=Side(style="medium", color="404040"))
    for col in ("D", "F", "H", "N"):
        ws.conditional_formatting.add(f"{col}{r}", FormulaRule(formula=[f'AND(ISNUMBER({col}{r}),{col}{r}>0)'], font=Font(name=FONT, color=GOOD_F, bold=True)))
        ws.conditional_formatting.add(f"{col}{r}", FormulaRule(formula=[f'AND(ISNUMBER({col}{r}),{col}{r}<0)'], font=Font(name=FONT, color=BAD_F, bold=True)))
ws["B13"] = "Como ler este quadro"; ws["B13"].font = f(10, True)
guide = ["Faturamos mais em peso e em dinheiro?  →  compare Faturamento (t) e Faturamento (R$) com o mês de comparação.",
         "Estamos vendendo melhor ou pior?  →  Preço médio (R$/kg) = faturamento em R$ ÷ faturamento em kg (mostra o efeito do mix de clientes e produtos).",
         "Quanto a perda nos custa?  →  Custo da perda (R$) = perda em kg × custo do material perdido (R$/kg).",
         "Quanto da receita a perda consome?  →  % perda financeira = custo da perda ÷ faturamento em R$.",
         "Melhorou de verdade?  →  Receita − custo da perda, comparada com o mês de comparação. (Não é lucro: não considera outros custos.)"]
for i, g in enumerate(guide):
    ws.merge_cells(f"B{14+i}:N{14+i}"); ws[f"B{14+i}"] = "•  " + g; ws[f"B{14+i}"].font = f(9, color="404040")
for r in range(13, 19):
    for col in range(2, 15): ws.cell(r, col).fill = fill(C_SOFT)
ws["B20"] = "Total da empresa: só aparece quando as 4 subsidiárias tiverem valores lançados no mês."; ws["B20"].font = f(8, color="7F7F7F", italic=True)
setup_print(ws, "A1:N20")

# =====================================================================
# INSTRUÇÕES
# =====================================================================
ws = wsI
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 110
txt = [
    (f"Planilha de Faturamento e Perdas · {EMPRESA} (dados fictícios de exemplo)", f(18, True)),
    ("", None),
    ("Todo mês, ao receber os relatórios", f(13, True)),
    ("1. Abra a aba Lançamentos e localize a linha do mês (os meses já estão carregados até dez/2028).", f(10)),
    ("2. Digite o faturamento e as perdas em kg, como vêm no relatório. A planilha converte para toneladas sozinha.", f(10)),
    ("3. O faturamento fica no bloco azul e as perdas no bloco laranja. Célula em branco = ainda não lançado (não é zero). Células cinzas = sem dado separado (G1/G2 antes de out/2025).", f(10)),
    ("4. Pronto: todas as páginas passam a mostrar o último mês lançado, com gráficos, cores e frases de destaque atualizados.", f(10)),
    ("", None),
    ("Para apresentar à diretoria", f(13, True)),
    ("Ordem sugerida: 1) Painel Consolidado (a empresa no mês) → 2) Onde focar (onde está a perda e para onde ela está indo) → 3) painéis das subsidiárias, se houver perguntas sobre um setor.", f(10)),
    ("As 3 frases de \"Destaques\" no topo de cada página já resumem o mês e podem ser usadas como roteiro da fala.", f(10)),
    ("O Resumo é uma tabela de consulta: funciona melhor impresso ou enviado como anexo do que projetado.", f(10)),
    ("", None),
    ("Para ver outro mês", f(13, True)),
    ("Cada painel, o Onde focar, o Resumo e o Financeiro têm o seu próprio seletor de mês (célula amarela à direita, fora da área de impressão). Vazio = último mês lançado.", f(10)),
    ("No Resumo e no Financeiro, a célula \"Comparar com\" escolhe o mês de comparação (vazio = mês anterior). Ex.: mai/2026 × mai/2025.", f(10)),
    ("Antes de exportar o PDF, confira o mês no cabeçalho de cada página: cada aba mostra o mês escolhido nela.", f(10)),
    ("", None),
    ("Para exportar em PDF", f(13, True)),
    ("1. Clique na aba Painel Consolidado, segure Shift e clique na aba Resumo (seleciona Consolidado + Onde focar + 4 painéis + Resumo, já na ordem da apresentação).", f(10)),
    ("2. Arquivo → Exportar → Criar PDF/XPS. Cada aba vira 1 página A4 deitada (7 páginas).", f(10)),
    ("3. Para tirar uma página, segure Ctrl e clique na aba antes de exportar. O Financeiro só entra se for incluído com Ctrl + clique (use quando estiver preenchido).", f(10)),
    ("4. Depois de exportar, clique em qualquer outra aba para desfazer a seleção (evita digitar em várias abas ao mesmo tempo).", f(10)),
    ("", None),
    ("Como ler os indicadores", f(13, True)),
    ("% de perda", f(10, True)),
    ("•  Perda ÷ faturamento. É o indicador principal.", f(10)),
    ("•  p.p. (pontos percentuais) = diferença entre dois percentuais. Ex.: o % de perda caiu de 1,0% para 0,5% = queda de 0,5 p.p.", f(10)),
    ("", None),
    ("Comparação com o mês anterior", f(10, True)),
    ("•  Feita pelo % de perda, que é justo entre meses com volumes diferentes.", f(10)),
    ("•  Verde = melhorou.", f(10, True, GOOD_F)),
    ("•  Vermelho = piorou (mesmo que a diferença seja pequena).", f(10, True, BAD_F)),
    ("", None),
    ("Comparação com a meta", f(10, True)),
    ("•  A meta é o limite máximo de perda do mês, em toneladas, definido pela diretoria (aba Metas).", f(10)),
    ("•  Verde = perda abaixo da meta.", f(10, True, GOOD_F)),
    ("•  Âmbar = perda ainda abaixo da meta, mas a menos de 1% do limite. Merece atenção.", f(10, True, AMB_F)),
    ("•  Vermelho = perda acima da meta, mesmo que pouco.", f(10, True, BAD_F)),
    ("", None),
    ("Tendência (aba Onde focar)", f(10, True)),
    ("•  Compara a média do % de perda dos últimos 3 meses com a média dos 3 meses anteriores.", f(10)),
    ("•  Não é a comparação com o mês anterior: olhar 3 meses evita reagir a um mês isolado.", f(10)),
    ("", None),
    ("Prioridade (aba Onde focar). Uma etapa \"pesa muito\" quando responde por 10% ou mais de toda a perda dos últimos 12 meses.", f(10, True)),
    ("•  Agir agora: a etapa pesa muito na perda total e está piorando. Investigar a causa já.", f(10)),
    ("•  Padronizar: a etapa pesa muito na perda total e está melhorando. Descobrir o que deu certo e transformar em padrão.", f(10)),
    ("•  Acompanhar: a etapa pesa pouco, mas está piorando. Observar mês a mês antes que cresça.", f(10)),
    ("•  Manter: a etapa pesa pouco e está estável ou melhorando. Seguir como está.", f(10)),
    ("", None),
    ("Ganho potencial (aba Onde focar)", f(10, True)),
    ("•  Estimativa de quantas toneladas a etapa teria deixado de perder se, nos meses ruins, tivesse ficado no nível dos seus 3 melhores meses. Não é meta.", f(10)),
]
for i, (t, ft) in enumerate(txt, 2):
    ws[f"B{i}"] = t
    if ft: ws[f"B{i}"].font = ft
    ws[f"B{i}"].alignment = Alignment(wrap_text=True, vertical="top", indent=2 if t.startswith("•") else 0)
setup_print(ws, f"A1:B{len(txt)+1}", landscape=False)

for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            if c.font and c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size, bold=c.font.bold, color=c.font.color, italic=c.font.italic)
for col, w in (("A", 37.7), ("C", 17.4), ("M", 60.6), ("N", 9.1)): wsL.column_dimensions[col].width = w
wsC.sheet_state = "hidden"   # aba de cálculo oculta (Página Inicial → Formatar → Reexibir, se precisar)
wb.calculation.fullCalcOnLoad = True
wb.active = 2
for w in wb.worksheets: w.sheet_view.tabSelected = (w.title == P1)
wb.save(OUT)
print("ok", nxt[0], "colunas de cálculo")

# --- pós-processamento: formato fixo nos rótulos dos gráficos (sourceLinked=0) ---
import zipfile, shutil, re as _re
tmp = OUT + ".tmp"
with zipfile.ZipFile(OUT) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        data = zi.read(it.filename)
        if it.filename.startswith("xl/charts/chart"):
            t = data.decode("utf-8")
            t = _re.sub(r'<numFmt formatCode="([^"]*)"/>', r'<numFmt formatCode="\1" sourceLinked="0"/>', t)
            t = t.replace("<a:r><a:t></a:t></a:r>", "")
            data = t.encode("utf-8")
        zo.writestr(it, data)
shutil.move(tmp, OUT)
print("rótulos ok")
