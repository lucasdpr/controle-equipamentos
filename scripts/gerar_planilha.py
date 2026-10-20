"""Gera controle_equipamentos.xlsx (planilha própria, separada do arquivo do e-mail).

Fontes diárias (anexos do e-mail "Relatórios GLA", lidos pelo robô):
- VIDA MOLDE.xlsx  -> acumulado de corridas por veio (C..H) e nº do molde.
- CORTESEQ.xls     -> aba DADOS DA PRODUÇÃO: toneladas reais por MCC por dia.

Regras:
- Molde: corridas (acumulado do último dia importado, por veio). Meta fixa por veio.
- Oscilador: dias em máquina (entrada conta, saída não; em aberto conta até hoje).
- Componentes: toneladas reais do MCC4 nos dias entre entrada e saída.
- Falta = meta - acumulado. % = acumulado / meta. Sem previsão.

Abas: Painel, Resumo, Moldes, TrocasEquip*, Paradas*, Historico, Producao, Trocas,
Alertas, Regras, Auxiliar (oculta).   * = preenchidas à mão; as demais pelo robô.
"""

import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

SAIDA = Path(__file__).resolve().parent.parent / "controle_equipamentos.xlsx"
ULTIMA = 1000   # linhas com fórmula pré-preenchida
LIM = 5000      # intervalo usado nas fórmulas de Historico/Trocas/Alertas

AZUL = "1F4E78"
AZUL_CLARO = "DDEBF7"
CINZA = "595959"
CINZA_CLARO = "F2F2F2"
AMARELO = "FFF2CC"
VERDE = "C6EFCE"
AMBAR = "FFEB9C"
VERMELHO = "FFC7CE"

fino = Side(style="thin", color="BFBFBF")
BORDA = Border(left=fino, right=fino, top=fino, bottom=fino)


def cabecalho(ws, colunas, cor=AZUL, altura=32):
    ws.append(colunas)
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor=cor)
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celula.border = BORDA
    ws.row_dimensions[1].height = altura
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def larguras(ws, valores):
    for letra, largura in valores.items():
        ws.column_dimensions[letra].width = largura


def bordas(ws, faixa):
    for linha in ws[faixa]:
        for c in linha:
            c.border = BORDA


def preencher(ws, faixa, cor):
    for linha in ws[faixa]:
        for c in linha:
            c.fill = PatternFill("solid", fgColor=cor)


def cores_situacao(ws, faixa):
    ws.conditional_formatting.add(faixa, CellIsRule(operator="equal", formula=['"Além da meta"'],
                                  fill=PatternFill("solid", bgColor=VERMELHO), font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(faixa, CellIsRule(operator="equal", formula=['"Atenção"'],
                                  fill=PatternFill("solid", bgColor=AMBAR), font=Font(bold=True, color="9C5700")))
    ws.conditional_formatting.add(faixa, CellIsRule(operator="equal", formula=['"OK"'],
                                  fill=PatternFill("solid", bgColor=VERDE), font=Font(bold=True, color="006100")))


def barra_percentual(ws, faixa):
    ws.conditional_formatting.add(
        faixa,
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5", showValue=True),
    )


wb = Workbook()

# =====================================================================
# Resumo (uma linha por equipamento/tag) - tudo automático, igual ao molde
# =====================================================================
rs = wb.active
rs.title = "Resumo"
cabecalho(
    rs,
    ["Tag / Equipamento", "Tipo", "Unidade", "MCC", "Meta", "Desde", "Acumulado inicial",
     "Acumulado", "Falta", "%", "Situação", "Veio"],
)
TIPOS = ["Molde", "Oscilador", "Bender", "Bow", "Straightener R1", "Straightener R2", "Horizontal"]
META_PADRAO = {"Molde": 1100, "Oscilador": 730, "Bender": 1100000, "Bow": 1900000,
               "Straightener R1": 1700000, "Straightener R2": 1700000, "Horizontal": 3300000}
equip = (
    [("Molde G", "Molde", "G"), ("Molde H", "Molde", "H"), ("Oscilador", "Oscilador", None),
     ("Bender", "Bender", None)]
    + [(f"Bow {n}", "Bow", None) for n in range(1, 6)]
    + [("Straightener R1", "Straightener R1", None), ("Straightener R2", "Straightener R2", None)]
    + [(f"Horizontal {n:02d}", "Horizontal", None) for n in range(1, 11)]
)
N = len(equip)          # 21 equipamentos já listados
RS_ULT = 41             # linhas com fórmula (cabem até 40 equipamentos)
INICIO_DADOS = datetime.date(2026, 8, 1)  # primeiro dia com produção importada (provisório)
ult_hist = f"SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=L{{i}})*Historico!$A$2:$A${LIM}))"
for i in range(2, RS_ULT + 1):
    k = i - 2
    if k < N:
        nome, tipo, veio = equip[k]
        molde = tipo == "Molde"
        rs.cell(i, 1, nome)
        rs.cell(i, 2, tipo)
        rs.cell(i, 4, None if molde else 4)
        rs.cell(i, 5, META_PADRAO[tipo])
        rs.cell(i, 6, None if molde else INICIO_DADOS)
        rs.cell(i, 12, veio)
    rs[f"C{i}"] = f'=IF(B{i}="","",IF(B{i}="Molde","corridas",IF(B{i}="Oscilador","dias","t")))'
    uh = ult_hist.format(i=i)
    rs[f"H{i}"] = (
        f'=IF(A{i}="","",IF(C{i}="corridas",'
        f"SUMIFS(Historico!$E:$E,Historico!$B:$B,L{i},Historico!$A:$A,{uh}),"
        f'N(G{i})+IF(F{i}="",0,IF(C{i}="dias",MAX(0,TODAY()-F{i}),'
        f'IF(D{i}=2,SUMIFS(Producao!$D:$D,Producao!$A:$A,">="&F{i}),'
        f'IF(D{i}=3,SUMIFS(Producao!$E:$E,Producao!$A:$A,">="&F{i}),'
        f'SUMIFS(Producao!$B:$B,Producao!$A:$A,">="&F{i})))))))'
    )
    rs[f"I{i}"] = f'=IF(OR(A{i}="",E{i}=""),"",E{i}-H{i})'
    rs[f"J{i}"] = f'=IF(OR(A{i}="",N(E{i})<=0),"",H{i}/E{i})'
    rs[f"K{i}"] = f'=IF(I{i}="","",IF(I{i}<0,"Além da meta",IF(J{i}>=Regras!$B$3,"Atenção","OK")))'
    for col, fmt in (("E", "#,##0.##"), ("G", "#,##0.##"), ("H", "#,##0.##"), ("I", "#,##0.##"),
                     ("F", "dd/mm/yyyy"), ("J", "0.0%")):
        rs[f"{col}{i}"].number_format = fmt
    for c in rs[i]:
        c.border = BORDA
        c.alignment = Alignment(horizontal="left" if c.column == 1 else
                                "center" if c.column in (2, 3, 4, 6, 11, 12) else "right")
preencher(rs, f"A2:B{RS_ULT}", AMARELO)     # tag e tipo
preencher(rs, f"D2:G{RS_ULT}", AMARELO)     # MCC, meta, desde, acumulado inicial
barra_percentual(rs, f"J2:J{RS_ULT}")
cores_situacao(rs, f"K2:K{RS_ULT}")
rs[f"A{RS_ULT + 3}"] = "Cada linha é um equipamento com a SUA tag e a SUA meta. Digite a tag real em 'Tag / Equipamento' e a meta de cada um."
rs[f"A{RS_ULT + 4}"] = "Molde: automático pelo relatório (corridas do veio G ou H). Oscilador e componentes: automático também, a partir da data em 'Desde'."
rs[f"A{RS_ULT + 5}"] = "Acumulado = 'Acumulado inicial' + toneladas da MCC escolhida (ou dias, no oscilador) desde a data 'Desde'."
rs[f"A{RS_ULT + 6}"] = "ATENÇÃO: metas padrão por tipo e 'Desde' = 01/08/2026 são PROVISÓRIOS. Digite a meta, a data de instalação e o acumulado inicial reais de cada tag."
rs[f"A{RS_ULT + 7}"] = "Trocou um equipamento? Registre a saída em 'TrocasEquip', depois digite a nova data em 'Desde' e zere o 'Acumulado inicial'."
for r_ in range(RS_ULT + 3, RS_ULT + 8):
    rs[f"A{r_}"].font = Font(italic=True, color="9C0006" if r_ == RS_ULT + 6 else CINZA)
larguras(rs, {"A": 22, "B": 17, "C": 10, "D": 6, "E": 13, "F": 12, "G": 16, "H": 14, "I": 14, "J": 15, "K": 14, "L": 7})
dv_tp = DataValidation(type="list", formula1='"' + ",".join(TIPOS) + '"', allow_blank=True)
dv_tp.add(f"B2:B{RS_ULT}")
rs.add_data_validation(dv_tp)
dv_mcc = DataValidation(type="list", formula1='"2,3,4"', allow_blank=True)
dv_mcc.add(f"D2:D{RS_ULT}")
rs.add_data_validation(dv_mcc)
dv_ds = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                       error="Digite uma data válida (dd/mm/aaaa).", errorTitle="Data inválida")
dv_ds.add(f"F2:F{RS_ULT}")
rs.add_data_validation(dv_ds)
rs.sheet_properties.tabColor = AZUL

# =====================================================================
# Moldes (todos os veios)
# =====================================================================
ml = wb.create_sheet("Moldes")
cabecalho(
    ml,
    ["Veio", "MCC", "Molde atual", "Meta", "Acumulado", "Falta", "%", "Situação",
     "Trocas (qtd)", "Corridas médias na saída", "Última troca", "Dias parado", "Dados de"],
    altura=40,
)
veios = [("C", 2), ("D", 2), ("E", 3), ("F", 3), ("G", 4), ("H", 4)]
for i, (veio, mcc) in enumerate(veios, start=2):
    ml.append([veio, f"MCC{mcc}"])
    ml[f"M{i}"] = f"=SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=A{i})*Historico!$A$2:$A${LIM}))"
    base = f"Historico!$B:$B,A{i},Historico!$A:$A,M{i}"
    ml[f"C{i}"] = f"=IF(M{i}=0,\"\",SUMIFS(Historico!$C:$C,{base}))"
    ml[f"D{i}"] = f"=IF(M{i}=0,\"\",SUMIFS(Historico!$D:$D,{base}))"
    ml[f"E{i}"] = f"=IF(M{i}=0,\"\",SUMIFS(Historico!$E:$E,{base}))"
    ml[f"F{i}"] = f'=IF(M{i}=0,"",D{i}-E{i})'
    ml[f"G{i}"] = f'=IF(M{i}=0,"",IF(N(D{i})>0,E{i}/D{i},""))'
    ml[f"H{i}"] = (
        f'=IF(M{i}=0,"",IF(F{i}<0,"Além da meta",IF(G{i}>=Regras!$B$3,"Atenção","OK")))'
    )
    ml[f"I{i}"] = f"=COUNTIFS(Trocas!$B:$B,A{i})"
    ml[f"J{i}"] = f'=IFERROR(AVERAGEIFS(Trocas!$E:$E,Trocas!$B:$B,A{i}),"-")'
    ml[f"K{i}"] = f'=IF(I{i}=0,"-",SUMPRODUCT(MAX((Trocas!$B$2:$B${LIM}=A{i})*Trocas!$A$2:$A${LIM})))'
    ml[f"L{i}"] = f"=SUMIFS(Paradas!$E:$E,Paradas!$C:$C,A{i})"
    ml[f"M{i}"].number_format = "dd/mm/yyyy;;\"-\""
    ml[f"K{i}"].number_format = "dd/mm/yyyy"
    ml[f"G{i}"].number_format = "0.0%"
    ml[f"J{i}"].number_format = "#,##0"
    ml[f"D{i}"].number_format = "#,##0"
    ml[f"E{i}"].number_format = "#,##0"
    ml[f"F{i}"].number_format = "#,##0"
bordas(ml, "A2:M7")
for linha in ml["A2:M7"]:
    for c in linha:
        c.alignment = Alignment(horizontal="center")
barra_percentual(ml, "G2:G7")
cores_situacao(ml, "H2:H7")
ml["A9"] = "Mostra todos os veios (C a H) com o último dado importado pelo robô. Corridas médias na saída = média das trocas já registradas."
ml["A9"].font = Font(italic=True, color=CINZA)
larguras(ml, {"A": 7, "B": 8, "C": 12, "D": 10, "E": 12, "F": 10, "G": 16, "H": 14,
              "I": 12, "J": 18, "K": 14, "L": 12, "M": 13})
ml.sheet_properties.tabColor = AZUL

# =====================================================================
# TrocasEquip (manual): registro das saídas de oscilador e componentes
# =====================================================================
te = wb.create_sheet("TrocasEquip")
cabecalho(te, ["Data da saída", "Equipamento", "Acumulado na saída", "Unidade", "Motivo", "Observação"], cor="BF8F00")
for r in range(2, ULTIMA + 1):
    te[f"A{r}"].number_format = "dd/mm/yyyy"
preencher(te, f"A2:F{ULTIMA}", AMARELO)
larguras(te, {"A": 15, "B": 20, "C": 20, "D": 10, "E": 45, "F": 40})
dv_eq = DataValidation(type="list", formula1="=Resumo!$A$2:$A$200", allow_blank=True)
dv_eq.add(f"B2:B{ULTIMA}")
te.add_data_validation(dv_eq)
dv_dt = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                       error="Digite uma data válida (dd/mm/aaaa).", errorTitle="Data inválida")
dv_dt.add(f"A2:A{ULTIMA}")
te.add_data_validation(dv_dt)
te.sheet_properties.tabColor = "BF8F00"

# =====================================================================
# Paradas (manual)
# =====================================================================
pa = wb.create_sheet("Paradas")
cabecalho(pa, ["Início", "Fim", "Veio / equipamento", "Motivo", "Dias parado", "Observação"], cor="BF8F00")
for r in range(2, ULTIMA + 1):
    pa[f"E{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",B{r}-A{r})'
    pa[f"A{r}"].number_format = "dd/mm/yyyy"
    pa[f"B{r}"].number_format = "dd/mm/yyyy"
preencher(pa, f"A2:D{ULTIMA}", AMARELO)
preencher(pa, f"F2:F{ULTIMA}", AMARELO)
larguras(pa, {"A": 13, "B": 13, "C": 20, "D": 45, "E": 13, "F": 40})
dv_vp = DataValidation(type="list", formula1='"C,D,E,F,G,H"', allow_blank=True, showErrorMessage=False)
dv_vp.add(f"C2:C{ULTIMA}")
pa.add_data_validation(dv_vp)
pa.sheet_properties.tabColor = "BF8F00"

# =====================================================================
# Abas do robô
# =====================================================================
hs = wb.create_sheet("Historico")
cabecalho(hs, ["Data", "Veio", "Molde nº", "Meta (vida progr.)", "Acumulado (corridas)", "Observação"], cor=CINZA)
larguras(hs, {"A": 13, "B": 7, "C": 11, "D": 18, "E": 20, "F": 70})
hs.sheet_properties.tabColor = CINZA

pr = wb.create_sheet("Producao")
cabecalho(pr, ["Data", "MCC4 real (t)", "MCC4 prog (t)", "MCC2 real (t)", "MCC3 real (t)", "Arquivo"], cor=CINZA)
larguras(pr, {"A": 13, "B": 15, "C": 15, "D": 15, "E": 15, "F": 36})
pr.sheet_properties.tabColor = CINZA

tr = wb.create_sheet("Trocas")
cabecalho(
    tr,
    ["Data", "Veio", "Molde que saiu", "Vida programada", "Corridas na saída",
     "Molde que entrou", "Motivo", "Conferir", "Origem"],
    cor=CINZA,
)
larguras(tr, {"A": 13, "B": 7, "C": 15, "D": 16, "E": 17, "F": 16, "G": 60, "H": 45, "I": 10})
tr.conditional_formatting.add(f"H2:H{LIM}", FormulaRule(formula=["LEN($H2)>0"], fill=PatternFill("solid", bgColor=AMBAR)))
tr.sheet_properties.tabColor = CINZA

al = wb.create_sheet("Alertas")
cabecalho(al, ["Data/hora", "Tipo", "Item", "Detalhe", "Resolvido"], cor="C00000")
larguras(al, {"A": 18, "B": 28, "C": 16, "D": 90, "E": 11})
al.conditional_formatting.add(
    f"A2:E{LIM}", FormulaRule(formula=['AND($A2<>"",$E2<>"Sim")'], fill=PatternFill("solid", bgColor=VERMELHO))
)
dv_res = DataValidation(type="list", formula1='"Sim"', allow_blank=True)
dv_res.add(f"E2:E{LIM}")
al.add_data_validation(dv_res)
al.sheet_properties.tabColor = "C00000"

# =====================================================================
# Regras
# =====================================================================
rg = wb.create_sheet("Regras")
rg["A1"], rg["B1"] = "Parâmetro", "Valor"
for celula in (rg["A1"], rg["B1"]):
    celula.font = Font(bold=True, color="FFFFFF")
    celula.fill = PatternFill("solid", fgColor=AZUL)
rg["A3"], rg["B3"] = "Limiar de atenção (%)", 0.8
rg["B3"].number_format = "0%"
rg["B3"].fill = PatternFill("solid", fgColor=AMARELO)
rg["A5"] = "Como a planilha calcula"
rg["A5"].font = Font(bold=True, size=12)
textos = [
    "Molde (corridas): acumulado do último dia em Historico, por veio. O robô grava um registro por data e veio; reenvio do mesmo dia substitui, não soma.",
    "Troca de molde: quando o nº do molde do veio muda de um dia para o outro, o robô registra em Trocas (molde que saiu, corridas na saída, molde que entrou, motivo).",
    "Oscilador (dias): dias corridos desde a data em Desde (Resumo) + Acumulado inicial.",
    "Componentes (t): toneladas reais do MCC4 (aba Producao) desde a data em Desde (Resumo) + Acumulado inicial. O robô atualiza a produção todo dia.",
    "Falta = meta − acumulado. % = acumulado ÷ meta. Acima de 100% = além da meta. Sem previsão de data de troca.",
    "Alertas: acumulado diminuiu sem troca; veio não andou enquanto a MCC produziu; layout dos anexos mudou. Escolha 'Sim' em Resolvido quando tratar.",
    "Cores das abas: azul = painel e resumos; amarelo/ouro = você digita (Desde e Acumulado inicial no Resumo, TrocasEquip, Paradas); cinza = o robô preenche; vermelho = alertas.",
]
for i, t in enumerate(textos, start=6):
    rg[f"A{i}"] = t
rg["A14"] = "Autor: Lucas Gabriel de Paula Rafael  |  Matrícula cbk3574  |  github.com/lucasdpr/controle-equipamentos"
rg["A14"].font = Font(italic=True, color=CINZA)
rg.sheet_view.showGridLines = False
larguras(rg, {"A": 140, "B": 12})
rg.sheet_properties.tabColor = CINZA

# =====================================================================
# Auxiliar (oculta): dados dos gráficos, últimos 30 dias
# =====================================================================
ax = wb.create_sheet("Auxiliar")
ax.append(["Dia", "Veio G (corridas)", "Veio H (corridas)", "MCC4 real (t)"])
for n in range(30):
    r = n + 2
    ax[f"A{r}"] = f"=TODAY()-{29 - n}"
    ax[f"A{r}"].number_format = "dd/mm"
    for col, veio in (("B", "G"), ("C", "H")):
        ax[f"{col}{r}"] = (
            f'=IF(COUNTIFS(Historico!$A:$A,$A{r},Historico!$B:$B,"{veio}")=0,NA(),'
            f'SUMIFS(Historico!$E:$E,Historico!$A:$A,$A{r},Historico!$B:$B,"{veio}"))'
        )
    ax[f"D{r}"] = f"=SUMIFS(Producao!$B:$B,Producao!$A:$A,$A{r})"
ax.sheet_state = "hidden"

# =====================================================================
# Painel (primeira aba)
# =====================================================================
pn = wb.create_sheet("Painel", 0)
pn.sheet_view.showGridLines = False
larguras(pn, {"A": 20, "B": 11, "C": 14, "D": 14, "E": 14, "F": 18, "G": 14, "H": 18})

pn.merge_cells("A1:H1")
pn["A1"] = "Controle de Vida dos Equipamentos  |  MCC4"
pn["A1"].font = Font(size=18, bold=True, color="FFFFFF")
pn["A1"].fill = PatternFill("solid", fgColor=AZUL)
pn["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
pn.row_dimensions[1].height = 34

pn.merge_cells("A2:H2")
def _data_txt(ref):
    # sem TEXT(): os codigos de formato mudam por idioma (aaaa no Excel em portugues)
    return f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'


_ult = f"MAX(Historico!$A$2:$A${LIM})"
pn["A2"] = (
    f'="Dados do robô até: "&IF({_ult}=0,"nenhum e-mail processado ainda",{_data_txt(_ult)})'
    f'&"   |   Hoje: "&{_data_txt("TODAY()")}'
)
pn["A2"].font = Font(italic=True, color=CINZA)
pn["A2"].alignment = Alignment(indent=1)
pn.merge_cells("A3:H3")
pn["A3"] = "Lucas Gabriel de Paula Rafael  |  Matrícula cbk3574  |  github.com/lucasdpr/controle-equipamentos"
pn["A3"].font = Font(size=8, italic=True, color="808080")
pn["A3"].alignment = Alignment(indent=1)

# Indicadores
kpis = [
    ("A4:B4", "A5:B5", "Além da meta", f'=COUNTIF(Resumo!$K$2:$K${RS_ULT},"Além da meta")', VERMELHO, "9C0006"),
    ("C4:D4", "C5:D5", "Atenção", f'=COUNTIF(Resumo!$K$2:$K${RS_ULT},"Atenção")', AMBAR, "9C5700"),
    ("E4:F4", "E5:F5", "OK", f'=COUNTIF(Resumo!$K$2:$K${RS_ULT},"OK")', VERDE, "006100"),
    ("G4:H4", "G5:H5", "Alertas abertos",
     f'=COUNTA(Alertas!$A$2:$A${LIM})-COUNTIFS(Alertas!$E$2:$E${LIM},"Sim")', AZUL_CLARO, AZUL),
]
for topo, valor, titulo, formula, fundo, cor in kpis:
    pn.merge_cells(topo)
    pn.merge_cells(valor)
    celula_t = pn[topo.split(":")[0]]
    celula_v = pn[valor.split(":")[0]]
    celula_t.value = titulo
    celula_t.font = Font(bold=True, color=cor)
    celula_t.alignment = Alignment(horizontal="center")
    celula_v.value = formula
    celula_v.font = Font(size=26, bold=True, color=cor)
    celula_v.alignment = Alignment(horizontal="center", vertical="center")
    preencher(pn, topo, fundo)
    preencher(pn, valor, fundo)
pn.row_dimensions[5].height = 42

# Tabela do painel (espelha o Resumo)
PN = N + 3  # linhas do painel (equipamentos + 3 de folga)
cab = ["Tag / Equipamento", "Tipo", "Meta", "Acumulado", "Falta", "%", "Situação", "Desde"]
for j, t in enumerate(cab, start=1):
    c = pn.cell(row=7, column=j, value=t)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDA
origem = {1: "A", 2: "B", 3: "E", 4: "H", 5: "I", 6: "J", 7: "K", 8: "F"}
for k in range(PN):
    r = 8 + k
    for j, col in origem.items():
        ref = f"Resumo!{col}{k + 2}"
        if j == 8:
            f = f'=IF(OR(Resumo!A{k + 2}="",{ref}=""),"",{ref})'
        else:
            f = f'=IF(Resumo!A{k + 2}="","",{ref})'
        c = pn.cell(row=r, column=j, value=f)
        c.border = BORDA
        c.alignment = Alignment(horizontal="left" if j == 1 else "center" if j in (2, 7, 8) else "right")
        if j in (3, 4, 5):
            c.number_format = "#,##0"
        if j == 6:
            c.number_format = "0.0%"
        if j == 8:
            c.number_format = "dd/mm/yyyy"
barra_percentual(pn, f"F8:F{7 + PN}")
cores_situacao(pn, f"G8:G{7 + PN}")

# Gráficos
c1 = BarChart()
c1.type = "bar"
c1.title = "% da vida consumida"
c1.style = 10
c1.add_data(Reference(pn, min_col=6, min_row=7, max_row=7 + PN), titles_from_data=True)
c1.set_categories(Reference(pn, min_col=1, min_row=8, max_row=7 + PN))
c1.y_axis.number_format = "0%"
c1.y_axis.scaling.min = 0
c1.legend = None
c1.height, c1.width = 12.5, 12
pn.add_chart(c1, f"A{9 + PN}")

c2 = LineChart()
c2.title = "Corridas dos moldes G e H (30 dias)"
c2.style = 12
c2.add_data(Reference(ax, min_col=2, max_col=3, min_row=1, max_row=31), titles_from_data=True)
c2.set_categories(Reference(ax, min_col=1, min_row=2, max_row=31))
c2.y_axis.scaling.min = 0
c2.height, c2.width = 6, 12
pn.add_chart(c2, f"E{9 + PN}")

c3 = BarChart()
c3.type = "col"
c3.title = "Toneladas reais do MCC4 por dia (30 dias)"
c3.style = 10
c3.add_data(Reference(ax, min_col=4, min_row=1, max_row=31), titles_from_data=True)
c3.set_categories(Reference(ax, min_col=1, min_row=2, max_row=31))
c3.legend = None
c3.height, c3.width = 6, 12
pn.add_chart(c3, f"E{21 + PN}")

pn.sheet_properties.tabColor = AZUL
pn.page_setup.orientation = "landscape"
pn.page_setup.fitToWidth = 1
pn.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

wb.active = 0
wb.save(SAIDA)
print(f"Gerado: {SAIDA}")
