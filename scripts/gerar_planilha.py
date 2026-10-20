"""Gera controle_equipamentos.xlsx (planilha própria, separada do arquivo do e-mail).

Autor: Lucas Gabriel de Paula Rafael | Matrícula cbk3574 | github.com/lucasdpr/controle-equipamentos

Fontes diárias (anexos do e-mail "Relatórios GLA", lidos pelo robô atualizar_controle.vbs):
- VIDA MOLDE.xlsx  -> acumulado de corridas por veio (C..H) e nº do molde   -> aba Historico
- CORTESEQ.xls     -> DADOS DA PRODUÇÃO: toneladas reais por MCC por dia      -> aba Producao

Regras:
- Molde: corridas (acumulado do último dia importado, pelo veio).
- Oscilador: dias corridos desde "Desde" + acumulado inicial.
- Componentes: toneladas reais da MCC escolhida desde "Desde" + acumulado inicial.
- Falta = meta - acumulado. % = acumulado / meta. Sem previsão.

As abas Historico, Producao, Trocas e Alertas são escritas pelo robô: nome, cabeçalho na
linha 1 e ordem das colunas NÃO podem mudar sem mudar o robô junto.
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
MARCA = "Lucas Gabriel de Paula Rafael  |  Matrícula cbk3574  |  github.com/lucasdpr/controle-equipamentos"
ULTIMA = 1000   # linhas preparadas nas abas manuais
LIM = 5000      # intervalo usado nas fórmulas das abas do robô
EQ_ULT = 41     # aba Equipamentos: linhas 2..41 (até 40 equipamentos)
INICIO_DADOS = datetime.date(2026, 8, 1)  # primeiro dia com produção importada (provisório)

# ---- Paleta ----------------------------------------------------------------
NAVY = "1F3864"
AZUL = "2E75B6"
AZUL_CLARO = "DDEBF7"
CINZA = "595959"
CINZA_CLARO = "F2F2F2"
OURO = "BF8F00"
AMARELO = "FFF2CC"
VERDE = "C6EFCE"
AMBAR = "FFEB9C"
VERMELHO = "FFC7CE"
VERMELHO_ESC = "C00000"

fino = Side(style="thin", color="D9D9D9")
BORDA = Border(left=fino, right=fino, top=fino, bottom=fino)
CENTRO = Alignment(horizontal="center", vertical="center")
ESQ = Alignment(horizontal="left", vertical="center", indent=1)
DIR = Alignment(horizontal="right", vertical="center")


# ---- Utilitários de formatação -----------------------------------------------
def cabecalho(ws, colunas, cor=NAVY, altura=30, filtro=True):
    for j, t in enumerate(colunas, start=1):
        c = ws.cell(row=1, column=j, value=t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=cor)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA
    ws.row_dimensions[1].height = altura
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if filtro:
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(colunas)).column_letter}{LIM}"


def larguras(ws, valores):
    for letra, largura in valores.items():
        ws.column_dimensions[letra].width = largura


def preencher(ws, faixa, cor):
    for linha in ws[faixa]:
        for c in linha:
            c.fill = PatternFill("solid", fgColor=cor)


def bordas(ws, faixa):
    for linha in ws[faixa]:
        for c in linha:
            c.border = BORDA


def zebra(ws, faixa, col_chave="A"):
    primeira = faixa.split(":")[0]
    linha = "".join(ch for ch in primeira if ch.isdigit())
    ws.conditional_formatting.add(
        faixa,
        FormulaRule(formula=[f'AND(${col_chave}{linha}<>"",MOD(ROW(),2)=0)'],
                    fill=PatternFill("solid", bgColor=CINZA_CLARO)),
    )


def cores_situacao(ws, faixa):
    regras = (('"Além da meta"', VERMELHO, "9C0006"), ('"Atenção"', AMBAR, "9C5700"), ('"OK"', VERDE, "006100"))
    for valor, fundo, letra in regras:
        ws.conditional_formatting.add(faixa, CellIsRule(operator="equal", formula=[valor],
                                      fill=PatternFill("solid", bgColor=fundo), font=Font(bold=True, color=letra)))


def barra_percentual(ws, faixa):
    ws.conditional_formatting.add(
        faixa, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5", showValue=True))


def titulo_secao(ws, faixa, texto, cor=AZUL):
    ws.merge_cells(faixa)
    c = ws[faixa.split(":")[0]]
    c.value = texto
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=cor)
    c.alignment = ESQ
    preencher(ws, faixa, cor)


def cab_linha(ws, linha, col_ini, textos, cor=NAVY):
    for j, t in enumerate(textos):
        c = ws.cell(row=linha, column=col_ini + j, value=t)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=cor)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA


def data_txt(ref):
    # sem TEXT(): os códigos de formato mudam com o idioma do Excel (aaaa x yyyy)
    return f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'


def validacao_data(ws, faixa):
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                        error="Digite uma data válida (dd/mm/aaaa).", errorTitle="Data inválida")
    dv.add(faixa)
    ws.add_data_validation(dv)


wb = Workbook()

# =====================================================================
# Equipamentos (uma linha por tag) - cálculo de tudo
# =====================================================================
eq = wb.active
eq.title = "Equipamentos"
cabecalho(eq, ["Tag / Equipamento", "Tipo", "Unidade", "MCC", "Meta", "Desde", "Acumulado inicial",
               "Acumulado", "Falta", "% da vida", "Situação", "Veio", "chave"], filtro=False)
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
N = len(equip)  # 21
for i in range(2, EQ_ULT + 1):
    k = i - 2
    if k < N:
        nome, tipo, veio = equip[k]
        molde = tipo == "Molde"
        eq.cell(i, 1, nome)
        eq.cell(i, 2, tipo)
        eq.cell(i, 4, None if molde else 4)
        eq.cell(i, 5, META_PADRAO[tipo])
        eq.cell(i, 6, None if molde else INICIO_DADOS)
        eq.cell(i, 12, veio)
    ult_hist = f"SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=L{i})*Historico!$A$2:$A${LIM}))"
    eq[f"C{i}"] = f'=IF(B{i}="","",IF(B{i}="Molde","corridas",IF(B{i}="Oscilador","dias","t")))'
    eq[f"H{i}"] = (
        f'=IF(A{i}="","",IF(C{i}="corridas",'
        f"SUMIFS(Historico!$E:$E,Historico!$B:$B,L{i},Historico!$A:$A,{ult_hist}),"
        f'N(G{i})+IF(F{i}="",0,IF(C{i}="dias",MAX(0,TODAY()-F{i}),'
        f'IF(D{i}=2,SUMIFS(Producao!$D:$D,Producao!$A:$A,">="&F{i}),'
        f'IF(D{i}=3,SUMIFS(Producao!$E:$E,Producao!$A:$A,">="&F{i}),'
        f'SUMIFS(Producao!$B:$B,Producao!$A:$A,">="&F{i})))))))'
    )
    eq[f"I{i}"] = f'=IF(OR(A{i}="",E{i}=""),"",E{i}-H{i})'
    eq[f"J{i}"] = f'=IF(OR(A{i}="",N(E{i})<=0),"",H{i}/E{i})'
    eq[f"K{i}"] = f'=IF(I{i}="","",IF(I{i}<0,"Além da meta",IF(J{i}>=Regras!$B$3,"Atenção","OK")))'
    eq[f"M{i}"] = f'=IF(J{i}="",-1,J{i}+ROW()/1000000000)'  # chave única para o ranking do Painel
    for col, fmt in (("E", "#,##0"), ("G", "#,##0"), ("H", "#,##0"), ("I", "#,##0"),
                     ("F", "dd/mm/yyyy"), ("J", "0.0%")):
        eq[f"{col}{i}"].number_format = fmt
    for c in eq[i][:12]:
        c.border = BORDA
        c.alignment = ESQ if c.column == 1 else CENTRO if c.column in (2, 3, 4, 6, 11, 12) else DIR
preencher(eq, f"A2:B{EQ_ULT}", AMARELO)
preencher(eq, f"D2:G{EQ_ULT}", AMARELO)
barra_percentual(eq, f"J2:J{EQ_ULT}")
cores_situacao(eq, f"K2:K{EQ_ULT}")
eq.column_dimensions["M"].hidden = True
notas = [
    ("Como preencher (células amarelas):", True, CINZA),
    ("•  Tag / Equipamento: a tag real de cada peça.   •  Tipo: escolha na lista.   •  MCC: de onde vem a produção (2, 3 ou 4).", False, CINZA),
    ("•  Meta: vida útil daquela tag.   •  Desde: data de instalação.   •  Acumulado inicial: o que já tinha até o dia anterior a 'Desde'.", False, CINZA),
    ("Molde: automático pelo relatório (corridas do veio G ou H). Oscilador conta dias; componentes somam as toneladas da MCC desde 'Desde'.", False, CINZA),
    ("ATENÇÃO: metas padrão por tipo e 'Desde = 01/08/2026' são PROVISÓRIOS. Substitua pelos valores reais de cada tag.", True, VERMELHO_ESC),
    ("Trocou um equipamento? Registre a saída na aba TrocasEquip, depois digite a nova data em 'Desde' e zere o 'Acumulado inicial'.", False, CINZA),
]
for n_, (txt, negrito, cor) in enumerate(notas):
    c = eq.cell(EQ_ULT + 3 + n_, 1, txt)
    c.font = Font(italic=not negrito, bold=negrito, color=cor)
larguras(eq, {"A": 22, "B": 17, "C": 10, "D": 7, "E": 13, "F": 12, "G": 16, "H": 14, "I": 14, "J": 16, "K": 14, "L": 7})
dv = DataValidation(type="list", formula1='"' + ",".join(TIPOS) + '"', allow_blank=True)
dv.add(f"B2:B{EQ_ULT}")
eq.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"2,3,4"', allow_blank=True)
dv.add(f"D2:D{EQ_ULT}")
eq.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"C,D,E,F,G,H"', allow_blank=True)
dv.add(f"L2:L{EQ_ULT}")
eq.add_data_validation(dv)
validacao_data(eq, f"F2:F{EQ_ULT}")
eq.sheet_properties.tabColor = AZUL

EQ = lambda col: f"Equipamentos!${col}$2:${col}${EQ_ULT}"  # noqa: E731

# =====================================================================
# Moldes (todos os veios, pelo relatório)
# =====================================================================
ml = wb.create_sheet("Moldes")
cabecalho(ml, ["Veio", "MCC", "Molde atual", "Meta", "Acumulado", "Falta", "% da vida", "Situação",
               "Trocas (qtd)", "Corridas médias na saída", "Última troca", "Dias parado", "Dados de"],
          altura=40, filtro=False)
veios = [("C", 2), ("D", 2), ("E", 3), ("F", 3), ("G", 4), ("H", 4)]
for i, (veio, mcc) in enumerate(veios, start=2):
    ml.cell(i, 1, veio)
    ml.cell(i, 2, f"MCC{mcc}")
    ml[f"M{i}"] = f"=SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=A{i})*Historico!$A$2:$A${LIM}))"
    base = f"Historico!$B:$B,A{i},Historico!$A:$A,M{i}"
    ml[f"C{i}"] = f'=IF(M{i}=0,"",SUMIFS(Historico!$C:$C,{base}))'
    ml[f"D{i}"] = f'=IF(M{i}=0,"",SUMIFS(Historico!$D:$D,{base}))'
    ml[f"E{i}"] = f'=IF(M{i}=0,"",SUMIFS(Historico!$E:$E,{base}))'
    ml[f"F{i}"] = f'=IF(M{i}=0,"",D{i}-E{i})'
    ml[f"G{i}"] = f'=IF(M{i}=0,"",IF(N(D{i})>0,E{i}/D{i},""))'
    ml[f"H{i}"] = f'=IF(M{i}=0,"",IF(F{i}<0,"Além da meta",IF(G{i}>=Regras!$B$3,"Atenção","OK")))'
    ml[f"I{i}"] = f"=COUNTIFS(Trocas!$B:$B,A{i})"
    ml[f"J{i}"] = f'=IFERROR(AVERAGEIFS(Trocas!$E:$E,Trocas!$B:$B,A{i}),"-")'
    ml[f"K{i}"] = f'=IF(I{i}=0,"-",SUMPRODUCT(MAX((Trocas!$B$2:$B${LIM}=A{i})*Trocas!$A$2:$A${LIM})))'
    ml[f"L{i}"] = f"=SUMIFS(Paradas!$E:$E,Paradas!$C:$C,A{i})"
    for col, fmt in (("D", "#,##0"), ("E", "#,##0"), ("F", "#,##0"), ("G", "0.0%"), ("J", "#,##0"),
                     ("K", "dd/mm/yyyy"), ("M", 'dd/mm/yyyy;;"-"')):
        ml[f"{col}{i}"].number_format = fmt
    for c in ml[i]:
        c.border = BORDA
        c.alignment = CENTRO
    ml.cell(i, 1).font = Font(bold=True, size=12, color=NAVY)
barra_percentual(ml, "G2:G7")
cores_situacao(ml, "H2:H7")
ml["A9"] = "Todos os veios (C a H) com o último relatório importado. Corridas médias na saída = média das trocas registradas."
ml["A9"].font = Font(italic=True, color=CINZA)
larguras(ml, {"A": 7, "B": 8, "C": 12, "D": 10, "E": 12, "F": 10, "G": 16, "H": 14,
              "I": 12, "J": 18, "K": 14, "L": 12, "M": 13})
ml.sheet_properties.tabColor = AZUL

# =====================================================================
# Abas manuais
# =====================================================================
te = wb.create_sheet("TrocasEquip")
cabecalho(te, ["Data da saída", "Equipamento", "Acumulado na saída", "Unidade", "Motivo", "Observação"], cor=OURO)
for r in range(2, ULTIMA + 1):
    te[f"A{r}"].number_format = "dd/mm/yyyy"
    te[f"C{r}"].number_format = "#,##0"
preencher(te, f"A2:F{ULTIMA}", AMARELO)
larguras(te, {"A": 15, "B": 22, "C": 20, "D": 10, "E": 45, "F": 40})
dv = DataValidation(type="list", formula1=f"={EQ('A')}", allow_blank=True)
dv.add(f"B2:B{ULTIMA}")
te.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"t,corridas,dias"', allow_blank=True)
dv.add(f"D2:D{ULTIMA}")
te.add_data_validation(dv)
validacao_data(te, f"A2:A{ULTIMA}")
te.sheet_properties.tabColor = OURO

pa = wb.create_sheet("Paradas")
cabecalho(pa, ["Início", "Fim", "Veio / equipamento", "Motivo", "Dias parado", "Observação"], cor=OURO)
for r in range(2, ULTIMA + 1):
    pa[f"E{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",B{r}-A{r})'
    pa[f"A{r}"].number_format = "dd/mm/yyyy"
    pa[f"B{r}"].number_format = "dd/mm/yyyy"
preencher(pa, f"A2:D{ULTIMA}", AMARELO)
preencher(pa, f"F2:F{ULTIMA}", AMARELO)
larguras(pa, {"A": 13, "B": 13, "C": 20, "D": 45, "E": 13, "F": 40})
dv = DataValidation(type="list", formula1='"C,D,E,F,G,H"', allow_blank=True, showErrorMessage=False)
dv.add(f"C2:C{ULTIMA}")
pa.add_data_validation(dv)
validacao_data(pa, f"A2:B{ULTIMA}")
pa.sheet_properties.tabColor = OURO

# =====================================================================
# Abas do robô (NÃO mudar nome, linha 1 nem ordem das colunas)
# =====================================================================
hs = wb.create_sheet("Historico")
cabecalho(hs, ["Data", "Veio", "Molde nº", "Meta (vida progr.)", "Acumulado (corridas)", "Observação"], cor=CINZA)
larguras(hs, {"A": 13, "B": 7, "C": 11, "D": 18, "E": 20, "F": 75})
zebra(hs, f"A2:F{LIM}")
hs.sheet_properties.tabColor = CINZA

pr = wb.create_sheet("Producao")
cabecalho(pr, ["Data", "MCC4 real (t)", "MCC4 prog (t)", "MCC2 real (t)", "MCC3 real (t)", "Arquivo"], cor=CINZA)
for r in range(2, 400):
    for col in "BCDE":
        pr[f"{col}{r}"].number_format = "#,##0"
larguras(pr, {"A": 13, "B": 15, "C": 15, "D": 15, "E": 15, "F": 36})
zebra(pr, f"A2:F{LIM}")
pr.sheet_properties.tabColor = CINZA

tr = wb.create_sheet("Trocas")
cabecalho(tr, ["Data", "Veio", "Molde que saiu", "Vida programada", "Corridas na saída",
               "Molde que entrou", "Motivo", "Conferir", "Origem"], cor=CINZA)
larguras(tr, {"A": 13, "B": 7, "C": 15, "D": 16, "E": 17, "F": 16, "G": 60, "H": 45, "I": 10})
tr.conditional_formatting.add(f"H2:H{LIM}", FormulaRule(formula=["LEN($H2)>0"], fill=PatternFill("solid", bgColor=AMBAR)))
zebra(tr, f"A2:G{LIM}")
tr.sheet_properties.tabColor = CINZA

al = wb.create_sheet("Alertas")
cabecalho(al, ["Data/hora", "Tipo", "Item", "Detalhe", "Resolvido"], cor=VERMELHO_ESC)
larguras(al, {"A": 18, "B": 28, "C": 16, "D": 95, "E": 11})
al.conditional_formatting.add(
    f"A2:E{LIM}", FormulaRule(formula=['AND($A2<>"",$E2<>"Sim")'], fill=PatternFill("solid", bgColor=VERMELHO)))
dv = DataValidation(type="list", formula1='"Sim"', allow_blank=True)
dv.add(f"E2:E{LIM}")
al.add_data_validation(dv)
al.sheet_properties.tabColor = VERMELHO_ESC

# =====================================================================
# Regras
# =====================================================================
rg = wb.create_sheet("Regras")
rg.sheet_view.showGridLines = False
cab_linha(rg, 1, 1, ["Parâmetro", "Valor"])
rg["A3"], rg["B3"] = "Limiar de atenção (% da vida)", 0.8
rg["B3"].number_format = "0%"
rg["B3"].fill = PatternFill("solid", fgColor=AMARELO)
rg["B3"].border = BORDA
rg["A5"] = "Como a planilha calcula"
rg["A5"].font = Font(bold=True, size=12, color=NAVY)
textos = [
    "Molde (corridas): acumulado do último dia em Historico, pelo veio. Um registro por data e veio; reenvio do mesmo dia substitui, não soma.",
    "Troca de molde: quando o nº do molde do veio muda, o robô registra em Trocas (molde que saiu, corridas na saída pela observação do relatório, molde que entrou).",
    "Oscilador (dias): dias corridos desde 'Desde' + Acumulado inicial.",
    "Componentes (t): toneladas reais da MCC escolhida (aba Producao) desde 'Desde' + Acumulado inicial.",
    "Falta = meta − acumulado.   % = acumulado ÷ meta.   Acima de 100% = além da meta.   Sem previsão de data de troca.",
    "Alertas: acumulado diminuiu sem troca; veio não andou enquanto a MCC produziu; layout dos anexos mudou; troca para conferir.",
    "Cores das abas: azul = painel e cadastros; dourado = você digita; cinza = o robô preenche; vermelho = alertas.",
    "As abas Historico, Producao, Trocas e Alertas são escritas pelo robô: não renomeie nem mude a ordem das colunas.",
]
for i, t in enumerate(textos, start=6):
    rg[f"A{i}"] = "•  " + t
rg["A16"] = "Autor: " + MARCA
rg["A16"].font = Font(italic=True, color=CINZA)
larguras(rg, {"A": 150, "B": 12})
rg.sheet_properties.tabColor = CINZA

# =====================================================================
# Auxiliar (oculta): dados dos gráficos e do ranking
# =====================================================================
ax = wb.create_sheet("Auxiliar")
ax.append(["Dia", "Veio G (corridas)", "Veio H (corridas)", "MCC4 real (t)", None, "Top (linha)"])
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
TOP = 6
for n in range(1, TOP + 1):
    ax[f"F{n + 1}"] = (
        f"=IFERROR(IF(LARGE({EQ('M')},{n})<0,0,MATCH(LARGE({EQ('M')},{n}),{EQ('M')},0)),0)"
    )
ax.sheet_state = "hidden"

# =====================================================================
# Painel
# =====================================================================
pn = wb.create_sheet("Painel", 0)
pn.sheet_view.showGridLines = False
larguras(pn, {"A": 22, "B": 16, "C": 11, "D": 13, "E": 13, "F": 13, "G": 15, "H": 14, "I": 13})

pn.merge_cells("A1:I1")
pn["A1"] = "Controle de Vida dos Equipamentos  |  MCC4"
pn["A1"].font = Font(size=18, bold=True, color="FFFFFF")
pn["A1"].fill = PatternFill("solid", fgColor=NAVY)
pn["A1"].alignment = ESQ
pn.row_dimensions[1].height = 36
pn.merge_cells("A2:I2")
_ult = f"MAX(Historico!$A$2:$A${LIM})"
pn["A2"] = (f'="Último relatório importado: "&IF({_ult}=0,"nenhum ainda",{data_txt(_ult)})'
            f'&"     |     Hoje: "&{data_txt("TODAY()")}')
pn["A2"].font = Font(italic=True, color=CINZA)
pn["A2"].alignment = ESQ
pn.merge_cells("A3:I3")
pn["A3"] = MARCA
pn["A3"].font = Font(size=8, italic=True, color="808080")
pn["A3"].alignment = ESQ

# Indicadores
EQK = EQ("K")
kpis = [
    ("A", "A", "Equipamentos", f"=COUNTA({EQ('A')})", AZUL_CLARO, NAVY),
    ("B", "C", "Além da meta", f'=COUNTIF({EQK},"Além da meta")', VERMELHO, "9C0006"),
    ("D", "E", "Atenção", f'=COUNTIF({EQK},"Atenção")', AMBAR, "9C5700"),
    ("F", "G", "OK", f'=COUNTIF({EQK},"OK")', VERDE, "006100"),
    ("H", "I", "Alertas abertos", f'=COUNTA(Alertas!$A$2:$A${LIM})-COUNTIFS(Alertas!$E$2:$E${LIM},"Sim")', CINZA_CLARO, VERMELHO_ESC),
]
for c_ini, c_fim, titulo, formula, fundo, cor in kpis:
    if c_ini != c_fim:
        pn.merge_cells(f"{c_ini}5:{c_fim}5")
        pn.merge_cells(f"{c_ini}6:{c_fim}6")
    t = pn[f"{c_ini}5"]
    v = pn[f"{c_ini}6"]
    t.value, v.value = titulo, formula
    t.font = Font(bold=True, color=cor)
    t.alignment = CENTRO
    v.font = Font(size=26, bold=True, color=cor)
    v.alignment = CENTRO
    preencher(pn, f"{c_ini}5:{c_fim}6", fundo)
pn.row_dimensions[6].height = 44

# Mais desgastados (A..D) e Por tipo (F..I)
titulo_secao(pn, "A8:D8", "Mais desgastados")
titulo_secao(pn, "F8:I8", "Por tipo")
cab_linha(pn, 9, 1, ["Tag / Equipamento", "Tipo", "% da vida", "Situação"])
cab_linha(pn, 9, 6, ["Tipo", "Qtd", "Maior % da vida", "Atenção / além"])
for n in range(1, TOP + 1):
    r = 9 + n
    idx = f"Auxiliar!$F${n + 1}"
    for j, col in ((1, "A"), (2, "B"), (3, "J"), (4, "K")):
        c = pn.cell(r, j, f'=IF({idx}=0,"",INDEX({EQ(col)},{idx}))')
        c.border = BORDA
        c.alignment = ESQ if j == 1 else CENTRO
    pn.cell(r, 3).number_format = "0.0%"
barra_percentual(pn, f"C10:C{9 + TOP}")
cores_situacao(pn, f"D10:D{9 + TOP}")
for n, tipo in enumerate(TIPOS):
    r = 10 + n
    pn.cell(r, 6, tipo)
    pn.cell(r, 7, f'=COUNTIF({EQ("B")},F{r})')
    pn.cell(r, 8, f'=IF(G{r}=0,"",SUMPRODUCT(MAX(({EQ("B")}=F{r})*{EQ("M")})))')
    pn.cell(r, 9, f'=COUNTIFS({EQ("B")},F{r},{EQK},"Atenção")+COUNTIFS({EQ("B")},F{r},{EQK},"Além da meta")')
    pn.cell(r, 8).number_format = "0.0%"
    for j in range(6, 10):
        pn.cell(r, j).border = BORDA
        pn.cell(r, j).alignment = ESQ if j == 6 else CENTRO
barra_percentual(pn, f"H10:H{9 + len(TIPOS)}")
pn.conditional_formatting.add(f"I10:I{9 + len(TIPOS)}", CellIsRule(operator="greaterThan", formula=["0"],
                              fill=PatternFill("solid", bgColor=AMBAR), font=Font(bold=True, color="9C5700")))

# Todos os equipamentos
T0 = 19  # linha do título da seção
PN = N + 3
titulo_secao(pn, f"A{T0}:I{T0}", "Todos os equipamentos")
cab_linha(pn, T0 + 1, 1, ["Tag / Equipamento", "Tipo", "Unidade", "Meta", "Acumulado", "Falta", "% da vida", "Situação", "Desde"])
origem = {1: "A", 2: "B", 3: "C", 4: "E", 5: "H", 6: "I", 7: "J", 8: "K", 9: "F"}
for k in range(PN):
    r = T0 + 2 + k
    for j, col in origem.items():
        ref = f"Equipamentos!{col}{k + 2}"
        vazio = f'Equipamentos!A{k + 2}=""'
        f = f'=IF(OR({vazio},{ref}=""),"",{ref})' if j == 9 else f'=IF({vazio},"",{ref})'
        c = pn.cell(r, j, f)
        c.border = BORDA
        c.alignment = ESQ if j == 1 else CENTRO if j in (2, 3, 8, 9) else DIR
        if j in (4, 5, 6):
            c.number_format = "#,##0"
        elif j == 7:
            c.number_format = "0.0%"
        elif j == 9:
            c.number_format = "dd/mm/yyyy"
L1, L2 = T0 + 2, T0 + 1 + PN
barra_percentual(pn, f"G{L1}:G{L2}")
cores_situacao(pn, f"H{L1}:H{L2}")
zebra(pn, f"A{L1}:F{L2}")

# Gráficos
G0 = L2 + 2
titulo_secao(pn, f"A{G0}:I{G0}", "Gráficos")
c1 = BarChart()
c1.type = "bar"
c1.title = "% da vida consumida"
c1.style = 10
c1.add_data(Reference(pn, min_col=7, min_row=T0 + 1, max_row=L2), titles_from_data=True)
c1.set_categories(Reference(pn, min_col=1, min_row=L1, max_row=L2))
c1.y_axis.number_format = "0%"
c1.y_axis.scaling.min = 0
c1.legend = None
c1.height, c1.width = 13, 12.5
pn.add_chart(c1, f"A{G0 + 1}")

c2 = LineChart()
c2.title = "Corridas dos moldes G e H (30 dias)"
c2.style = 12
c2.add_data(Reference(ax, min_col=2, max_col=3, min_row=1, max_row=31), titles_from_data=True)
c2.set_categories(Reference(ax, min_col=1, min_row=2, max_row=31))
c2.y_axis.scaling.min = 0
c2.height, c2.width = 6.3, 12.5
pn.add_chart(c2, f"E{G0 + 1}")

c3 = BarChart()
c3.type = "col"
c3.title = "Toneladas reais do MCC4 por dia (30 dias)"
c3.style = 10
c3.add_data(Reference(ax, min_col=4, min_row=1, max_row=31), titles_from_data=True)
c3.set_categories(Reference(ax, min_col=1, min_row=2, max_row=31))
c3.legend = None
c3.height, c3.width = 6.3, 12.5
pn.add_chart(c3, f"E{G0 + 14}")

pn.sheet_properties.tabColor = NAVY
pn.page_setup.orientation = "portrait"
pn.page_setup.fitToWidth = 1
pn.page_setup.fitToHeight = 0
pn.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# =====================================================================
# Início (como usar)
# =====================================================================
ini = wb.create_sheet("Início", 0)
ini.sheet_view.showGridLines = False
larguras(ini, {"A": 3, "B": 26, "C": 95})
ini.merge_cells("B2:C2")
ini["B2"] = "Controle de Vida dos Equipamentos"
ini["B2"].font = Font(size=20, bold=True, color="FFFFFF")
ini["B2"].fill = PatternFill("solid", fgColor=NAVY)
ini["B2"].alignment = ESQ
preencher(ini, "B2:C2", NAVY)
ini.row_dimensions[2].height = 40
ini.merge_cells("B3:C3")
ini["B3"] = MARCA
ini["B3"].font = Font(size=9, italic=True, color="808080")

linha = 5
titulo_secao(ini, f"B{linha}:C{linha}", "Abas (clique para ir)")
abas = [
    ("Painel", "Visão geral: indicadores, mais desgastados, resumo por tipo, todos os equipamentos e gráficos.", NAVY),
    ("Equipamentos", "Cadastro de cada tag (meta, MCC, Desde, acumulado inicial) e o cálculo da vida.", AZUL),
    ("Moldes", "Todos os veios (C a H) pelo relatório: molde atual, corridas, trocas, dias parados.", AZUL),
    ("TrocasEquip", "Você registra a saída de oscilador e componentes.", OURO),
    ("Paradas", "Você registra paradas (início, fim, motivo).", OURO),
    ("Historico", "Robô: corridas de cada veio por dia.", CINZA),
    ("Producao", "Robô: toneladas reais por MCC por dia.", CINZA),
    ("Trocas", "Robô: trocas de molde detectadas no relatório.", CINZA),
    ("Alertas", "Robô: avisos para conferir. Escolha 'Sim' em Resolvido quando tratar.", VERMELHO_ESC),
    ("Regras", "Limiar de atenção e explicação dos cálculos.", CINZA),
]
for k, (nome, desc, cor) in enumerate(abas, start=linha + 1):
    c = ini.cell(k, 2, nome)
    c.hyperlink = f"#'{nome}'!A1"
    c.font = Font(bold=True, color=cor, underline="single")
    c.fill = PatternFill("solid", fgColor=CINZA_CLARO)
    c.alignment = ESQ
    c.border = Border(left=Side(style="thick", color=cor), right=fino, top=fino, bottom=fino)
    d = ini.cell(k, 3, desc)
    d.border = BORDA
    d.alignment = ESQ
linha = linha + len(abas) + 2

passos = [
    ("Primeira vez", [
        "Na aba Equipamentos, troque os nomes pelas tags reais e ajuste meta, MCC, Desde e Acumulado inicial (células amarelas).",
        "Coloque esta planilha e o robô atualizar_controle_GRAVANDO.vbs na mesma pasta, fora do OneDrive.",
        "Feche a planilha, abra o Outlook clássico e dê dois cliques no robô. Espere o aviso 'Robô terminou'.",
    ]),
    ("Todo dia", [
        "Rode o robô (ou deixe agendado no Agendador de Tarefas do Windows) com a planilha fechada.",
        "Abra o Painel: veja os quadros do topo e os mais desgastados.",
        "Se 'Alertas abertos' for maior que 0, confira na aba Alertas e marque 'Sim' em Resolvido.",
    ]),
    ("Quando trocar", [
        "Molde: nada a fazer, o robô detecta a troca pelo relatório.",
        "Oscilador ou componente: registre a saída em TrocasEquip, depois na aba Equipamentos digite a nova data em Desde e zere o Acumulado inicial.",
    ]),
    ("Cores", [
        "Amarelo = você digita.   Verde = OK.   Âmbar = atenção (a partir do limiar em Regras).   Vermelho = além da meta ou alerta aberto.",
    ]),
]
for titulo, itens_ in passos:
    titulo_secao(ini, f"B{linha}:C{linha}", titulo)
    for n_, txt in enumerate(itens_, start=1):
        linha += 1
        ini.cell(linha, 2, f"{n_}.").alignment = Alignment(horizontal="right", vertical="top")
        c = ini.cell(linha, 3, txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    linha += 2
ini.sheet_properties.tabColor = NAVY
ini.page_setup.fitToWidth = 1
ini.page_setup.fitToHeight = 1
ini.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

wb.active = 1  # abre no Painel
wb.save(SAIDA)
print(f"Gerado: {SAIDA}")
