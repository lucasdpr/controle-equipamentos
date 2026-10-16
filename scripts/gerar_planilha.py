"""Gera controle_equipamentos.xlsx (planilha própria, separada do arquivo do e-mail).

Fontes diárias (anexos do e-mail "Relatórios GLA", lidos pelo robô):
- VIDA MOLDE.xlsx  -> acumulado de corridas por veio (C..H) e nº do molde.
- CORTESEQ.xls     -> aba DADOS DA PRODUÇÃO: toneladas reais por MCC por dia.

Regras:
- Molde: corridas (acumulado do último dia importado, por veio). Meta fixa por veio.
- Oscilador: dias em máquina (entrada conta, saída não; em aberto conta até hoje).
- Componentes: toneladas reais do MCC4 nos dias entre entrada e saída.
- Falta = meta - acumulado. % = acumulado / meta. Sem previsão.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

SAIDA = Path(__file__).resolve().parent.parent / "controle_equipamentos.xlsx"
ULTIMA = 1000   # linhas com fórmula pré-preenchida (Movimentos)
LIM = 5000      # intervalo usado nas fórmulas de Historico/Alertas

CABECALHO = Font(bold=True, color="FFFFFF")
FUNDO = PatternFill("solid", fgColor="1F4E78")
FUNDO_ENTRADA = PatternFill("solid", fgColor="FFF2CC")
FUNDO_ALERTA = PatternFill("solid", fgColor="F8CBAD")


def cabecalho(ws, colunas):
    ws.append(colunas)
    for celula in ws[1]:
        celula.font = CABECALHO
        celula.fill = FUNDO
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def larguras(ws, valores):
    for letra, largura in valores.items():
        ws.column_dimensions[letra].width = largura


wb = Workbook()

# ---------- Resumo ----------
rs = wb.active
rs.title = "Resumo"
cabecalho(
    rs,
    ["Item", "Tipo", "Unidade", "Meta", "Acumulado inicial", "Acumulado", "Falta",
     "%", "Situação", "Em máquina agora", "Veio"],
)
itens = [
    ("Molde G", "Molde", "corridas", 1100, "G"),
    ("Molde H", "Molde", "corridas", 1100, "H"),
    ("Oscilador", "Oscilador", "dias", 730, None),
    ("Bender", "Componente", "t", 1100000, None),
    ("Bow", "Componente", "t", 1900000, None),
    ("Straightener R1", "Componente", "t", 1700000, None),
    ("Straightener R2", "Componente", "t", 1700000, None),
    ("Horizontal", "Componente", "t", 3300000, None),
]
for i, (nome, tipo, unidade, meta, veio) in enumerate(itens, start=2):
    rs.append([nome, tipo, unidade, meta, None, None, None, None, None, None, veio])
    ultima_data = f"SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=K{i})*Historico!$A$2:$A${LIM}))"
    rs[f"F{i}"] = (
        f'=IF(C{i}="corridas",'
        f"SUMIFS(Historico!$E:$E,Historico!$B:$B,K{i},Historico!$A:$A,{ultima_data}),"
        f'N(E{i})+IF(C{i}="dias",'
        f"SUMIFS(Movimentos!$D:$D,Movimentos!$A:$A,A{i}),"
        f"SUMIFS(Movimentos!$F:$F,Movimentos!$A:$A,A{i})))"
    )
    rs[f"G{i}"] = f'=IF(D{i}="","",D{i}-F{i})'
    rs[f"H{i}"] = f'=IF(N(D{i})>0,F{i}/D{i},"")'
    rs[f"I{i}"] = (
        f'=IF(G{i}="","",IF(G{i}<0,"Além da meta",'
        f'IF(H{i}>=Regras!$B$3,"Atenção","OK")))'
    )
    rs[f"J{i}"] = (
        f'=IF(C{i}="corridas","-",'
        f'IF(COUNTIFS(Movimentos!$A:$A,A{i},Movimentos!$C:$C,"")>0,"Sim","Não"))'
    )
for linha in rs.iter_rows(min_row=2, max_row=len(itens) + 1):
    for idx in (3, 4, 5, 6):
        linha[idx].number_format = "#,##0.##"
    linha[7].number_format = "0.0%"
rs["M1"] = (
    f'="Alertas abertos: "&(COUNTA(Alertas!$A$2:$A${LIM})-COUNTIFS(Alertas!$E$2:$E${LIM},"Sim"))'
)
rs["M1"].font = Font(bold=True)
rs.conditional_formatting.add(
    "M1",
    FormulaRule(
        formula=[f'COUNTA(Alertas!$A$2:$A${LIM})-COUNTIFS(Alertas!$E$2:$E${LIM},"Sim")>0'],
        fill=FUNDO_ALERTA,
    ),
)
larguras(rs, {"A": 18, "B": 13, "C": 11, "D": 14, "E": 16, "F": 14, "G": 14,
              "H": 9, "I": 14, "J": 16, "K": 7, "M": 22})
dv_un = DataValidation(type="list", formula1='"t,corridas,dias"', allow_blank=True)
dv_un.add("C2:C200")
rs.add_data_validation(dv_un)

# ---------- Movimentos (entrada e saída, manual) ----------
mv = wb.create_sheet("Movimentos")
cabecalho(mv, ["Equipamento", "Entrada", "Saída", "Dias rodados", "Alerta", "Toneladas do período"])
for r in range(2, ULTIMA + 1):
    mv[f"D{r}"] = f'=IF(B{r}="","",IF(C{r}="",TODAY(),C{r})-B{r})'
    mv[f"E{r}"] = (
        f'=IF(A{r}="","",IF(COUNTIFS($A:$A,A{r},$C:$C,"")>1,"Mais de uma entrada em aberto",'
        f'IF(AND(C{r}<>"",C{r}<B{r}),"Saída antes da entrada","")))'
    )
    mv[f"F{r}"] = (
        f'=IF(OR(A{r}="",B{r}=""),"",SUMIFS(Producao!$B:$B,Producao!$A:$A,">="&B{r},'
        f'Producao!$A:$A,"<"&IF(C{r}="",TODAY()+1,C{r})))'
    )
    mv[f"B{r}"].number_format = "dd/mm/yyyy"
    mv[f"C{r}"].number_format = "dd/mm/yyyy"
    mv[f"F{r}"].number_format = "#,##0"
larguras(mv, {"A": 20, "B": 13, "C": 13, "D": 14, "E": 30, "F": 20})
dv_eq = DataValidation(type="list", formula1="=Resumo!$A$2:$A$200", allow_blank=True)
dv_eq.add(f"A2:A{ULTIMA}")
mv.add_data_validation(dv_eq)
mv.conditional_formatting.add(f"E2:E{ULTIMA}", FormulaRule(formula=["LEN($E2)>0"], fill=FUNDO_ALERTA))

# ---------- Producao (toneladas reais por MCC, preenchida pelo robô) ----------
pr = wb.create_sheet("Producao")
cabecalho(pr, ["Data", "MCC4 real (t)", "MCC4 prog (t)", "MCC2 real (t)", "MCC3 real (t)", "Arquivo"])
for r in range(2, 3):
    pr[f"A{r}"].number_format = "dd/mm/yyyy"
larguras(pr, {"A": 13, "B": 15, "C": 15, "D": 15, "E": 15, "F": 36})

# ---------- Historico (um registro por data e veio, preenchido pelo robô) ----------
hs = wb.create_sheet("Historico")
cabecalho(hs, ["Data", "Veio", "Molde nº", "Meta (vida progr.)", "Acumulado (corridas)", "Observação"])
larguras(hs, {"A": 13, "B": 7, "C": 11, "D": 18, "E": 20, "F": 70})

# ---------- Trocas ----------
tr = wb.create_sheet("Trocas")
cabecalho(
    tr,
    ["Data", "Veio", "Molde que saiu", "Vida programada", "Corridas na saída",
     "Molde que entrou", "Motivo", "Conferir", "Origem"],
)
larguras(tr, {"A": 13, "B": 7, "C": 15, "D": 16, "E": 17, "F": 16, "G": 60, "H": 45, "I": 10})

# ---------- Paradas (manual) ----------
pa = wb.create_sheet("Paradas")
cabecalho(pa, ["Início", "Fim", "Veio / equipamento", "Motivo", "Dias parado", "Observação"])
for r in range(2, ULTIMA + 1):
    pa[f"E{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",B{r}-A{r})'
    pa[f"A{r}"].number_format = "dd/mm/yyyy"
    pa[f"B{r}"].number_format = "dd/mm/yyyy"
larguras(pa, {"A": 13, "B": 13, "C": 20, "D": 45, "E": 13, "F": 40})

# ---------- Alertas (robô escreve; você marca "Sim" quando resolver) ----------
al = wb.create_sheet("Alertas")
cabecalho(al, ["Data/hora", "Tipo", "Item", "Detalhe", "Resolvido"])
larguras(al, {"A": 18, "B": 28, "C": 16, "D": 90, "E": 11})

# ---------- Regras ----------
rg = wb.create_sheet("Regras")
rg["A1"], rg["B1"] = "Parâmetro", "Valor"
for celula in (rg["A1"], rg["B1"]):
    celula.font = CABECALHO
    celula.fill = FUNDO
rg["A3"], rg["B3"] = "Limiar de atenção (%)", 0.8
rg["B3"].number_format = "0%"
rg["B3"].fill = FUNDO_ENTRADA
rg["A5"] = "Como a planilha calcula"
rg["A5"].font = Font(bold=True)
textos = [
    "Molde (corridas): acumulado do último dia em Historico, por veio. O robô grava um registro por data e veio; reenvio do mesmo dia substitui, não soma.",
    "Troca de molde: quando o nº do molde do veio muda de um dia para o outro, o robô registra em Trocas (molde que saiu, corridas na saída, molde que entrou, motivo).",
    "Oscilador (dias): saída − entrada em Movimentos. A entrada conta, a saída não. Sem saída, conta até hoje.",
    "Componentes (t): toneladas reais do MCC4 (aba Producao) nos dias entre entrada e saída. Fora desse período não desconta.",
    "Falta = meta − acumulado. % = acumulado ÷ meta. Acima de 100% = além da meta. Sem previsão de data de troca.",
    "Alertas: acumulado diminuiu sem troca; veio não andou enquanto a MCC produziu; layout dos anexos mudou. Marque 'Sim' na coluna Resolvido quando tratar.",
    "Paradas e Movimentos são preenchidos à mão. Historico, Producao, Trocas (robô) e Alertas são preenchidos pelo robô.",
]
for i, t in enumerate(textos, start=6):
    rg[f"A{i}"] = t
larguras(rg, {"A": 140, "B": 12})

wb.save(SAIDA)
print(f"Gerado: {SAIDA}")
