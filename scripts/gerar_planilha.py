"""Gera controle_equipamentos.xlsx com as regras de controle de vida.

Regras (confirmadas pelo usuário):
- Meta: limite de vida do item, até a troca.
- Dias rodados por período = saída - entrada (entrada conta, saída não).
  Período em aberto (sem saída) conta até hoje.
- Acumulado de dias (unidade "dias") = soma dos períodos.
- Acumulado de t/corridas = soma dos lançamentos feitos dentro de um período
  em que o item estava em máquina (entrada <= data < saída).
- Falta = meta - acumulado. % = acumulado / meta.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

SAIDA = Path(__file__).resolve().parent.parent / "controle_equipamentos.xlsx"
ULTIMA = 1000  # linhas com fórmula pré-preenchida nas abas de registro

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

# ---------- Equipamentos (cadastro + resumo) ----------
eq = wb.active
eq.title = "Equipamentos"
cabecalho(
    eq,
    [
        "Equipamento", "Tipo", "Unidade", "Meta", "Acumulado inicial",
        "Acumulado", "Falta", "%", "Situação", "Em máquina agora",
    ],
)

itens = [
    ("Molde MCC4", "Molde", "corridas", 1100, 1080),
    ("Oscilador", "Oscilador", "dias", 730, None),
    ("Bender", "Componente", "t", 1100000, None),
    ("Bow", "Componente", "t", 1900000, None),
    ("Straightener R1", "Componente", "t", 1700000, None),
    ("Straightener R2", "Componente", "t", 1700000, None),
    ("Horizontal", "Componente", "t", 3300000, None),
]
for i, (nome, tipo, unidade, meta, inicial) in enumerate(itens, start=2):
    eq.append([nome, tipo, unidade, meta, inicial])
    eq[f"F{i}"] = (
        f'=N(E{i})+IF(C{i}="dias",'
        f"SUMIFS(Movimentos!$D:$D,Movimentos!$A:$A,A{i}),"
        f'SUMIFS(Lancamentos!$C:$C,Lancamentos!$B:$B,A{i},Lancamentos!$E:$E,"Sim"))'
    )
    eq[f"G{i}"] = f'=IF(D{i}="","",D{i}-F{i})'
    eq[f"H{i}"] = f'=IF(N(D{i})>0,F{i}/D{i},"")'
    eq[f"I{i}"] = (
        f'=IF(G{i}="","",IF(G{i}<0,"Além da meta",'
        f'IF(H{i}>=Regras!$B$3,"Atenção","OK")))'
    )
    eq[f"J{i}"] = (
        f'=IF(COUNTIFS(Movimentos!$A:$A,A{i},Movimentos!$C:$C,"")>0,"Sim","Não")'
    )

for linha in eq.iter_rows(min_row=2, max_row=len(itens) + 1):
    linha[3].number_format = "#,##0"
    linha[4].number_format = "#,##0"
    linha[5].number_format = "#,##0.##"
    linha[6].number_format = "#,##0.##"
    linha[7].number_format = "0.0%"

larguras(eq, {"A": 20, "B": 13, "C": 11, "D": 14, "E": 16, "F": 14, "G": 14, "H": 9, "I": 16, "J": 16})

dv_unidade = DataValidation(type="list", formula1='"t,corridas,dias"', allow_blank=True)
dv_unidade.add(f"C2:C{ULTIMA}")
eq.add_data_validation(dv_unidade)

# ---------- Movimentos (entrada e saída) ----------
mov = wb.create_sheet("Movimentos")
cabecalho(mov, ["Equipamento", "Entrada", "Saída", "Dias rodados", "Alerta"])
for r in range(2, ULTIMA + 1):
    mov[f"D{r}"] = f'=IF(B{r}="","",IF(C{r}="",TODAY(),C{r})-B{r})'
    mov[f"E{r}"] = (
        f'=IF(A{r}="","",IF(COUNTIFS($A:$A,A{r},$C:$C,"")>1,"Mais de uma entrada em aberto",'
        f'IF(AND(C{r}<>"",C{r}<B{r}),"Saída antes da entrada","")))'
    )
    mov[f"B{r}"].number_format = "dd/mm/yyyy"
    mov[f"C{r}"].number_format = "dd/mm/yyyy"
larguras(mov, {"A": 20, "B": 13, "C": 13, "D": 14, "E": 30})

dv_equip_mov = DataValidation(type="list", formula1=f"=Equipamentos!$A$2:$A${ULTIMA}", allow_blank=True)
dv_equip_mov.add(f"A2:A{ULTIMA}")
mov.add_data_validation(dv_equip_mov)
mov.conditional_formatting.add(
    f"E2:E{ULTIMA}",
    FormulaRule(formula=['LEN($E2)>0'], fill=FUNDO_ALERTA),
)

# ---------- Lançamentos (produção: t ou corridas) ----------
lan = wb.create_sheet("Lancamentos")
cabecalho(lan, ["Data", "Equipamento", "Valor", "Origem", "Válido", "Observação"])
for r in range(2, ULTIMA + 1):
    lan[f"E{r}"] = (
        f'=IF(B{r}="","",IF(COUNTIFS(Movimentos!$A:$A,B{r},Movimentos!$B:$B,"<="&A{r},'
        f'Movimentos!$C:$C,">"&A{r})+COUNTIFS(Movimentos!$A:$A,B{r},Movimentos!$B:$B,"<="&A{r},'
        f'Movimentos!$C:$C,"")>0,"Sim","Não"))'
    )
    lan[f"A{r}"].number_format = "dd/mm/yyyy"
larguras(lan, {"A": 13, "B": 20, "C": 14, "D": 14, "E": 10, "F": 40})

dv_equip_lan = DataValidation(type="list", formula1=f"=Equipamentos!$A$2:$A${ULTIMA}", allow_blank=True)
dv_equip_lan.add(f"B2:B{ULTIMA}")
lan.add_data_validation(dv_equip_lan)
lan.conditional_formatting.add(
    f"E2:E{ULTIMA}",
    FormulaRule(formula=['$E2="Não"'], fill=FUNDO_ALERTA),
)

# ---------- Regras ----------
reg = wb.create_sheet("Regras")
reg["A1"] = "Parâmetro"
reg["B1"] = "Valor"
for celula in (reg["A1"], reg["B1"]):
    celula.font = CABECALHO
    celula.fill = FUNDO
reg["A3"] = "Limiar de atenção (%)"
reg["B3"] = 0.8
reg["B3"].number_format = "0%"
reg["B3"].fill = FUNDO_ENTRADA

reg["A5"] = "Como a planilha calcula"
reg["A5"].font = Font(bold=True)
textos = [
    "Meta: limite de vida do item, até a troca.",
    "Dias rodados = saída − entrada. A entrada conta, a saída não. Sem saída, conta até hoje.",
    "Lançamento de t ou corridas só conta se a data cair dentro de um período em que o item estava em máquina (entrada ≤ data < saída).",
    "Lançamento fora de período aparece como 'Não' em Válido, com destaque, e não entra no acumulado.",
    "Falta = meta − acumulado. % = acumulado ÷ meta. Acima de 100% = além da meta.",
    "Sem previsão de dias: a planilha não calcula ritmo nem data de troca.",
    "Acumulado inicial: valor que o item já tinha antes desta planilha começar a ser usada.",
]
for i, t in enumerate(textos, start=6):
    reg[f"A{i}"] = t
reg.column_dimensions["A"].width = 110
reg.column_dimensions["B"].width = 12

wb.save(SAIDA)
print(f"Gerado: {SAIDA}")
