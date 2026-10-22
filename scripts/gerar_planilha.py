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
import importlib.util
import os
import re
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
EQ_ULT = 61     # aba Equipamentos: linhas 2..61 (uma linha por POSIÇÃO, até 60)
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
cabecalho(eq, ["Tag atual", "Tipo", "Unidade", "MCC", "Meta", "Desde", "Acumulado inicial",
               "Acumulado", "Falta", "% da vida", "Situação", "Veio", "chave", "Posição", "ID no app",
               "Veio - Posição", "ativo", "Última troca", "Tag original", "Meta original",
               "Desde original", "Acum. inicial original", "idx"], filtro=False)
TIPOS = ["Molde", "Oscilador", "Bender", "Bow", "Straightener R1", "Straightener R2", "Horizontal"]
META_PADRAO = {"Molde": 1100, "Oscilador": 730, "Bender": 1100000, "Bow": 1900000,
               "Straightener R1": 1700000, "Straightener R2": 1700000, "Horizontal": 3300000}
SNAPSHOT = datetime.date(2026, 10, 20)  # data em que os acumulados do app foram lidos


def _tipo(t):
    t = (t or "").upper()
    if t.startswith("HORIZONTAL"):
        return "Horizontal"
    if t in ("R1", "STRAIGHTENER R1"):
        return "Straightener R1"
    if t in ("R2", "STRAIGHTENER R2"):
        return "Straightener R2"
    return {"BOW": "Bow", "BENDER": "Bender", "OSCILADOR": "Oscilador", "MOLDE": "Molde"}.get(t, t.title())


def _posicao(tipo, pos):
    m = re.search(r"(\d+)", pos or "")
    if tipo == "Horizontal" and m:
        return f"Horizontal {int(m.group(1)) - 7:02d}"  # HOR-8 = Horizontal 01 ... HOR-17 = Horizontal 10
    if tipo == "Bow" and m:
        return f"Bow {m.group(1)}"
    if tipo == "Oscilador":
        return "Oscilador " + ("Norte" if (pos or "").endswith("N") else "Sul")
    return tipo


def _data(txt):
    try:
        d, m_, a_ = (int(x) for x in txt.split("/"))
        return datetime.date(a_, m_, d)
    except Exception:  # noqa: BLE001
        return None


def carregar_equipamentos():
    """Lista do app (Neon) se DADOS_EQUIP apontar para o snapshot; senão, lista genérica."""
    caminho = os.environ.get("DADOS_EQUIP")
    itens = [dict(tag="Molde G", tipo="Molde", veio="G", meta=1100, pos="Molde (relatório)", id=""),
             dict(tag="Molde H", tipo="Molde", veio="H", meta=1100, pos="Molde (relatório)", id="")]
    if not caminho:
        genericos = ([("Oscilador", "Oscilador")] + [("Bender", "Bender")] + [(f"Bow {n}", "Bow") for n in range(1, 6)]
                     + [("Straightener R1", "Straightener R1"), ("Straightener R2", "Straightener R2")]
                     + [(f"Horizontal {n:02d}", "Horizontal") for n in range(1, 11)])
        for nome, tipo in genericos:
            itens.append(dict(tag=nome, tipo=tipo, veio="", meta=META_PADRAO[tipo], pos=nome, id="",
                              desde=INICIO_DADOS, inicial=None))
        return itens
    spec = importlib.util.spec_from_file_location("dados_equip", caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    ordem = {t: n for n, t in enumerate(TIPOS)}
    lista = []
    for (id_, tag, tipo, local, pos, ton, meta, entrada) in mod.ROWS:
        tipo_n = _tipo(tipo)
        if tipo_n == "Molde":
            continue  # molde vem do relatório diário (Molde G / Molde H)
        veio = local.strip()[-1]
        posicao = _posicao(tipo_n, pos)
        if tipo_n == "Oscilador":  # conta dias desde a entrada
            desde, inicial = _data(entrada) or SNAPSHOT, None
        else:                      # parte do acumulado do app e soma a produção a partir do snapshot
            desde, inicial = SNAPSHOT, ton
        lista.append(dict(tag=tag or id_, tipo=tipo_n, veio=veio, meta=meta, pos=posicao, id=id_,
                          desde=desde, inicial=inicial))
    lista.sort(key=lambda e: (e["veio"], ordem.get(e["tipo"], 99), e["pos"]))
    return itens + lista


equip = carregar_equipamentos()
N = len(equip)
for i in range(2, EQ_ULT + 1):
    k = i - 2
    molde = k < N and equip[k]["tipo"] == "Molde"
    if k < N:
        e = equip[k]
        eq.cell(i, 2, e["tipo"])
        eq.cell(i, 4, None if molde else 4)
        eq.cell(i, 12, e["veio"] or None)
        eq.cell(i, 14, e["pos"])
        eq.cell(i, 15, e["id"] or None)
        eq.cell(i, 19, e["tag"])                      # S: tag original
        eq.cell(i, 20, e["meta"])                     # T: meta original
        eq.cell(i, 21, None if molde else e.get("desde"))    # U: desde original
        eq.cell(i, 22, None if molde else e.get("inicial"))  # V: acumulado inicial original
    ult_hist = f"SUMPRODUCT(MAX((Historico!$B$2:$B${LIM}=L{i})*Historico!$A$2:$A${LIM}))"
    eq[f"C{i}"] = f'=IF(B{i}="","",IF(B{i}="Molde","corridas",IF(B{i}="Oscilador","dias","t")))'
    if molde:
        eq[f"A{i}"] = f"=S{i}"
        eq[f"E{i}"] = f"=IF({ult_hist}=0,T{i},SUMIFS(Historico!$D:$D,Historico!$B:$B,L{i},Historico!$A:$A,{ult_hist}))"
    else:
        # posição -> última troca registrada em TrocasEquip (se houver) define a peça atual
        eq[f"P{i}"] = f'=IF(OR(N{i}="",L{i}=""),"",L{i}&" - "&N{i})'
        eq[f"R{i}"] = f'=IF(P{i}="",0,_xlfn.MAXIFS(TrocasEquip!$A$2:$A$1000,TrocasEquip!$B$2:$B$1000,P{i}))'
        eq[f"W{i}"] = f'=IF(R{i}=0,0,MATCH(1,INDEX((TrocasEquip!$B$2:$B$1000=P{i})*(TrocasEquip!$A$2:$A$1000=R{i}),0),0))'
        eq[f"A{i}"] = f'=IF(S{i}="","",IF(W{i}=0,S{i},INDEX(TrocasEquip!$G$2:$G$1000,W{i})))'
        eq[f"E{i}"] = f'=IF(W{i}=0,T{i},IF(INDEX(TrocasEquip!$H$2:$H$1000,W{i})="",T{i},INDEX(TrocasEquip!$H$2:$H$1000,W{i})))'
        eq[f"F{i}"] = f'=IF(W{i}=0,IF(U{i}="","",U{i}),R{i})'
        eq[f"G{i}"] = f'=IF(W{i}=0,V{i},N(INDEX(TrocasEquip!$I$2:$I$1000,W{i})))'
    eq[f"H{i}"] = (
        f'=IF(A{i}="","",IF(C{i}="corridas",'
        f"SUMIFS(Historico!$E:$E,Historico!$B:$B,L{i},Historico!$A:$A,{ult_hist}),"
        f'N(G{i})+IF(F{i}="",0,IF(C{i}="dias",MAX(0,TODAY()-F{i}),'
        f'IF(D{i}=2,SUMIFS(Producao!$D:$D,Producao!$A:$A,">="&F{i}),'
        f'IF(D{i}=3,SUMIFS(Producao!$E:$E,Producao!$A:$A,">="&F{i}),'
        f'SUMIFS(Producao!$B:$B,Producao!$A:$A,">="&F{i})))/Regras!$B$4))))'
    )
    eq[f"I{i}"] = f'=IF(OR(A{i}="",E{i}=""),"",E{i}-H{i})'
    eq[f"J{i}"] = f'=IF(OR(A{i}="",N(E{i})<=0),"",H{i}/E{i})'
    eq[f"K{i}"] = f'=IF(I{i}="","",IF(I{i}<0,"Além da meta",IF(J{i}>=Regras!$B$3,"Atenção","OK")))'
    eq[f"M{i}"] = f'=IF(J{i}="",-1,J{i}+ROW()/1000000000)'  # chave do ranking
    eq[f"Q{i}"] = f'=IF(A{i}<>"",ROW(),"")'                   # linhas com peça, para o Painel
    for col, fmt in (("E", "#,##0"), ("G", "#,##0"), ("H", "#,##0"), ("I", "#,##0"), ("T", "#,##0"), ("V", "#,##0"),
                     ("F", "dd/mm/yyyy"), ("U", "dd/mm/yyyy"), ("R", 'dd/mm/yyyy;;"-"'), ("J", "0.0%")):
        eq[f"{col}{i}"].number_format = fmt
    for c in list(eq[i][:12]) + [eq.cell(i, j) for j in (14, 15, 16, 18, 19, 20, 21, 22)]:
        c.border = BORDA
        c.alignment = ESQ if c.column in (1, 16, 19) else CENTRO if c.column in (2, 3, 4, 6, 11, 12, 18, 21) else DIR
preencher(eq, f"B2:B{EQ_ULT}", AMARELO)
preencher(eq, f"D2:D{EQ_ULT}", AMARELO)
preencher(eq, f"L2:L{EQ_ULT}", AMARELO)
preencher(eq, f"N2:O{EQ_ULT}", AMARELO)
preencher(eq, f"S2:V{EQ_ULT}", AMARELO)
for k_, e_ in enumerate(equip):
    if e_["tipo"] == "Molde":  # molde vem do relatório: não digitar
        preencher(eq, f"D{k_ + 2}:D{k_ + 2}", CINZA_CLARO)
        preencher(eq, f"U{k_ + 2}:V{k_ + 2}", CINZA_CLARO)
eq["A1"].fill = PatternFill("solid", fgColor=AZUL)
barra_percentual(eq, f"J2:J{EQ_ULT}")
cores_situacao(eq, f"K2:K{EQ_ULT}")
for col in ("M", "Q", "W"):
    eq.column_dimensions[col].hidden = True
notas = [
    ("Uma linha por POSIÇÃO. Tag atual, Meta, Desde e Acumulado inicial são automáticos: vêm da última troca em TrocasEquip (ou dos valores 'original').", True, CINZA),
    ("Amarelo = você digita: Tipo, MCC, Veio, Posição e os valores 'original' (a peça que estava na posição quando a planilha começou).", False, CINZA),
    ("Trocou uma peça? Só na aba TrocasEquip: data, posição, tag que entrou, motivo (meta e acumulado inicial da nova se precisar). Aqui atualiza sozinho.", False, CINZA),
    ("Molde: automático pelo relatório (corridas e vida programada do veio G ou H). Oscilador conta dias; componentes somam as toneladas da MCC desde 'Desde'.", False, CINZA),
    ("ATENÇÃO: os valores 'original' vieram do app em 20/10/2026. Confira tags TESTE, tags repetidas e osciladores.", True, VERMELHO_ESC),
]
for n_, (txt, negrito, cor) in enumerate(notas):
    c = eq.cell(EQ_ULT + 3 + n_, 1, txt)
    c.font = Font(italic=not negrito, bold=negrito, color=cor)
larguras(eq, {"A": 16, "B": 15, "C": 9, "D": 6, "E": 12, "F": 12, "G": 15, "H": 13, "I": 13, "J": 15, "K": 13, "L": 6,
              "N": 17, "O": 15, "P": 22, "R": 13, "S": 14, "T": 13, "U": 14, "V": 18})
dv = DataValidation(type="list", formula1='"' + ",".join(TIPOS) + '"', allow_blank=True)
dv.add(f"B2:B{EQ_ULT}")
eq.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"2,3,4"', allow_blank=True)
dv.add(f"D2:D{EQ_ULT}")
eq.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"C,D,E,F,G,H"', allow_blank=True)
dv.add(f"L2:L{EQ_ULT}")
eq.add_data_validation(dv)
validacao_data(eq, f"U2:U{EQ_ULT}")
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
cabecalho(te, ["Data da troca", "Posição (veio - posição)", "Tag que saiu", "Acumulado na saída", "Unidade",
               "% da vida na saída", "Tag que entrou", "Meta da nova (vazio = mesma)", "Acum. inicial da nova (vazio = 0)",
               "Motivo", "Observação", "data anterior", "idx anterior", "linha posição", "desde anterior",
               "inicial anterior", "meta anterior"], cor=OURO, altura=42)
for r in range(2, ULTIMA + 1):
    te[f"N{r}"] = f'=IF(B{r}="","",IFERROR(MATCH(B{r},{EQ("P")},0),""))'
    te[f"L{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",_xlfn.MAXIFS(TrocasEquip!$A$2:$A$1000,TrocasEquip!$B$2:$B$1000,B{r},TrocasEquip!$A$2:$A$1000,"<"&A{r}))'
    te[f"M{r}"] = f'=IF(N(L{r})=0,0,MATCH(1,INDEX((TrocasEquip!$B$2:$B$1000=B{r})*(TrocasEquip!$A$2:$A$1000=L{r}),0),0))'
    te[f"O{r}"] = f'=IF(N{r}="","",IF(M{r}=0,INDEX({EQ("U")},N{r}),L{r}))'
    te[f"P{r}"] = f'=IF(N{r}="","",IF(M{r}=0,N(INDEX({EQ("V")},N{r})),N(INDEX(TrocasEquip!$I$2:$I$1000,M{r}))))'
    te[f"Q{r}"] = (f'=IF(N{r}="","",IF(M{r}=0,INDEX({EQ("T")},N{r}),'
                   f'IF(INDEX(TrocasEquip!$H$2:$H$1000,M{r})="",INDEX({EQ("T")},N{r}),INDEX(TrocasEquip!$H$2:$H$1000,M{r}))))')
    te[f"C{r}"] = (f'=IF(B{r}="","",IF(N{r}="","posição não encontrada",'
                   f'IF(M{r}=0,INDEX({EQ("S")},N{r}),INDEX(TrocasEquip!$G$2:$G$1000,M{r}))))')
    te[f"E{r}"] = f'=IF(N{r}="","",INDEX({EQ("C")},N{r}))'
    mcc = f'INDEX({EQ("D")},N{r})'
    ate = f'Producao!$A:$A,">="&O{r},Producao!$A:$A,"<"&A{r}'
    te[f"D{r}"] = (f'=IF(OR(A{r}="",N{r}=""),"",P{r}+IF(E{r}="dias",MAX(0,A{r}-O{r}),'
                   f'IF({mcc}=2,SUMIFS(Producao!$D:$D,{ate}),IF({mcc}=3,SUMIFS(Producao!$E:$E,{ate}),'
                   f'SUMIFS(Producao!$B:$B,{ate})))/Regras!$B$4))')
    te[f"F{r}"] = f'=IF(OR(D{r}="",N(Q{r})<=0),"",D{r}/Q{r})'
    te[f"R{r}"] = f'=IF(N{r}="","",INDEX({EQ("B")},N{r}))'
    te[f"A{r}"].number_format = "dd/mm/yyyy"
    for col in ("D", "H", "I"):
        te[f"{col}{r}"].number_format = "#,##0"
    te[f"F{r}"].number_format = "0.0%"
preencher(te, f"A2:B{ULTIMA}", AMARELO)
preencher(te, f"G2:K{ULTIMA}", AMARELO)
for col in "LMNOPQR":
    te.column_dimensions[col].hidden = True
larguras(te, {"A": 13, "B": 24, "C": 14, "D": 18, "E": 9, "F": 15, "G": 15, "H": 16, "I": 18, "J": 35, "K": 30})
dv = DataValidation(type="list", formula1=f"={EQ('P')}", allow_blank=True)
dv.add(f"B2:B{ULTIMA}")
te.add_data_validation(dv)
validacao_data(te, f"A2:A{ULTIMA}")
# Instruções ao lado da tabela (colunas S:T)
te.column_dimensions["S"].width = 5
te.column_dimensions["T"].width = 78
te.merge_cells("S1:T1")
te["S1"] = "COMO PREENCHER UMA TROCA (uma linha por troca)"
te["S1"].font = Font(bold=True, color="FFFFFF", size=12)
te["S1"].fill = PatternFill("solid", fgColor=OURO)
te["S1"].alignment = ESQ
passos_te = [
    ("1", "Data da troca: o dia em que a peça foi trocada (dd/mm/aaaa)."),
    ("2", "Posição: escolha na lista. Ex.: \"H - Horizontal 10\" = veio H, horizontal 10."),
    ("3", "Tag que entrou: a tag da peça nova (ex.: H-30)."),
    ("4", "Meta da nova: a VIDA TOTAL da peça (ex.: 3.300.000). NÃO é quanto falta. Vazio = mesma meta da posição."),
    ("5", "Acum. inicial da nova: quanto a peça JÁ RODOU até a troca (ex.: 1.359.462). Se você só sabe quanto falta: meta − falta. Vazio = zero."),
    ("6", "Motivo e Observação: o porquê da troca (fim de vida, quebra, reparo...)."),
    ("", ""),
    ("✔", "Aparecem sozinhos: Tag que saiu, Acumulado na saída, Unidade e % da vida na saída."),
    ("✔", "A aba Equipamentos e o Painel passam a mostrar a peça nova na posição. Não precisa mexer lá."),
    ("!", "Molde NÃO entra aqui: a troca de molde o robô pega sozinho do relatório (aba Trocas)."),
    ("!", "Errou? Corrija a própria linha. Não apague trocas antigas: elas guardam o histórico da posição."),
    ("", ""),
    ("Ex.", "20/10/2026  |  H - Horizontal 10  |  H-30  |  (vazio)  |  (vazio)  |  fim de vida"),
]
for n_, (num, txt) in enumerate(passos_te, start=2):
    a_ = te.cell(n_, 19, num)
    b_ = te.cell(n_, 20, txt)
    a_.font = Font(bold=True, color=OURO if num.isdigit() else ("006100" if num == "✔" else VERMELHO_ESC if num == "!" else NAVY))
    a_.alignment = CENTRO
    b_.alignment = Alignment(wrap_text=True, vertical="center")
    if txt:
        for c_ in (a_, b_):
            c_.fill = PatternFill("solid", fgColor=CINZA_CLARO)
            c_.border = BORDA
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
               "Molde que entrou", "Motivo", "Conferir", "Origem", "% da vida na saída"], cor=CINZA)
for r in range(2, 2001):
    tr[f"J{r}"] = f'=IF(OR(E{r}="",N(D{r})=0),"",E{r}/D{r})'
    tr[f"J{r}"].number_format = "0.0%"
larguras(tr, {"A": 13, "B": 7, "C": 15, "D": 16, "E": 17, "F": 16, "G": 60, "H": 45, "I": 10, "J": 17})
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
rg["A4"], rg["B4"] = "Veios por máquina (a tonelada do relatório é da máquina e é dividida entre os veios)", 2
rg["B4"].fill = PatternFill("solid", fgColor=AMARELO)
rg["B4"].border = BORDA
rg["B3"].number_format = "0%"
rg["B3"].fill = PatternFill("solid", fgColor=AMARELO)
rg["B3"].border = BORDA
rg["A5"] = "Como a planilha calcula"
rg["A5"].font = Font(bold=True, size=12, color=NAVY)
textos = [
    "Molde (corridas): acumulado do último dia em Historico, pelo veio. Um registro por data e veio; reenvio do mesmo dia substitui, não soma.",
    "Troca de molde: quando o nº do molde do veio muda, o robô registra em Trocas (molde que saiu, corridas na saída pela observação do relatório, molde que entrou).",
    "Oscilador (dias): dias corridos desde 'Desde' + Acumulado inicial.",
    "Componentes (t): toneladas reais da MCC escolhida (aba Producao) desde 'Desde', DIVIDIDAS pelos veios da máquina (Regras B4 = 2: cada veio fez metade), + Acumulado inicial.",
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
for k in range(1, N + 7):
    ax[f"H{k + 1}"] = f"=IFERROR(SMALL({EQ('Q')},{k}),0)"
ax["H1"] = "Linha ativa"
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
    ("A", "A", "Equipamentos", f"=COUNT({EQ('Q')})", AZUL_CLARO, NAVY),
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
    pn.cell(r, 7, f'=COUNTIFS({EQ("B")},F{r},{EQ("A")},"<>")')
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
# Vida real na troca (média por tipo)
V0 = 18
titulo_secao(pn, f"A{V0}:I{V0}", "Vida real na troca (média das trocas registradas)")
cab_linha(pn, V0 + 1, 1, ["Tipo", "Trocas", "Média na saída", "Unidade", "Meta média", "% da meta (média)"])
TE_ = lambda col: f"TrocasEquip!${col}$2:${col}${ULTIMA}"  # noqa: E731
TR_ = lambda col: f"Trocas!${col}$2:${col}$2000"  # noqa: E731
for n_, tipo in enumerate(TIPOS):
    r = V0 + 2 + n_
    if tipo == "Molde":
        pn.cell(r, 1, "Molde (G e H)")
        q = f'(COUNTIFS({TR_("B")},"G",{TR_("E")},">=0")+COUNTIFS({TR_("B")},"H",{TR_("E")},">=0"))'
        soma = lambda col: f'(SUMIFS({TR_(col)},{TR_("B")},"G")+SUMIFS({TR_(col)},{TR_("B")},"H"))'  # noqa: E731
        pn.cell(r, 2, f"={q}")
        pn.cell(r, 3, f'=IF(B{r}=0,"-",{soma("E")}/B{r})')
        pn.cell(r, 4, "corridas")
        pn.cell(r, 5, f'=IF(B{r}=0,"-",{soma("D")}/B{r})')
        pn.cell(r, 6, f'=IF(B{r}=0,"-",{soma("J")}/B{r})')
    else:
        pn.cell(r, 1, tipo)
        pn.cell(r, 2, f'=COUNTIFS({TE_("R")},A{r},{TE_("D")},">=0")')
        pn.cell(r, 3, f'=IFERROR(AVERAGEIFS({TE_("D")},{TE_("R")},A{r}),"-")')
        pn.cell(r, 4, "dias" if tipo == "Oscilador" else "t")
        pn.cell(r, 5, f'=IFERROR(AVERAGEIFS({TE_("Q")},{TE_("R")},A{r},{TE_("D")},">=0"),"-")')
        pn.cell(r, 6, f'=IFERROR(AVERAGEIFS({TE_("F")},{TE_("R")},A{r}),"-")')
    for j in range(1, 7):
        c = pn.cell(r, j)
        c.border = BORDA
        c.alignment = ESQ if j == 1 else CENTRO
    pn.cell(r, 3).number_format = "#,##0"
    pn.cell(r, 5).number_format = "#,##0"
    pn.cell(r, 6).number_format = "0.0%"
barra_percentual(pn, f"F{V0 + 2}:F{V0 + 1 + len(TIPOS)}")
pn.cell(V0 + 2 + len(TIPOS), 1, "Abaixo de 100% = a peça sai antes da meta (vida real menor que a prevista). Acima de 100% = durou mais que a meta.").font = Font(italic=True, size=9, color=CINZA)

T0 = V0 + len(TIPOS) + 4  # linha do título da seção
PN = N + 6
titulo_secao(pn, f"A{T0}:I{T0}", "Todos os equipamentos")
cab_linha(pn, T0 + 1, 1, ["Tag", "Posição", "Veio", "Meta", "Acumulado", "Falta", "% da vida", "Situação", "Desde"])
origem = {1: "A", 2: "N", 3: "L", 4: "E", 5: "H", 6: "I", 7: "J", 8: "K", 9: "F"}
for k in range(PN):
    r = T0 + 2 + k
    for j, col in origem.items():
        idx = f"Auxiliar!$H${k + 2}"
        ref = f"INDEX(Equipamentos!${col}:${col},{idx})"
        f = f'=IF({idx}=0,"",IF({ref}="","",{ref}))'
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
c1.height, c1.width = max(13, PN * 0.42), 12.5
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
    ("TrocasEquip", "Uma linha por troca: data, posição, tag que entrou e motivo. Quem saiu e com quanto saiu é automático.", OURO),
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
        "Oscilador ou componente: só na aba TrocasEquip, uma linha: Data da troca, Posição (escolha na lista), Tag que entrou e Motivo.",
        "Se a peça nova tiver outra meta ou já tiver uso, preencha também 'Meta da nova' e 'Acum. inicial da nova'.",
        "A planilha mostra sozinha a tag que saiu, o acumulado e a % da vida na saída, e coloca a peça nova na posição (aba Equipamentos e Painel).",
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
