' atualizar_controle.vbs
' Le os e-mails "Relatorios GLA" do Outlook CLASSICO, extrai os anexos
' "VIDA MOLDE" e "CORTESEQ" e atualiza a planilha de controle.
'
' - Historico: um registro por data e veio (reenvio do mesmo dia SUBSTITUI, nao soma)
' - Producao : toneladas reais por MCC por dia (aba DADOS DA PRODUCAO)
' - Trocas   : registrada quando o numero do molde do veio muda
' - Alertas  : acumulado diminuiu sem troca / veio nao andou / layout mudou
'
' Nao move nem marca nenhum e-mail no Outlook. Guarda o ID de cada e-mail ja
' processado em processados.txt (funciona lido ou nao lido).
' Feche a planilha antes de rodar. Comece com DRY_RUN = True (nao grava nada,
' so escreve robo.log) e troque para False depois de conferir o log.
Option Explicit

' ===== CONFIGURACAO =====
Const PASTA = "C:\Controle"
Const PLANILHA = "C:\Controle\controle_equipamentos.xlsx"
Const ASSUNTO = "GLA"
Const DIAS_RETROATIVOS = 45
Const DRY_RUN = True
' ========================

Dim fso, xl, wb, logPath
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = PASTA & "\robo.log"

Sub Registrar(msg)
  Dim f
  On Error Resume Next
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & " | " & msg
  f.Close
  On Error GoTo 0
End Sub

Function UltimaLinha(ws, col)
  UltimaLinha = ws.Cells(ws.Rows.Count, col).End(-4162).Row
End Function

Function Txt(v)
  Txt = Trim(CStr(v & ""))
End Function

Function AchaAba(book, trecho)
  Dim s
  Set AchaAba = Nothing
  For Each s In book.Worksheets
    If InStr(UCase(s.Name), trecho) > 0 Then
      Set AchaAba = s
      Exit Function
    End If
  Next
End Function

Sub Alerta(tipo, item, detalhe)
  Dim ws, r, ult
  Registrar "ALERTA [" & tipo & "] " & item & ": " & detalhe
  If DRY_RUN Then Exit Sub
  Set ws = wb.Worksheets("Alertas")
  ult = UltimaLinha(ws, 1)
  For r = 2 To ult
    If Txt(ws.Cells(r, 2).Value) = tipo And Txt(ws.Cells(r, 3).Value) = item And Txt(ws.Cells(r, 4).Value) = detalhe Then Exit Sub
  Next
  ws.Cells(ult + 1, 1).Value = Now
  ws.Cells(ult + 1, 2).Value = tipo
  ws.Cells(ult + 1, 3).Value = item
  ws.Cells(ult + 1, 4).Value = detalhe
End Sub

' ---- Producao -------------------------------------------------------------
Sub GravarProducao(dt, mcc4real, mcc4prog, mcc2real, mcc3real, arquivo)
  Dim ws, r, ult, achou, k
  Registrar "Producao " & dt & " MCC4 real=" & mcc4real
  If DRY_RUN Then Exit Sub
  Set ws = wb.Worksheets("Producao")
  k = CLng(dt)
  ult = UltimaLinha(ws, 1)
  achou = 0
  For r = 2 To ult
    If IsDate(ws.Cells(r, 1).Value) Then
      If CLng(CDate(ws.Cells(r, 1).Value)) = k Then
        achou = r
        Exit For
      End If
    End If
  Next
  If achou = 0 Then achou = ult + 1
  ws.Cells(achou, 1).Value = CDate(dt)
  ws.Cells(achou, 1).NumberFormat = "dd/mm/yyyy"
  ws.Cells(achou, 2).Value = mcc4real
  ws.Cells(achou, 3).Value = mcc4prog
  ws.Cells(achou, 4).Value = mcc2real
  ws.Cells(achou, 5).Value = mcc3real
  ws.Cells(achou, 6).Value = arquivo
End Sub

' Toneladas reais da MCC do veio na data; -1 se nao houver registro
Function ProducaoDoVeio(veio, dt)
  Dim ws, r, ult, col
  ProducaoDoVeio = -1
  Select Case veio
    Case "C", "D": col = 4
    Case "E", "F": col = 5
    Case Else: col = 2
  End Select
  Set ws = wb.Worksheets("Producao")
  ult = UltimaLinha(ws, 1)
  For r = 2 To ult
    If IsDate(ws.Cells(r, 1).Value) Then
      If CLng(CDate(ws.Cells(r, 1).Value)) = CLng(dt) Then
        If Txt(ws.Cells(r, col).Value) <> "" Then ProducaoDoVeio = CDbl(ws.Cells(r, col).Value)
        Exit Function
      End If
    End If
  Next
End Function

' ---- Historico e Trocas -----------------------------------------------------
' Linha do ultimo registro do veio com data anterior a dt (0 se nao existir)
Function LinhaAnterior(veio, dt)
  Dim ws, r, ult, melhor, dataMelhor, d
  Set ws = wb.Worksheets("Historico")
  ult = UltimaLinha(ws, 1)
  melhor = 0
  dataMelhor = 0
  For r = 2 To ult
    If IsDate(ws.Cells(r, 1).Value) Then
      d = CLng(CDate(ws.Cells(r, 1).Value))
      If UCase(Txt(ws.Cells(r, 2).Value)) = veio And d < CLng(dt) And d > dataMelhor Then
        melhor = r
        dataMelhor = d
      End If
    End If
  Next
  LinhaAnterior = melhor
End Function

Sub GravarHistorico(dt, veio, molde, meta, acum, obs)
  Dim ws, r, ult, achou, k
  If DRY_RUN Then Exit Sub
  Set ws = wb.Worksheets("Historico")
  k = CLng(dt)
  ult = UltimaLinha(ws, 1)
  achou = 0
  For r = 2 To ult
    If IsDate(ws.Cells(r, 1).Value) Then
      If CLng(CDate(ws.Cells(r, 1).Value)) = k And UCase(Txt(ws.Cells(r, 2).Value)) = veio Then
        achou = r
        Exit For
      End If
    End If
  Next
  If achou = 0 Then achou = ult + 1
  ws.Cells(achou, 1).Value = CDate(dt)
  ws.Cells(achou, 1).NumberFormat = "dd/mm/yyyy"
  ws.Cells(achou, 2).Value = veio
  ws.Cells(achou, 3).Value = molde
  ws.Cells(achou, 4).Value = meta
  ws.Cells(achou, 5).Value = acum
  ws.Cells(achou, 6).Value = obs
End Sub

Sub RegistrarTroca(dt, veio, moldeSaiu, vidaProg, acumSaida, moldeEntrou, obs)
  Dim ws, r, ult, conferir, re, m
  Registrar "TROCA veio " & veio & ": saiu " & moldeSaiu & " com " & acumSaida & ", entrou " & moldeEntrou
  conferir = ""
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "MOLDE\s+(?:N\S?\s*)?(\d+)\s+COM\s+(\d+)\s+COR"
  re.IgnoreCase = True
  Set m = Nothing
  If re.Test(obs) Then Set m = re.Execute(obs)(0)
  If m Is Nothing Then
    conferir = "Observacao nao informa molde/corridas de saida"
  Else
    If CLng(m.SubMatches(0)) <> CLng(moldeSaiu) Then conferir = "Observacao cita molde " & m.SubMatches(0) & ", planilha indica " & moldeSaiu
    If CLng(m.SubMatches(1)) <> CLng(acumSaida) Then
      If conferir <> "" Then conferir = conferir & "; "
      conferir = conferir & "Observacao diz " & m.SubMatches(1) & " corridas, ultimo acumulado " & acumSaida
    End If
  End If
  If DRY_RUN Then Exit Sub
  Set ws = wb.Worksheets("Trocas")
  ult = UltimaLinha(ws, 1)
  For r = 2 To ult
    If IsDate(ws.Cells(r, 1).Value) Then
      If CLng(CDate(ws.Cells(r, 1).Value)) = CLng(dt) And UCase(Txt(ws.Cells(r, 2).Value)) = veio And Txt(ws.Cells(r, 6).Value) = Txt(moldeEntrou) Then Exit Sub
    End If
  Next
  ws.Cells(ult + 1, 1).Value = CDate(dt)
  ws.Cells(ult + 1, 1).NumberFormat = "dd/mm/yyyy"
  ws.Cells(ult + 1, 2).Value = veio
  ws.Cells(ult + 1, 3).Value = moldeSaiu
  ws.Cells(ult + 1, 4).Value = vidaProg
  ws.Cells(ult + 1, 5).Value = acumSaida
  ws.Cells(ult + 1, 6).Value = moldeEntrou
  ws.Cells(ult + 1, 7).Value = obs
  ws.Cells(ult + 1, 8).Value = conferir
  ws.Cells(ult + 1, 9).Value = "robo"
  If conferir <> "" Then Alerta "Troca para conferir", "Veio " & veio, conferir
End Sub

Sub AplicarVeio(dt, veio, molde, meta, acum, obs)
  Dim ws, prev, pMolde, pAcum, pMeta, pData, prod
  Set ws = wb.Worksheets("Historico")
  prev = LinhaAnterior(veio, dt)
  If prev > 0 Then
    pMolde = Txt(ws.Cells(prev, 3).Value)
    pMeta = ws.Cells(prev, 4).Value
    pAcum = CDbl(ws.Cells(prev, 5).Value)
    pData = CDate(ws.Cells(prev, 1).Value)
    If pMolde <> Txt(molde) Then
      RegistrarTroca dt, veio, pMolde, pMeta, pAcum, Txt(molde), obs
    ElseIf CDbl(acum) < pAcum Then
      Alerta "Acumulado diminuiu sem troca", "Veio " & veio, "Dia " & dt & ": " & acum & " (antes " & pAcum & " em " & pData & ", molde " & molde & ")"
    ElseIf CDbl(acum) = pAcum Then
      prod = ProducaoDoVeio(veio, dt - 1)
      If prod > 0 Then Alerta "Veio nao andou", "Veio " & veio, "Acumulado " & acum & " igual ao de " & pData & ", mas a MCC produziu " & prod & " t em " & (dt - 1) & ". Verificar parada ou e-mail repetido."
    End If
  End If
  GravarHistorico dt, veio, molde, meta, acum, obs
End Sub

' ---- Leitura dos anexos -----------------------------------------------------
' Retorna True se leu sem problema de layout
Function ProcessarVidaMolde(caminho, ByRef dataRel)
  Dim b, ws, r, veio, molde, meta, acum, obs, ok
  ProcessarVidaMolde = False
  Set b = xl.Workbooks.Open(caminho, 0, True)
  Set ws = AchaAba(b, "VIDA MOLDE")
  If ws Is Nothing Then
    Alerta "Layout alterado", "VIDA MOLDE", "Aba VIDA MOLDE nao encontrada em " & fso.GetFileName(caminho)
    b.Close False
    Exit Function
  End If
  If UCase(Txt(ws.Cells(3, 2).Value)) <> "VEIOS" Or InStr(UCase(Txt(ws.Cells(4, 4).Value)), "PROGR") = 0 Or UCase(Txt(ws.Cells(4, 5).Value)) <> "REAL" Or Not IsDate(ws.Cells(2, 9).Value) Then
    Alerta "Layout alterado", "VIDA MOLDE", "Cabecalhos (B3, D4, E4) ou data (I2) fora do esperado em " & fso.GetFileName(caminho)
    b.Close False
    Exit Function
  End If
  dataRel = CDate(ws.Cells(2, 9).Value)
  ok = True
  For r = 5 To 10
    veio = UCase(Txt(ws.Cells(r, 2).Value))
    molde = ws.Cells(r, 3).Value
    meta = ws.Cells(r, 4).Value
    acum = ws.Cells(r, 5).Value
    obs = Txt(ws.Cells(r, 8).Value)
    If Len(veio) <> 1 Or InStr("CDEFGH", veio) = 0 Or Not IsNumeric(molde) Or Not IsNumeric(meta) Or Not IsNumeric(acum) Or Txt(molde) = "" Or Txt(acum) = "" Then
      Alerta "Layout alterado", "VIDA MOLDE", "Linha " & r & " invalida (veio=" & veio & ", molde=" & Txt(molde) & ", acumulado=" & Txt(acum) & ")"
      ok = False
    Else
      AplicarVeio dataRel, veio, CLng(molde), CDbl(meta), CDbl(acum), obs
    End If
  Next
  b.Close False
  ProcessarVidaMolde = ok
End Function

Function ProcessarCorteseq(caminho, nomeArq, refData)
  Dim b, ws, r, dia, ano, mes, dt, re, m, real4, prog4, real2, real3
  ProcessarCorteseq = False
  ano = Year(refData)
  mes = Month(refData)
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "(\d{1,2})-(\d{4})"
  If re.Test(nomeArq) Then
    Set m = re.Execute(nomeArq)(0)
    mes = CInt(m.SubMatches(0))
    ano = CInt(m.SubMatches(1))
  End If
  Set b = xl.Workbooks.Open(caminho, 0, True)
  Set ws = AchaAba(b, "DADOS DA PRODU")
  If ws Is Nothing Then
    Alerta "Layout alterado", "CORTESEQ", "Aba DADOS DA PRODUCAO nao encontrada em " & nomeArq
    b.Close False
    Exit Function
  End If
  If UCase(Txt(ws.Cells(3, 1).Value)) <> "DIA" Or InStr(UCase(Txt(ws.Cells(2, 8).Value)), "MCC#4") = 0 Then
    Alerta "Layout alterado", "CORTESEQ", "Cabecalhos (A3, H2) fora do esperado em " & nomeArq
    b.Close False
    Exit Function
  End If
  For r = 4 To 34
    If IsNumeric(ws.Cells(r, 1).Value) And Txt(ws.Cells(r, 1).Value) <> "" Then
      dia = CInt(ws.Cells(r, 1).Value)
      If dia >= 1 And dia <= 31 And Txt(ws.Cells(r, 9).Value) <> "" Then
        dt = DateSerial(ano, mes, dia)
        If Month(dt) = mes Then
          real4 = CDbl(ws.Cells(r, 9).Value)
          prog4 = 0
          If IsNumeric(ws.Cells(r, 8).Value) And Txt(ws.Cells(r, 8).Value) <> "" Then prog4 = CDbl(ws.Cells(r, 8).Value)
          real2 = ""
          If IsNumeric(ws.Cells(r, 3).Value) And Txt(ws.Cells(r, 3).Value) <> "" Then real2 = CDbl(ws.Cells(r, 3).Value)
          real3 = ""
          If IsNumeric(ws.Cells(r, 6).Value) And Txt(ws.Cells(r, 6).Value) <> "" Then real3 = CDbl(ws.Cells(r, 6).Value)
          GravarProducao dt, real4, prog4, real2, real3, nomeArq
        End If
      End If
    End If
  Next
  b.Close False
  ProcessarCorteseq = True
End Function

' ---- Carga do que ja foi processado -----------------------------------------
Function CarregarProcessados()
  Dim d, f, linha, p
  Set d = CreateObject("Scripting.Dictionary")
  p = PASTA & "\processados.txt"
  If fso.FileExists(p) Then
    Set f = fso.OpenTextFile(p, 1)
    Do While Not f.AtEndOfStream
      linha = Trim(f.ReadLine)
      If linha <> "" Then d(linha) = 1
    Loop
    f.Close
  End If
  Set CarregarProcessados = d
End Function

Sub MarcarProcessado(id)
  Dim f
  If DRY_RUN Then Exit Sub
  Set f = fso.OpenTextFile(PASTA & "\processados.txt", 8, True)
  f.WriteLine id
  f.Close
End Sub

' ---- Processa um e-mail -----------------------------------------------------
Function ProcessarEmail(msg, tmp, seq)
  Dim i, att, nome, u, pVida, pCorte, nCorte, ok, dataRel, temVida, temCorte
  ProcessarEmail = False
  temVida = False
  temCorte = False
  For i = 1 To msg.Attachments.Count
    Set att = msg.Attachments.Item(i)
    nome = att.FileName
    u = UCase(nome)
    If Right(u, 5) = ".XLSX" Or Right(u, 4) = ".XLS" Then
      If InStr(u, "VIDA MOLDE") > 0 And Not temVida Then
        pVida = tmp & "\" & seq & "_vida_" & i & Right(u, Len(u) - InStrRev(u, ".") + 1)
        att.SaveAsFile pVida
        temVida = True
      ElseIf InStr(u, "CORTESEQ") > 0 And Not temCorte Then
        pCorte = tmp & "\" & seq & "_corte_" & i & Right(u, Len(u) - InStrRev(u, ".") + 1)
        att.SaveAsFile pCorte
        nCorte = nome
        temCorte = True
      End If
    End If
  Next
  If Not temVida And Not temCorte Then Exit Function
  ok = True
  dataRel = CDate(msg.ReceivedTime)
  ' Vida molde primeiro para saber a data do relatorio; producao depois
  If temVida Then
    If Not ProcessarVidaMolde(pVida, dataRel) Then ok = False
  End If
  If temCorte Then
    If Not ProcessarCorteseq(pCorte, nCorte, dataRel) Then ok = False
  End If
  If temVida Then fso.DeleteFile pVida, True
  If temCorte Then fso.DeleteFile pCorte, True
  ProcessarEmail = ok
End Function

' ---- Principal ---------------------------------------------------------------
Sub Main()
  Dim ol, ns, pasta, itens, msg, proc, tmp, seq, nOk, nFalha, nPulados, filtro, corte, res

  If Not fso.FolderExists(PASTA) Then fso.CreateFolder PASTA
  tmp = PASTA & "\temp"
  If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp
  If DRY_RUN Then
    Registrar "==== Inicio (DRY_RUN: nada sera gravado)"
  Else
    Registrar "==== Inicio"
  End If

  If Not fso.FileExists(PLANILHA) Then
    Registrar "ERRO: planilha nao encontrada em " & PLANILHA
    Exit Sub
  End If

  Set xl = CreateObject("Excel.Application")
  xl.Visible = False
  xl.DisplayAlerts = False
  Set wb = xl.Workbooks.Open(PLANILHA)
  If wb.ReadOnly Then
    Registrar "ERRO: planilha aberta por outra pessoa/programa (somente leitura). Feche e rode de novo."
    wb.Close False
    xl.Quit
    Exit Sub
  End If

  Set ol = CreateObject("Outlook.Application")
  Set ns = ol.GetNamespace("MAPI")
  Set pasta = ns.GetDefaultFolder(6)
  filtro = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%" & ASSUNTO & "%'"
  Set itens = pasta.Items.Restrict(filtro)
  itens.Sort "[ReceivedTime]", False
  Set proc = CarregarProcessados()
  corte = Date - DIAS_RETROATIVOS
  seq = 0
  nOk = 0
  nFalha = 0
  nPulados = 0

  For Each msg In itens
    If msg.Class = 43 Then
      If CDate(msg.ReceivedTime) >= corte And Not proc.Exists(msg.EntryID) Then
        seq = seq + 1
        Registrar "E-mail " & msg.ReceivedTime & " | " & msg.Subject
        res = False
        On Error Resume Next
        res = ProcessarEmail(msg, tmp, seq)
        If Err.Number <> 0 Then
          Registrar "ERRO no e-mail: " & Err.Description
          Err.Clear
          res = False
        End If
        On Error GoTo 0
        If res Then
          MarcarProcessado msg.EntryID
          nOk = nOk + 1
        Else
          nFalha = nFalha + 1
        End If
      Else
        nPulados = nPulados + 1
      End If
    End If
  Next

  If Not DRY_RUN Then wb.Save
  wb.Close False
  xl.Quit
  Registrar "==== Fim: processados=" & nOk & " com problema/sem anexos=" & nFalha & " ja feitos/antigos=" & nPulados
End Sub

Main
WScript.Quit 0
