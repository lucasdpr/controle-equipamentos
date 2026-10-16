Attribute VB_Name = "Robo"
' Robo.bas - roda DENTRO do Excel (macro). Nao usa arquivo .vbs.
' Autor: Lucas Gabriel de Paula Rafael | Matricula cbk3574 | github.com/lucasdpr/controle-equipamentos
' Le os e-mails "Relatorios GLA" do Outlook CLASSICO, extrai os anexos
' "VIDA MOLDE" e "CORTESEQ" e atualiza esta planilha.
'
' - Historico: um registro por data e veio (reenvio do mesmo dia SUBSTITUI, nao soma)
' - Producao : toneladas reais por MCC por dia (aba DADOS DA PRODUCAO)
' - Trocas   : registrada quando o numero do molde do veio muda
' - Alertas  : acumulado diminuiu sem troca / veio nao andou / layout mudou
' - Log      : aba criada automaticamente com o que o robo fez
' - Processados: aba criada automaticamente com os e-mails ja lidos
'
' Nao move nem marca e-mails no Outlook. Funciona lido ou nao lido.
' Comece com DRY_RUN = True (nao grava nos dados, so escreve na aba Log).
Option Explicit

Private Const ASSUNTO As String = "GLA"
Private Const DIAS_RETROATIVOS As Long = 45
Private Const DRY_RUN As Boolean = True
Private Const MARCA As String = "Lucas Gabriel de Paula Rafael | Matricula cbk3574 | github.com/lucasdpr/controle-equipamentos"

Private wb As Workbook

' ---- Utilitarios -----------------------------------------------------------
Private Function AbaOuCria(ByVal nome As String, ByVal cab As String) As Worksheet
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = nome Then
            Set AbaOuCria = s
            Exit Function
        End If
    Next s
    Set s = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    s.Name = nome
    s.Cells(1, 1).Value = cab
    s.Cells(1, 1).Font.Bold = True
    Set AbaOuCria = s
End Function

Private Sub Registrar(ByVal msg As String)
    Dim ws As Worksheet, r As Long
    Set ws = AbaOuCria("Log", "Data/hora")
    ws.Cells(1, 2).Value = "Mensagem"
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(r, 2).Value = msg
End Sub

Private Function UltimaLinha(ByVal ws As Worksheet, ByVal col As Long) As Long
    UltimaLinha = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Private Function Txt(ByVal v As Variant) As String
    Txt = Trim$(CStr(v & ""))
End Function

Private Function AchaAba(ByVal book As Workbook, ByVal trecho As String) As Worksheet
    Dim s As Worksheet
    For Each s In book.Worksheets
        If InStr(UCase$(s.Name), trecho) > 0 Then
            Set AchaAba = s
            Exit Function
        End If
    Next s
End Function

Private Sub Alerta(ByVal tipo As String, ByVal item As String, ByVal detalhe As String)
    Dim ws As Worksheet, r As Long, ult As Long
    Registrar "ALERTA [" & tipo & "] " & item & ": " & detalhe
    If DRY_RUN Then Exit Sub
    Set ws = wb.Worksheets("Alertas")
    ult = UltimaLinha(ws, 1)
    For r = 2 To ult
        If Txt(ws.Cells(r, 2).Value) = tipo And Txt(ws.Cells(r, 3).Value) = item And Txt(ws.Cells(r, 4).Value) = detalhe Then Exit Sub
    Next r
    ws.Cells(ult + 1, 1).Value = Now
    ws.Cells(ult + 1, 2).Value = tipo
    ws.Cells(ult + 1, 3).Value = item
    ws.Cells(ult + 1, 4).Value = detalhe
End Sub

' ---- Producao ---------------------------------------------------------------
Private Sub GravarProducao(ByVal dt As Date, ByVal mcc4real As Double, ByVal mcc4prog As Double, ByVal mcc2real As Variant, ByVal mcc3real As Variant, ByVal arquivo As String)
    Dim ws As Worksheet, r As Long, ult As Long, achou As Long
    Registrar "Producao " & Format$(dt, "dd/mm/yyyy") & " MCC4 real=" & mcc4real
    If DRY_RUN Then Exit Sub
    Set ws = wb.Worksheets("Producao")
    ult = UltimaLinha(ws, 1)
    achou = 0
    For r = 2 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            If CLng(CDate(ws.Cells(r, 1).Value)) = CLng(dt) Then
                achou = r
                Exit For
            End If
        End If
    Next r
    If achou = 0 Then achou = ult + 1
    ws.Cells(achou, 1).Value = dt
    ws.Cells(achou, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(achou, 2).Value = mcc4real
    ws.Cells(achou, 3).Value = mcc4prog
    ws.Cells(achou, 4).Value = mcc2real
    ws.Cells(achou, 5).Value = mcc3real
    ws.Cells(achou, 6).Value = arquivo
End Sub

' Toneladas reais da MCC do veio na data; -1 se nao houver registro
Private Function ProducaoDoVeio(ByVal veio As String, ByVal dt As Date) As Double
    Dim ws As Worksheet, r As Long, ult As Long, col As Long
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
    Next r
End Function

' ---- Historico e Trocas -------------------------------------------------------
Private Function LinhaAnterior(ByVal veio As String, ByVal dt As Date) As Long
    Dim ws As Worksheet, r As Long, ult As Long, melhor As Long, dataMelhor As Long, d As Long
    Set ws = wb.Worksheets("Historico")
    ult = UltimaLinha(ws, 1)
    melhor = 0
    dataMelhor = 0
    For r = 2 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            d = CLng(CDate(ws.Cells(r, 1).Value))
            If UCase$(Txt(ws.Cells(r, 2).Value)) = veio And d < CLng(dt) And d > dataMelhor Then
                melhor = r
                dataMelhor = d
            End If
        End If
    Next r
    LinhaAnterior = melhor
End Function

Private Sub GravarHistorico(ByVal dt As Date, ByVal veio As String, ByVal molde As Long, ByVal meta As Double, ByVal acum As Double, ByVal obs As String)
    Dim ws As Worksheet, r As Long, ult As Long, achou As Long
    If DRY_RUN Then Exit Sub
    Set ws = wb.Worksheets("Historico")
    ult = UltimaLinha(ws, 1)
    achou = 0
    For r = 2 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            If CLng(CDate(ws.Cells(r, 1).Value)) = CLng(dt) And UCase$(Txt(ws.Cells(r, 2).Value)) = veio Then
                achou = r
                Exit For
            End If
        End If
    Next r
    If achou = 0 Then achou = ult + 1
    ws.Cells(achou, 1).Value = dt
    ws.Cells(achou, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(achou, 2).Value = veio
    ws.Cells(achou, 3).Value = molde
    ws.Cells(achou, 4).Value = meta
    ws.Cells(achou, 5).Value = acum
    ws.Cells(achou, 6).Value = obs
End Sub

Private Sub RegistrarTroca(ByVal dt As Date, ByVal veio As String, ByVal moldeSaiu As String, ByVal vidaProg As Variant, ByVal acumSaida As Double, ByVal moldeEntrou As String, ByVal obs As String)
    Dim ws As Worksheet, r As Long, ult As Long, conferir As String, re As Object, m As Object
    Registrar "TROCA veio " & veio & ": saiu " & moldeSaiu & " com " & acumSaida & ", entrou " & moldeEntrou
    conferir = ""
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "MOLDE\s+(?:N\S?\s*)?(\d+)\s+COM\s+(\d+)\s+COR"
    re.IgnoreCase = True
    If re.Test(obs) Then
        Set m = re.Execute(obs)(0)
        If CLng(m.SubMatches(0)) <> CLng(moldeSaiu) Then conferir = "Observacao cita molde " & m.SubMatches(0) & ", planilha indica " & moldeSaiu
        If CLng(m.SubMatches(1)) <> CLng(acumSaida) Then
            If conferir <> "" Then conferir = conferir & "; "
            conferir = conferir & "Observacao diz " & m.SubMatches(1) & " corridas, ultimo acumulado " & acumSaida
        End If
    Else
        conferir = "Observacao nao informa molde/corridas de saida"
    End If
    If DRY_RUN Then Exit Sub
    Set ws = wb.Worksheets("Trocas")
    ult = UltimaLinha(ws, 1)
    For r = 2 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            If CLng(CDate(ws.Cells(r, 1).Value)) = CLng(dt) And UCase$(Txt(ws.Cells(r, 2).Value)) = veio And Txt(ws.Cells(r, 6).Value) = moldeEntrou Then Exit Sub
        End If
    Next r
    ws.Cells(ult + 1, 1).Value = dt
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

Private Sub AplicarVeio(ByVal dt As Date, ByVal veio As String, ByVal molde As Long, ByVal meta As Double, ByVal acum As Double, ByVal obs As String)
    Dim ws As Worksheet, prev As Long, pMolde As String, pAcum As Double, pData As Date, prod As Double
    Set ws = wb.Worksheets("Historico")
    prev = LinhaAnterior(veio, dt)
    If prev > 0 Then
        pMolde = Txt(ws.Cells(prev, 3).Value)
        pAcum = CDbl(ws.Cells(prev, 5).Value)
        pData = CDate(ws.Cells(prev, 1).Value)
        If pMolde <> CStr(molde) Then
            RegistrarTroca dt, veio, pMolde, ws.Cells(prev, 4).Value, pAcum, CStr(molde), obs
        ElseIf acum < pAcum Then
            Alerta "Acumulado diminuiu sem troca", "Veio " & veio, "Dia " & Format$(dt, "dd/mm/yyyy") & ": " & acum & " (antes " & pAcum & " em " & Format$(pData, "dd/mm/yyyy") & ", molde " & molde & ")"
        ElseIf acum = pAcum Then
            prod = ProducaoDoVeio(veio, dt - 1)
            If prod > 0 Then Alerta "Veio nao andou", "Veio " & veio, "Acumulado " & acum & " igual ao de " & Format$(pData, "dd/mm/yyyy") & ", mas a MCC produziu " & prod & " t em " & Format$(dt - 1, "dd/mm/yyyy") & ". Verificar parada ou e-mail repetido."
        End If
    End If
    GravarHistorico dt, veio, molde, meta, acum, obs
End Sub

' ---- Leitura dos anexos -------------------------------------------------------
Private Function ProcessarVidaMolde(ByVal caminho As String, ByRef dataRel As Date) As Boolean
    Dim b As Workbook, ws As Worksheet, r As Long, veio As String, ok As Boolean
    Dim molde As Variant, meta As Variant, acum As Variant, obs As String
    ProcessarVidaMolde = False
    Set b = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True)
    Set ws = AchaAba(b, "VIDA MOLDE")
    If ws Is Nothing Then
        Alerta "Layout alterado", "VIDA MOLDE", "Aba VIDA MOLDE nao encontrada"
        b.Close False
        Exit Function
    End If
    If UCase$(Txt(ws.Cells(3, 2).Value)) <> "VEIOS" Or InStr(UCase$(Txt(ws.Cells(4, 4).Value)), "PROGR") = 0 Or UCase$(Txt(ws.Cells(4, 5).Value)) <> "REAL" Or Not IsDate(ws.Cells(2, 9).Value) Then
        Alerta "Layout alterado", "VIDA MOLDE", "Cabecalhos (B3, D4, E4) ou data (I2) fora do esperado"
        b.Close False
        Exit Function
    End If
    dataRel = CDate(ws.Cells(2, 9).Value)
    ok = True
    For r = 5 To 10
        veio = UCase$(Txt(ws.Cells(r, 2).Value))
        molde = ws.Cells(r, 3).Value
        meta = ws.Cells(r, 4).Value
        acum = ws.Cells(r, 5).Value
        obs = Txt(ws.Cells(r, 8).Value)
        If Len(veio) <> 1 Or InStr("CDEFGH", veio) = 0 Or Txt(molde) = "" Or Txt(acum) = "" Or Txt(meta) = "" Or Not IsNumeric(molde) Or Not IsNumeric(meta) Or Not IsNumeric(acum) Then
            Alerta "Layout alterado", "VIDA MOLDE", "Linha " & r & " invalida (veio=" & veio & ", molde=" & Txt(molde) & ", acumulado=" & Txt(acum) & ")"
            ok = False
        Else
            AplicarVeio dataRel, veio, CLng(molde), CDbl(meta), CDbl(acum), obs
        End If
    Next r
    b.Close False
    ProcessarVidaMolde = ok
End Function

Private Function ProcessarCorteseq(ByVal caminho As String, ByVal nomeArq As String, ByVal refData As Date) As Boolean
    Dim b As Workbook, ws As Worksheet, r As Long, dia As Long, ano As Long, mes As Long
    Dim dt As Date, re As Object, m As Object, real4 As Double, prog4 As Double, real2 As Variant, real3 As Variant
    ProcessarCorteseq = False
    ano = Year(refData)
    mes = Month(refData)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{1,2})-(\d{4})"
    If re.Test(nomeArq) Then
        Set m = re.Execute(nomeArq)(0)
        mes = CLng(m.SubMatches(0))
        ano = CLng(m.SubMatches(1))
    End If
    Set b = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True)
    Set ws = AchaAba(b, "DADOS DA PRODU")
    If ws Is Nothing Then
        Alerta "Layout alterado", "CORTESEQ", "Aba DADOS DA PRODUCAO nao encontrada em " & nomeArq
        b.Close False
        Exit Function
    End If
    If UCase$(Txt(ws.Cells(3, 1).Value)) <> "DIA" Or InStr(UCase$(Txt(ws.Cells(2, 8).Value)), "MCC#4") = 0 Then
        Alerta "Layout alterado", "CORTESEQ", "Cabecalhos (A3, H2) fora do esperado em " & nomeArq
        b.Close False
        Exit Function
    End If
    For r = 4 To 34
        If Txt(ws.Cells(r, 1).Value) <> "" And IsNumeric(ws.Cells(r, 1).Value) Then
            dia = CLng(ws.Cells(r, 1).Value)
            If dia >= 1 And dia <= 31 And Txt(ws.Cells(r, 9).Value) <> "" Then
                dt = DateSerial(ano, mes, dia)
                If Month(dt) = mes Then
                    real4 = CDbl(ws.Cells(r, 9).Value)
                    prog4 = 0
                    If Txt(ws.Cells(r, 8).Value) <> "" And IsNumeric(ws.Cells(r, 8).Value) Then prog4 = CDbl(ws.Cells(r, 8).Value)
                    real2 = ""
                    If Txt(ws.Cells(r, 3).Value) <> "" And IsNumeric(ws.Cells(r, 3).Value) Then real2 = CDbl(ws.Cells(r, 3).Value)
                    real3 = ""
                    If Txt(ws.Cells(r, 6).Value) <> "" And IsNumeric(ws.Cells(r, 6).Value) Then real3 = CDbl(ws.Cells(r, 6).Value)
                    GravarProducao dt, real4, prog4, real2, real3, nomeArq
                End If
            End If
        End If
    Next r
    b.Close False
    ProcessarCorteseq = True
End Function

' ---- Um e-mail --------------------------------------------------------------
Private Function ProcessarEmail(ByVal msg As Object, ByVal tmp As String, ByVal seq As Long) As Boolean
    Dim i As Long, att As Object, nome As String, u As String, ext As String
    Dim pVida As String, pCorte As String, nCorte As String
    Dim temVida As Boolean, temCorte As Boolean, ok As Boolean, dataRel As Date, fso As Object
    ProcessarEmail = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To msg.Attachments.Count
        Set att = msg.Attachments.Item(i)
        nome = att.FileName
        u = UCase$(nome)
        If Right$(u, 5) = ".XLSX" Or Right$(u, 4) = ".XLS" Then
            ext = Mid$(u, InStrRev(u, "."))
            If InStr(u, "VIDA MOLDE") > 0 And Not temVida Then
                pVida = tmp & "\" & seq & "_vida_" & i & ext
                att.SaveAsFile pVida
                temVida = True
            ElseIf InStr(u, "CORTESEQ") > 0 And Not temCorte Then
                pCorte = tmp & "\" & seq & "_corte_" & i & ext
                att.SaveAsFile pCorte
                nCorte = nome
                temCorte = True
            End If
        End If
    Next i
    If Not temVida And Not temCorte Then Exit Function
    ok = True
    dataRel = CDate(msg.ReceivedTime)
    If temVida Then
        If Not ProcessarVidaMolde(pVida, dataRel) Then ok = False
    End If
    If temCorte Then
        If Not ProcessarCorteseq(pCorte, nCorte, dataRel) Then ok = False
    End If
    On Error Resume Next
    If temVida Then fso.DeleteFile pVida, True
    If temCorte Then fso.DeleteFile pCorte, True
    On Error GoTo 0
    ProcessarEmail = ok
End Function

' ---- Principal: Alt+F8 > AtualizarControle > Executar ---------------------------
Public Sub AtualizarControle()
    Dim ol As Object, ns As Object, pasta As Object, itens As Object, msg As Object
    Dim proc As Object, wsP As Worksheet, tmp As String, fso As Object
    Dim seq As Long, nOk As Long, nFalha As Long, nPulados As Long, corte As Date
    Dim filtro As String, id As String, res As Boolean, r As Long, ult As Long

    Set wb = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = Environ$("TEMP") & "\controle_robo"
    If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Registrar IIf(DRY_RUN, "==== Inicio (DRY_RUN: nada sera gravado nos dados)", "==== Inicio")
    Registrar MARCA

    On Error GoTo falha
    Set wsP = AbaOuCria("Processados", "EntryID")
    Set proc = CreateObject("Scripting.Dictionary")
    ult = UltimaLinha(wsP, 1)
    For r = 2 To ult
        proc(Txt(wsP.Cells(r, 1).Value)) = 1
    Next r

    Set ol = CreateObject("Outlook.Application")
    Set ns = ol.GetNamespace("MAPI")
    Set pasta = ns.GetDefaultFolder(6)
    filtro = "@SQL=" & Chr$(34) & "urn:schemas:httpmail:subject" & Chr$(34) & " like '%" & ASSUNTO & "%'"
    Set itens = pasta.Items.Restrict(filtro)
    itens.Sort "[ReceivedTime]", False
    corte = Date - DIAS_RETROATIVOS

    For Each msg In itens
        If msg.Class = 43 Then
            id = msg.EntryID
            If CDate(msg.ReceivedTime) >= corte And Not proc.Exists(id) Then
                seq = seq + 1
                Registrar "E-mail " & Format$(msg.ReceivedTime, "dd/mm/yyyy hh:nn") & " | " & msg.Subject
                res = False
                On Error Resume Next
                res = ProcessarEmail(msg, tmp, seq)
                If Err.Number <> 0 Then
                    Registrar "ERRO no e-mail: " & Err.Description
                    Err.Clear
                    res = False
                End If
                On Error GoTo falha
                If res Then
                    If Not DRY_RUN Then wsP.Cells(UltimaLinha(wsP, 1) + 1, 1).Value = id
                    proc(id) = 1
                    nOk = nOk + 1
                Else
                    nFalha = nFalha + 1
                End If
            Else
                nPulados = nPulados + 1
            End If
        End If
    Next msg

    Registrar "==== Fim: processados=" & nOk & " | com problema ou sem os anexos=" & nFalha & " | ja feitos ou antigos=" & nPulados
    GoTo limpa

falha:
    Registrar "ERRO GERAL: " & Err.Description
limpa:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not DRY_RUN Then wb.Save
    On Error GoTo 0
    MsgBox "Robo terminou." & vbCrLf & "E-mails processados: " & nOk & vbCrLf & "Com problema/sem anexos: " & nFalha & vbCrLf & _
           IIf(DRY_RUN, "MODO TESTE: nada foi gravado nos dados. Veja a aba Log.", "Veja a aba Log e a aba Alertas."), vbInformation, "Controle de equipamentos - " & MARCA
End Sub
