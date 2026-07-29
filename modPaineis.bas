Attribute VB_Name = "modPaineis"
Option Explicit

'==============================================================================
' InfraControl SES
' Modulo.......: modPaineis
' Descricao....: Gera paineis visuais de IP a partir da Base_IPs
' Compativel...: Excel 2016+
'==============================================================================

Private Const PANEL_FIRST_GRID_ROW As Long = 10
Private Const PANEL_GRID_COLUMNS As Long = 32
Private Const PANEL_GRID_ROWS As Long = 8

Public Sub AtualizarPaineis()

    On Error GoTo TrataErro

    Dim redes As Collection
    Dim rede As Variant

    Set redes = ListarRedesPainel()

    For Each rede In redes
        AtualizarPainelRede CStr(rede)
    Next rede

    AtualizarDashboard
    DebugSuccess "Paineis atualizados a partir da Base_IPs."

    Exit Sub

TrataErro:
    DebugError "modPaineis.AtualizarPaineis: " & Err.Description
    ShowError "Erro ao atualizar paineis: " & Err.Description

End Sub

Public Sub AtualizarPainelRede(ByVal Rede As String)

    On Error GoTo TrataErro

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim octeto As Long
    Dim r As Long
    Dim c As Long
    Dim rw As ListRow
    Dim ipInfo As Object

    Set tbl = GetTable(TB_IPS)
    If tbl Is Nothing Then Err.Raise vbObjectError + 401, , "Tabela tbIPs nao encontrada."

    Set ws = EnsurePainelWorksheet(Rede)
    PrepararPainel ws, Rede

    For octeto = 0 To 255
        r = PANEL_FIRST_GRID_ROW + (octeto \ PANEL_GRID_COLUMNS)
        c = 1 + (octeto Mod PANEL_GRID_COLUMNS)

        With ws.Cells(r, c)
            .Value = octeto
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 255)
            .Font.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(120, 120, 120)
            .ClearComments
        End With
    Next octeto

    If Not tbl.DataBodyRange Is Nothing Then
        For Each rw In tbl.ListRows
            If StrComp(GetRowRede(rw), Rede, vbTextCompare) = 0 Then
                Set ipInfo = BuildIPInfo(rw)
                octeto = CLng(ipInfo("Octeto"))
                If octeto >= 0 And octeto <= 255 Then
                    AplicarIPNoPainel ws.Cells(PANEL_FIRST_GRID_ROW + (octeto \ PANEL_GRID_COLUMNS), _
                                               1 + (octeto Mod PANEL_GRID_COLUMNS)), ipInfo
                End If
            End If
        Next rw
    End If

    ws.Columns("A:AF").ColumnWidth = 3.5
    ws.Rows(PANEL_FIRST_GRID_ROW & ":" & PANEL_FIRST_GRID_ROW + PANEL_GRID_ROWS - 1).RowHeight = 18

    Exit Sub

TrataErro:
    DebugError "modPaineis.AtualizarPainelRede: " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub PrepararPainel(ByVal ws As Worksheet, ByVal Rede As String)

    ws.Cells.Clear
    ws.Cells.ClearComments

    With ws.Range("A1:AF1")
        .Merge
        .Value = "Relacao de IPs - " & Rede & ".XX"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
    End With

    EscreverLegenda ws

End Sub

Private Sub EscreverLegenda(ByVal ws As Worksheet)

    Dim startCol As Long

    startCol = 35

    ws.Cells(1, startCol).Value = "Legenda"
    ws.Cells(1, startCol).Font.Bold = True

    EscreverItemLegenda ws, 2, startCol, "POC", CorStatus("POC")
    EscreverItemLegenda ws, 3, startCol, "Utilizado", CorStatus("Utilizado")
    EscreverItemLegenda ws, 4, startCol, "Pre reservado", CorStatus("Pre reservado")
    EscreverItemLegenda ws, 5, startCol, "Disponivel", CorStatus("Disponivel")
    EscreverItemLegenda ws, 6, startCol, "Temporario", CorStatus("Temporario")

    ws.Columns(startCol).ColumnWidth = 4
    ws.Columns(startCol + 1).ColumnWidth = 18

End Sub

Private Sub EscreverItemLegenda(ByVal ws As Worksheet, _
                                ByVal linha As Long, _
                                ByVal coluna As Long, _
                                ByVal texto As String, _
                                ByVal cor As Long)

    ws.Cells(linha, coluna).Interior.Color = cor
    ws.Cells(linha, coluna).Borders.LineStyle = xlContinuous
    ws.Cells(linha, coluna + 1).Value = texto
    ws.Cells(linha, coluna + 1).Font.Bold = True

End Sub

Private Sub AplicarIPNoPainel(ByVal Target As Range, ByVal ipInfo As Object)

    Dim status As String
    Dim comentario As String

    status = CStr(ipInfo("Status"))

    With Target
        .Interior.Color = CorStatus(status)
        .Font.Color = CorFonteStatus(status)
        .Font.Bold = True
        .ClearComments

        comentario = CStr(ipInfo("Comentario"))
        If Len(comentario) > 0 Then .AddComment comentario
    End With

End Sub

Private Function BuildIPInfo(ByVal rw As ListRow) As Object

    Dim data As Object
    Dim ip As String
    Dim status As String

    Set data = CreateObject("Scripting.Dictionary")

    ip = CStr(ReadField(rw, COL_SERV_IP))
    status = Trim$(CStr(ReadField(rw, COL_SERV_STATUS_IP)))
    If Len(status) = 0 Then status = "Utilizado"

    data.Add "IP", ip
    data.Add "Octeto", ExtrairOcteto(ip)
    data.Add "Status", status
    data.Add "Comentario", BuildComentarioIP(rw)

    Set BuildIPInfo = data

End Function

Private Function BuildComentarioIP(ByVal rw As ListRow) As String

    Dim texto As String

    texto = AddComentarioLinha(texto, "IP", CStr(ReadField(rw, COL_SERV_IP)))
    texto = AddComentarioLinha(texto, "Hostname", CStr(ReadField(rw, COL_SERV_HOSTNAME)))
    texto = AddComentarioLinha(texto, "AD", CStr(ReadField(rw, COL_SERV_AD)))
    texto = AddComentarioLinha(texto, "Interface", CStr(ReadField(rw, COL_SERV_INTERFACE)))
    texto = AddComentarioLinha(texto, "Gerencia", CStr(ReadField(rw, COL_SERV_GERENCIA)))
    texto = AddComentarioLinha(texto, "Sistema", CStr(ReadField(rw, COL_SERV_SISTEMA)))
    texto = AddComentarioLinha(texto, "Projeto", CStr(ReadField(rw, COL_SERV_PROJETO)))
    texto = AddComentarioLinha(texto, "Observacoes", CStr(ReadField(rw, COL_SERV_OBSERVACOES)))

    BuildComentarioIP = texto

End Function

Private Function AddComentarioLinha(ByVal texto As String, _
                                    ByVal rotulo As String, _
                                    ByVal valor As String) As String

    valor = Trim$(valor)
    If Len(valor) = 0 Then
        AddComentarioLinha = texto
        Exit Function
    End If

    If Len(texto) > 0 Then texto = texto & vbCrLf
    AddComentarioLinha = texto & rotulo & ": " & valor

End Function

Private Function ListarRedesPainel() As Collection

    Dim redes As New Collection
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim rede As String

    AddRedeUnica redes, "10.233.66"
    AddRedeUnica redes, "10.233.87"
    AddRedeUnica redes, "10.233.166"
    AddRedeUnica redes, "10.233.187"

    Set tbl = GetTable(TB_IPS)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            For Each rw In tbl.ListRows
                rede = GetRowRede(rw)
                If Len(rede) > 0 Then AddRedeUnica redes, rede
            Next rw
        End If
    End If

    Set ListarRedesPainel = redes

End Function

Private Sub AddRedeUnica(ByVal redes As Collection, ByVal rede As String)

    On Error Resume Next
    redes.Add rede, rede
    On Error GoTo 0

End Sub

Private Function GetRowRede(ByVal rw As ListRow) As String

    Dim rede As String
    Dim ip As String

    rede = Trim$(CStr(ReadField(rw, COL_SERV_REDE)))
    If Len(rede) = 0 Then
        ip = Trim$(CStr(ReadField(rw, COL_SERV_IP)))
        rede = ExtrairRede(ip)
    End If

    GetRowRede = rede

End Function

Private Function EnsurePainelWorksheet(ByVal Rede As String) As Worksheet

    Dim sheetName As String

    sheetName = "Painel_" & Replace(Rede, "10.233.", "")

    Set EnsurePainelWorksheet = GetWorksheet(sheetName)
    If EnsurePainelWorksheet Is Nothing Then
        Set EnsurePainelWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsurePainelWorksheet.Name = sheetName
    End If

End Function

Private Function CorStatus(ByVal status As String) As Long

    Select Case NormalizeStatus(status)
        Case "POC"
            CorStatus = RGB(0, 176, 80)
        Case "UTILIZADO"
            CorStatus = RGB(255, 0, 0)
        Case "PRE RESERVADO", "RESERVADO"
            CorStatus = RGB(255, 255, 0)
        Case "TEMPORARIO"
            CorStatus = RGB(91, 155, 213)
        Case Else
            CorStatus = RGB(255, 255, 255)
    End Select

End Function

Private Function CorFonteStatus(ByVal status As String) As Long

    Select Case NormalizeStatus(status)
        Case "UTILIZADO"
            CorFonteStatus = RGB(255, 255, 255)
        Case Else
            CorFonteStatus = RGB(0, 0, 0)
    End Select

End Function

Private Function NormalizeStatus(ByVal status As String) As String

    status = UCase$(Trim$(status))

    If status = "POC" Then
        NormalizeStatus = "POC"
    ElseIf InStr(1, status, "UTILIZ", vbTextCompare) > 0 Then
        NormalizeStatus = "UTILIZADO"
    ElseIf InStr(1, status, "RESERVADO", vbTextCompare) > 0 Then
        NormalizeStatus = "PRE RESERVADO"
    ElseIf InStr(1, status, "TEMPOR", vbTextCompare) > 0 Then
        NormalizeStatus = "TEMPORARIO"
    ElseIf InStr(1, status, "DISPON", vbTextCompare) > 0 Then
        NormalizeStatus = "DISPONIVEL"
    Else
        NormalizeStatus = status
    End If

End Function
Private Function ExtrairRede(ByVal ip As String) As String

    Dim partes() As String

    partes = Split(Trim$(ip), ".")
    If UBound(partes) < 3 Then Exit Function

    ExtrairRede = partes(0) & "." & partes(1) & "." & partes(2)

End Function

Private Function ExtrairOcteto(ByVal ip As String) As Long

    Dim partes() As String

    partes = Split(Trim$(ip), ".")
    If UBound(partes) < 3 Then
        ExtrairOcteto = -1
    Else
        ExtrairOcteto = CLng(Val(partes(3)))
    End If

End Function

Private Sub AtualizarDashboard()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim total As Long
    Dim utilizados As Long
    Dim disponiveis As Long

    Set ws = GetWorksheet(WS_DASHBOARD)
    Set tbl = GetTable(TB_IPS)
    If ws Is Nothing Or tbl Is Nothing Then Exit Sub

    If Not tbl.DataBodyRange Is Nothing Then total = tbl.DataBodyRange.Rows.Count
    utilizados = CountByStatus("Utilizado")
    disponiveis = (ListarRedesPainel.Count * 256) - utilizados

    ws.Cells.Clear
    ws.Range("A1").Value = "Dashboard"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Registros na Base_IPs"
    ws.Range("B3").Value = total
    ws.Range("A4").Value = "IPs utilizados"
    ws.Range("B4").Value = utilizados
    ws.Range("A5").Value = "IPs disponiveis estimados"
    ws.Range("B5").Value = disponiveis
    ws.Range("A7").Value = "Atualizado em"
    ws.Range("B7").Value = Now

End Sub

Private Function CountByStatus(ByVal statusName As String) As Long

    Dim tbl As ListObject
    Dim rw As ListRow
    Dim status As String

    Set tbl = GetTable(TB_IPS)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    For Each rw In tbl.ListRows
        status = Trim$(CStr(ReadField(rw, COL_SERV_STATUS_IP)))
        If Len(status) = 0 Then status = "Utilizado"
        If NormalizeStatus(status) = NormalizeStatus(statusName) Then CountByStatus = CountByStatus + 1
    Next rw

End Function
