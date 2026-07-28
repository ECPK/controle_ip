Attribute VB_Name = "modImportacao"
Option Explicit

'==============================================================================
' InfraControl SES
' Modulo.......: modImportacao
' Descricao....: Importacao LEVANTAMENTO.xlsx -> Inventario -> Base_IPs
' Compativel...: Excel 2016+
'==============================================================================

Private Const IMPORT_HEADER_ROW As Long = 1

Public Sub ExecutarImportacaoLevantamento()

    ImportarLevantamento

End Sub
Public Function ImportarLevantamento(Optional ByVal FilePath As String = "") As Boolean

    On Error GoTo TrataErro

    Dim SelectedFile As Variant
    Dim SourceWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim RowsImported As Long
    Dim RowsProcessed As Long

    If Len(FilePath) = 0 Then
        SelectedFile = Application.GetOpenFilename( _
            FileFilter:="Arquivos Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
            Title:="Selecionar LEVANTAMENTO.xlsx")

        If VarType(SelectedFile) = vbBoolean Then
            RegistrarImportacao LOG_WARNING, MSG_IMPORTACAO_CANCELADA
            Exit Function
        End If

        FilePath = CStr(SelectedFile)
    End If

    DebugInfo "Importando levantamento: " & FilePath
    UpdateStatusBar "Importando levantamento"

    LimparInventario

    Set SourceWorkbook = Workbooks.Open(FileName:=FilePath, ReadOnly:=True)
    Set SourceWorksheet = GetSourceWorksheet(SourceWorkbook)

    RowsImported = CopiarParaInventario(SourceWorksheet)

    SourceWorkbook.Close SaveChanges:=False
    Set SourceWorkbook = Nothing

    RowsProcessed = ProcessarInventarioParaBase()
    LimparInventario

    RegistrarImportacao LOG_SUCCESS, MSG_IMPORTACAO_SUCESSO & _
        " Linhas importadas: " & RowsImported & _
        ". Linhas processadas: " & RowsProcessed & "."

    UpdateStatusBar APP_STATUS_PRONTO
    ImportarLevantamento = True

    Exit Function

TrataErro:
    On Error Resume Next
    If Not SourceWorkbook Is Nothing Then SourceWorkbook.Close SaveChanges:=False
    RegistrarImportacao LOG_ERROR, "Erro na importacao: " & Err.Description
    DebugError "modImportacao.ImportarLevantamento: " & Err.Description
    UpdateStatusBar APP_STATUS_ERRO
    ImportarLevantamento = False

End Function

Public Function ProcessarInventarioParaBase() As Long

    On Error GoTo TrataErro

    Dim wsInv As Worksheet
    Dim tblBase As ListObject
    Dim SourceRange As Range
    Dim RowIndex As Long
    Dim NewRow As ListRow

    Set wsInv = GetWorksheet(WS_INVENTARIO)
    Set tblBase = GetTable(TB_IPS)

    If wsInv Is Nothing Then Err.Raise vbObjectError + 301, , "Aba Inventario nao encontrada."
    If tblBase Is Nothing Then Err.Raise vbObjectError + 302, , "Tabela tbIPs nao encontrada."
    If WorksheetFunction.CountA(wsInv.Cells) = 0 Then Exit Function

    Set SourceRange = wsInv.UsedRange
    If SourceRange.Rows.Count <= IMPORT_HEADER_ROW Then Exit Function

    ClearTableData tblBase

    For RowIndex = IMPORT_HEADER_ROW + 1 To SourceRange.Rows.Count
        If WorksheetFunction.CountA(SourceRange.Rows(RowIndex)) > 0 Then
            Set NewRow = tblBase.ListRows.Add
            CopiarLinhaInventarioParaBase SourceRange, RowIndex, tblBase, NewRow
            ProcessarInventarioParaBase = ProcessarInventarioParaBase + 1
        End If
    Next RowIndex

    RegistrarImportacao LOG_INFO, "Base_IPs atualizada a partir do Inventario. Registros: " & ProcessarInventarioParaBase

    Exit Function

TrataErro:
    DebugError "modImportacao.ProcessarInventarioParaBase: " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub LimparInventario()

    Dim wsInv As Worksheet

    Set wsInv = GetWorksheet(WS_INVENTARIO)
    If wsInv Is Nothing Then Exit Sub

    wsInv.Cells.Clear

End Sub

Private Function GetSourceWorksheet(ByVal SourceWorkbook As Workbook) As Worksheet

    On Error Resume Next
    Set GetSourceWorksheet = SourceWorkbook.Worksheets(WS_INVENTARIO)
    On Error GoTo 0

    If GetSourceWorksheet Is Nothing Then
        Set GetSourceWorksheet = SourceWorkbook.Worksheets(1)
    End If

End Function

Private Function CopiarParaInventario(ByVal SourceWorksheet As Worksheet) As Long

    Dim wsInv As Worksheet
    Dim SourceRange As Range

    Set wsInv = GetWorksheet(WS_INVENTARIO)
    If wsInv Is Nothing Then Err.Raise vbObjectError + 303, , "Aba Inventario nao encontrada."

    Set SourceRange = SourceWorksheet.UsedRange
    If WorksheetFunction.CountA(SourceRange) = 0 Then Exit Function

    wsInv.Range("A1").Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).Value = SourceRange.Value

    If SourceRange.Rows.Count > IMPORT_HEADER_ROW Then
        CopiarParaInventario = SourceRange.Rows.Count - IMPORT_HEADER_ROW
    End If

End Function

Private Sub CopiarLinhaInventarioParaBase(ByVal SourceRange As Range, _
                                          ByVal SourceRowIndex As Long, _
                                          ByVal tblBase As ListObject, _
                                          ByVal TargetRow As ListRow)

    Dim BaseColumn As ListColumn
    Dim SourceColumnIndex As Long

    For Each BaseColumn In tblBase.ListColumns
        SourceColumnIndex = FindSourceColumnIndex(SourceRange, BaseColumn.Name)
        If SourceColumnIndex > 0 Then
            TargetRow.Range.Cells(1, BaseColumn.Index).Value = _
                SourceRange.Cells(SourceRowIndex, SourceColumnIndex).Value
        Else
            TargetRow.Range.Cells(1, BaseColumn.Index).Value = GetDefaultBaseValue(BaseColumn.Name, SourceRange, SourceRowIndex)
        End If
    Next BaseColumn

End Sub

Private Function FindSourceColumnIndex(ByVal SourceRange As Range, ByVal BaseHeaderName As String) As Long

    FindSourceColumnIndex = FindHeaderIndex(SourceRange, BaseHeaderName)
    If FindSourceColumnIndex > 0 Then Exit Function

    Select Case BaseHeaderName
        Case COL_SERV_HOSTNAME
            FindSourceColumnIndex = FindHeaderIndex(SourceRange, "Nome do Recurso")
        Case COL_SERV_INTERFACE
            FindSourceColumnIndex = FindHeaderIndex(SourceRange, "Nome (interface)")
        Case COL_SERV_SO
            FindSourceColumnIndex = FindHeaderIndex(SourceRange, "SO")
        Case COL_SERV_MEMORIA
            FindSourceColumnIndex = FindHeaderIndex(SourceRange, "RAM ATUAL")
        Case COL_SERV_RAM_RECOMENDADA
            FindSourceColumnIndex = FindHeaderIndex(SourceRange, "RAM Recomendada (MB)")
    End Select

End Function

Private Function FindHeaderIndex(ByVal SourceRange As Range, ByVal HeaderName As String) As Long

    Dim ColumnIndex As Long

    For ColumnIndex = 1 To SourceRange.Columns.Count
        If StrComp(CStr(SourceRange.Cells(IMPORT_HEADER_ROW, ColumnIndex).Value), HeaderName, vbTextCompare) = 0 Then
            FindHeaderIndex = ColumnIndex
            Exit Function
        End If
    Next ColumnIndex

End Function

Private Function GetDefaultBaseValue(ByVal BaseHeaderName As String, _
                                     ByVal SourceRange As Range, _
                                     ByVal SourceRowIndex As Long) As Variant

    Select Case BaseHeaderName
        Case COL_SERV_ID
            GetDefaultBaseValue = BuildImportedID(SourceRange, SourceRowIndex)
        Case COL_SERV_DATA_CADASTRO, COL_SERV_DATA_ATUALIZACAO
            GetDefaultBaseValue = Now
    End Select

End Function

Private Function BuildImportedID(ByVal SourceRange As Range, ByVal SourceRowIndex As Long) As String

    Dim SourceColumnIndex As Long
    Dim Value As String

    SourceColumnIndex = FindHeaderIndex(SourceRange, COL_SERV_IP)
    If SourceColumnIndex > 0 Then Value = Trim$(CStr(SourceRange.Cells(SourceRowIndex, SourceColumnIndex).Value))

    If Len(Value) = 0 Then
        SourceColumnIndex = FindSourceColumnIndex(SourceRange, COL_SERV_HOSTNAME)
        If SourceColumnIndex > 0 Then Value = Trim$(CStr(SourceRange.Cells(SourceRowIndex, SourceColumnIndex).Value))
    End If

    BuildImportedID = Value

End Function

Private Sub RegistrarImportacao(ByVal Nivel As String, ByVal Mensagem As String)

    Dim wsLog As Worksheet
    Dim NextRow As Long

    DebugMessage Nivel & " - " & Mensagem

    Set wsLog = GetWorksheet(WS_LOG)
    If wsLog Is Nothing Then Exit Sub

    NextRow = LastRow(wsLog, 1)
    If Len(CStr(wsLog.Cells(NextRow, 1).Value)) > 0 Then NextRow = NextRow + 1

    wsLog.Cells(NextRow, 1).Value = Now
    wsLog.Cells(NextRow, 2).Value = Nivel
    wsLog.Cells(NextRow, 3).Value = ACTION_IMPORT
    wsLog.Cells(NextRow, 4).Value = Mensagem

End Sub
