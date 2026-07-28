Attribute VB_Name = "modRepository"
Option Explicit

'==============================================================================
' InfraControl SES
' Módulo.......: modRepository
' Descrição....: Camada de acesso aos dados (ListObjects)
' Compatível...: Excel 2016+
'==============================================================================

Public Function GetRecordCount(ByVal TableName As String) As Long

    Dim tbl As ListObject

    Set tbl = GetTable(TableName)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    GetRecordCount = tbl.DataBodyRange.Rows.Count

End Function

Public Function FindRow(ByVal TableName As String, _
                        ByVal ColumnName As String, _
                        ByVal SearchValue As Variant) As ListRow

    Dim tbl As ListObject
    Dim lc As ListColumn
    Dim c As Range

    Set tbl = GetTable(TableName)
    If tbl Is Nothing Then Exit Function

    Set lc = GetTableColumn(tbl, ColumnName)
    If lc Is Nothing Then Exit Function

    If lc.DataBodyRange Is Nothing Then Exit Function

    For Each c In lc.DataBodyRange.Cells
        If CStr(c.Value) = CStr(SearchValue) Then
            Set FindRow = tbl.ListRows(c.Row - tbl.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next c

End Function

Public Function RecordExists(ByVal TableName As String, _
                             ByVal ColumnName As String, _
                             ByVal SearchValue As Variant) As Boolean

    RecordExists = Not FindRow(TableName, ColumnName, SearchValue) Is Nothing

End Function

Public Function ReadField(ByVal Row As ListRow, _
                          ByVal ColumnName As String) As Variant

    Dim idx As Long

    idx = Row.Parent.ListColumns(ColumnName).Index
    ReadField = Row.Range.Cells(1, idx).Value

End Function

Public Sub WriteField(ByVal Row As ListRow, _
                      ByVal ColumnName As String, _
                      ByVal Value As Variant)

    Dim idx As Long

    idx = Row.Parent.ListColumns(ColumnName).Index
    Row.Range.Cells(1, idx).Value = Value

End Sub

Public Function InsertRecord(ByVal TableName As String) As ListRow

    Dim tbl As ListObject

    Set tbl = GetTable(TableName)
    If tbl Is Nothing Then Exit Function

    Set InsertRecord = tbl.ListRows.Add

End Function

Public Function DeleteRecord(ByVal Row As ListRow) As Boolean

    On Error GoTo TrataErro

    Row.Delete

    DeleteRecord = True
    Exit Function

TrataErro:
    DeleteRecord = False

End Function

Public Function FindValue(ByVal TableName As String, _
                          ByVal SearchColumn As String, _
                          ByVal SearchValue As Variant, _
                          ByVal ReturnColumn As String) As Variant

    Dim rw As ListRow

    Set rw = FindRow(TableName, SearchColumn, SearchValue)

    If rw Is Nothing Then Exit Function

    FindValue = ReadField(rw, ReturnColumn)

End Function

Public Function EnsureRecord(ByVal TableName As String, _
                             ByVal KeyColumn As String, _
                             ByVal KeyValue As Variant) As ListRow

    Dim rw As ListRow

    Set rw = FindRow(TableName, KeyColumn, KeyValue)

    If rw Is Nothing Then
        Set rw = InsertRecord(TableName)
        If Not rw Is Nothing Then WriteField rw, KeyColumn, KeyValue
    End If

    Set EnsureRecord = rw

End Function

Public Function ReadFieldByKey(ByVal TableName As String, _
                               ByVal SearchColumn As String, _
                               ByVal SearchValue As Variant, _
                               ByVal ReturnColumn As String) As Variant

    Dim tbl As ListObject
    Dim rw As ListRow

    Set tbl = GetTable(TableName)
    If tbl Is Nothing Then Exit Function
    If GetTableColumn(tbl, ReturnColumn) Is Nothing Then Exit Function

    Set rw = FindRow(TableName, SearchColumn, SearchValue)
    If rw Is Nothing Then Exit Function

    ReadFieldByKey = ReadField(rw, ReturnColumn)

End Function

Public Function WriteFieldByKey(ByVal TableName As String, _
                                ByVal SearchColumn As String, _
                                ByVal SearchValue As Variant, _
                                ByVal UpdateColumn As String, _
                                ByVal NewValue As Variant) As Boolean

    Dim tbl As ListObject
    Dim rw As ListRow

    Set tbl = GetTable(TableName)
    If tbl Is Nothing Then Exit Function
    If GetTableColumn(tbl, UpdateColumn) Is Nothing Then Exit Function

    Set rw = EnsureRecord(TableName, SearchColumn, SearchValue)
    If rw Is Nothing Then Exit Function

    WriteField rw, UpdateColumn, NewValue

    WriteFieldByKey = True

End Function

Public Function DeleteRecordByKey(ByVal TableName As String, _
                                  ByVal SearchColumn As String, _
                                  ByVal SearchValue As Variant) As Boolean

    Dim rw As ListRow

    Set rw = FindRow(TableName, SearchColumn, SearchValue)
    If rw Is Nothing Then Exit Function

    DeleteRecordByKey = DeleteRecord(rw)

End Function

Public Function UpdateRecord(ByVal TableName As String, _
                             ByVal SearchColumn As String, _
                             ByVal SearchValue As Variant, _
                             ByVal UpdateColumn As String, _
                             ByVal NewValue As Variant) As Boolean

    Dim rw As ListRow

    Set rw = FindRow(TableName, SearchColumn, SearchValue)

    If rw Is Nothing Then Exit Function

    WriteField rw, UpdateColumn, NewValue

    UpdateRecord = True

End Function
