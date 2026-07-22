Attribute VB_Name = "modExcel"
Option Explicit

'==============================================================================
' InfraControl SES - modExcel
' Funções utilitárias para acesso ao Excel.
'==============================================================================

Public Function GetWorksheet(ByVal SheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, SheetName, vbTextCompare) = 0 Then
            Set GetWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Public Function WorksheetExists(ByVal SheetName As String) As Boolean
    WorksheetExists = Not GetWorksheet(SheetName) Is Nothing
End Function

Public Function GetTable(ByVal TableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, TableName, vbTextCompare) = 0 Then
                Set GetTable = tbl
                Exit Function
            End If
        Next tbl
    Next ws
End Function

Public Function TableExists(ByVal TableName As String) As Boolean
    TableExists = Not GetTable(TableName) Is Nothing
End Function

Public Function LastRow(ByVal ws As Worksheet, Optional ByVal Col As Variant = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, Col).End(xlUp).Row
End Function

Public Function LastColumn(ByVal ws As Worksheet, Optional ByVal RowNumber As Long = 1) As Long
    LastColumn = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function GetTableColumn(ByVal tbl As ListObject, ByVal ColumnName As String) As ListColumn
    Dim lc As ListColumn
    For Each lc In tbl.ListColumns
        If StrComp(lc.Name, ColumnName, vbTextCompare) = 0 Then
            Set GetTableColumn = lc
            Exit Function
        End If
    Next lc
End Function

Public Function TableColumnExists(ByVal tbl As ListObject, ByVal ColumnName As String) As Boolean
    TableColumnExists = Not GetTableColumn(tbl, ColumnName) Is Nothing
End Function

Public Sub ClearTableFilter(ByVal tbl As ListObject)
    On Error Resume Next
    If Not tbl Is Nothing Then
        If tbl.ShowAutoFilter Then
            tbl.AutoFilter.ShowAllData
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub ClearTableData(ByVal tbl As ListObject)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    tbl.DataBodyRange.Delete
End Sub

Public Function ReadCell(ByVal ws As Worksheet, ByVal Row As Long, ByVal Col As Variant) As Variant
    ReadCell = ws.Cells(Row, Col).Value
End Function

Public Sub WriteCell(ByVal ws As Worksheet, ByVal Row As Long, ByVal Col As Variant, ByVal Value As Variant)
    ws.Cells(Row, Col).Value = Value
End Sub

Public Function FindColumn(ByVal ws As Worksheet, ByVal Header As String, Optional HeaderRow As Long = 1) As Long
    Dim c As Range
    For Each c In ws.Rows(HeaderRow).Cells
        If Len(c.Value) = 0 Then Exit For
        If StrComp(CStr(c.Value), Header, vbTextCompare) = 0 Then
            FindColumn = c.Column
            Exit Function
        End If
    Next c
End Function

Public Sub ShowError(ByVal Msg As String, Optional ByVal Title As String = APP_NOME)
    MsgBox Msg, vbCritical + vbOKOnly, Title
End Sub
