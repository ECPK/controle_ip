Attribute VB_Name = "modApp"
Option Explicit

'==============================================================================
' InfraControl SES
' Módulo.......: modApp
' Descrição....: Informações e validações da aplicação
' Compatível...: Excel 2016+
'==============================================================================

Public Function GetAppName() As String
    GetAppName = APP_NOME
End Function

Public Function GetVersion() As String
    GetVersion = APP_VERSAO
End Function

Public Function GetBuild() As String
    GetBuild = APP_BUILD
End Function

Public Function GetFrameworkVersion() As String
    GetFrameworkVersion = FRAMEWORK_VERSAO
End Function

Public Function GetApplicationTitle() As String
    GetApplicationTitle = APP_NOME & " v" & APP_VERSAO
End Function

Public Function ValidateSheets() As Boolean

    ValidateSheets = _
        WorksheetExists(WS_INICIO) And _
        WorksheetExists(WS_BASE_IPS) And _
        WorksheetExists(WS_DASHBOARD) And _
        WorksheetExists(WS_CONFIG) And _
        WorksheetExists(WS_LOG)

End Function

Public Function ValidateTables() As Boolean

    ValidateTables = _
        TableExists(TB_IPS) And _
        TableExists(TB_CONFIG)

End Function

Public Function ValidateWorkbook() As Boolean

    ValidateWorkbook = ValidateSheets() And ValidateTables()

End Function

Private Function BuildAboutText() As String

    Dim S As String

    S = APP_NOME & vbCrLf
    S = S & "Versão: " & APP_VERSAO & vbCrLf
    S = S & "Build: " & APP_BUILD & vbCrLf
    S = S & "Framework: " & FRAMEWORK_NOME & " " & FRAMEWORK_VERSAO & vbCrLf
    S = S & "Empresa: " & APP_EMPRESA & vbCrLf
    S = S & "Compatibilidade: Excel 2016+" & vbCrLf

    BuildAboutText = S

End Function

Public Sub ShowAbout()

    MsgBox BuildAboutText(), _
           vbInformation + vbOKOnly, _
           APP_NOME

End Sub

Public Function GetWorkbookName() As String
    GetWorkbookName = ThisWorkbook.Name
End Function

Public Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.Path
End Function

Public Function IsWorkbookSaved() As Boolean
    IsWorkbookSaved = Len(ThisWorkbook.Path) > 0
End Function
