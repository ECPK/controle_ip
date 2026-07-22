Attribute VB_Name = "modCore"
Option Explicit

'==============================================================================
' InfraControl SES
' Módulo.......: modCore
' Descrição....: Núcleo da aplicação
' Compatível...: Excel 2016+
'==============================================================================

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mCalculation As XlCalculation
Private mStatusBar As Variant

'------------------------------------------------------------------------------
' Inicializa a aplicação
'------------------------------------------------------------------------------
Public Function InitializeApplication() As Boolean

    On Error GoTo TrataErro

    DebugSeparator
    DebugInfo "Inicializando aplicação..."

    If Not ValidateWorkbook() Then
        DebugError "Estrutura da pasta de trabalho inválida."
        ShowError "A estrutura da pasta de trabalho é inválida."
        InitializeApplication = False
        Exit Function
    End If

    EnterPerformanceMode

    UpdateStatusBar APP_STATUS_INICIALIZANDO

    DebugSuccess "Aplicação inicializada com sucesso."

    InitializeApplication = True

    Exit Function

TrataErro:

    DebugError "Erro durante a inicialização: " & Err.Description

    ExitPerformanceMode

    ShowError Err.Description

    InitializeApplication = False

End Function

'------------------------------------------------------------------------------
' Ativa modo de performance
'------------------------------------------------------------------------------
Public Sub EnterPerformanceMode()

    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mCalculation = Application.Calculation
    mStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

End Sub

'------------------------------------------------------------------------------
' Restaura configurações do Excel
'------------------------------------------------------------------------------
Public Sub ExitPerformanceMode()

    Application.ScreenUpdating = mScreenUpdating
    Application.EnableEvents = mEnableEvents
    Application.DisplayAlerts = mDisplayAlerts
    Application.Calculation = mCalculation
    Application.StatusBar = mStatusBar

End Sub

'------------------------------------------------------------------------------
' Atualiza barra de status
'------------------------------------------------------------------------------
Public Sub UpdateStatusBar(ByVal Texto As String)

    Application.StatusBar = APP_NOME & " - " & Texto

End Sub

'------------------------------------------------------------------------------
' Finaliza aplicação
'------------------------------------------------------------------------------
Public Sub ShutdownApplication()

    DebugInfo "Finalizando aplicação..."

    UpdateStatusBar APP_STATUS_FINALIZADO

    ExitPerformanceMode

    Application.StatusBar = False

    DebugSuccess "Aplicação finalizada."

End Sub

'------------------------------------------------------------------------------
' Reinicia a aplicação
'------------------------------------------------------------------------------
Public Function RestartApplication() As Boolean

    ShutdownApplication

    RestartApplication = InitializeApplication()

End Function
