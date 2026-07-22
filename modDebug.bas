Attribute VB_Name = "modDebug"
Option Explicit

'==============================================================================
' InfraControl SES - modDebug
' Utilitário para mensagens de depuração.
'==============================================================================

Private Function TimeStamp() As String
    TimeStamp = Format(Now, APP_DATAHORA_FORMATO)
End Function

Private Sub WriteLog(ByVal Nivel As String, ByVal Mensagem As String)

    If Not APP_DEBUG Then Exit Sub

    Debug.Print TimeStamp & " [" & Nivel & "] " & Mensagem

End Sub

Public Sub DebugInfo(ByVal Mensagem As String)
    WriteLog LOG_INFO, Mensagem
End Sub

Public Sub DebugWarning(ByVal Mensagem As String)
    WriteLog LOG_WARNING, Mensagem
End Sub

Public Sub DebugError(ByVal Mensagem As String)
    WriteLog LOG_ERROR, Mensagem
End Sub

Public Sub DebugSuccess(ByVal Mensagem As String)
    WriteLog LOG_SUCCESS, Mensagem
End Sub

Public Sub DebugMessage(ByVal Mensagem As String)
    WriteLog LOG_DEBUG, Mensagem
End Sub

Public Sub DebugSeparator(Optional ByVal Caracter As String = "=")

    If Not APP_DEBUG Then Exit Sub

    Debug.Print String(70, Left$(Caracter, 1))

End Sub

Public Sub DebugVariable(ByVal Nome As String, ByVal Valor As Variant)

    If Not APP_DEBUG Then Exit Sub

    Debug.Print TimeStamp & " [VAR] " & Nome & " = " & CStr(Valor)

End Sub

Public Function DebugStartTimer() As Double
    DebugStartTimer = Timer
End Function

Public Sub DebugEndTimer(ByVal Inicio As Double, Optional ByVal Processo As String = "")

    Dim Tempo As Double

    Tempo = Timer - Inicio

    WriteLog LOG_DEBUG, Processo & " Tempo: " & Format(Tempo, "0.000") & " s"

End Sub

Public Sub DebugObject(ByVal Nome As String, ByVal Objeto As Object)

    If Not APP_DEBUG Then Exit Sub

    If Objeto Is Nothing Then
        WriteLog LOG_WARNING, Nome & " = Nothing"
    Else
        WriteLog LOG_DEBUG, Nome & " carregado."
    End If

End Sub
