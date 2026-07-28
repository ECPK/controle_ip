VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServidorRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: clsServidorRepository
' Projeto....: InfraControl SES
' Build......: 2026.07.22.001
' Descrição..: Repositório especializado da entidade Servidor
' Compatível.: Excel 2016+
'==============================================================================

' Esta classe encapsula toda a persistência da entidade clsServidor.
' Ela utiliza exclusivamente o modRepository/modExcel para acesso às tabelas.

Public Function Exists(ByVal ID As String) As Boolean

    On Error GoTo TratarErro

    Exists = RecordExists(TB_SERVIDORES, COL_SERV_ID, ID)

    Exit Function

TratarErro:
    Exists = False
End Function


Public Function Save(ByVal Servidor As clsServidor) As Boolean

    On Error GoTo TratarErro

    Dim Mapper As New clsDataMapper

    If Servidor Is Nothing Then Exit Function
    If Len(Servidor.ID) = 0 Then Servidor.ID = Servidor.Hostname
    If Not Servidor.Validate Then Exit Function

    Servidor.Touch
    If Not Mapper.SaveServidor(Servidor) Then Exit Function

    Servidor.AcceptChanges
    Save = True

    Exit Function

TratarErro:

    DebugError "Erro ao salvar servidor: " & Err.Description
    Save = False

End Function


Public Function Delete(ByVal ID As String) As Boolean

    On Error GoTo TratarErro

    Delete = DeleteRecordByKey(TB_SERVIDORES, COL_SERV_ID, ID)

    Exit Function

TratarErro:

    DebugError "Erro ao excluir servidor: " & Err.Description
    Delete = False

End Function


Public Function Load(ByVal ID As String) As clsServidor

    On Error GoTo TratarErro

    Dim Mapper As New clsDataMapper

    If Not Exists(ID) Then Exit Function

    Set Load = Mapper.LoadServidor(ID)

    Exit Function

TratarErro:

    DebugError "Erro ao carregar servidor: " & Err.Description
    Set Load = Nothing

End Function
