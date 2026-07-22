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

    If Servidor Is Nothing Then Exit Function

    If Not Servidor.Validate Then Exit Function

    ' Nesta primeira versão a persistência ainda será integrada
    ' ao modRepository nas próximas entregas.

    Servidor.AcceptChanges

    Save = True

    Exit Function

TratarErro:

    DebugError "Erro ao salvar servidor: " & Err.Description
    Save = False

End Function


Public Function Delete(ByVal ID As String) As Boolean

    On Error GoTo TratarErro

    Delete = DeleteRecord(TB_SERVIDORES, COL_SERV_ID, ID)

    Exit Function

TratarErro:

    DebugError "Erro ao excluir servidor: " & Err.Description
    Delete = False

End Function


Public Function Load(ByVal ID As String) As clsServidor

    On Error GoTo TratarErro

    Dim Servidor As New clsServidor

    ' Próxima etapa:
    ' Ler os campos utilizando ReadField()
    ' Popular o objeto
    ' Chamar AcceptChanges()

    Set Load = Servidor

    Exit Function

TratarErro:

    DebugError "Erro ao carregar servidor: " & Err.Description
    Set Load = Nothing

End Function
