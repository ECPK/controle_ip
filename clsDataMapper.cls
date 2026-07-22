VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: clsDataMapper
' Projeto....: InfraControl SES
' Build......: 2026.07.22.001
' Descrição..: Mapeamento entre entidades e tabelas
' Compatível.: Excel 2016+
'==============================================================================

Private mTableName As String

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal Value As String)
    mTableName = Trim$(Value)
End Property

'------------------------------------------------------------------------------
' Salva uma entidade em uma linha da tabela.
' Nesta primeira versão o mapeamento é explícito e poderá evoluir para
' metadados/atributos centralizados nas próximas builds.
'------------------------------------------------------------------------------
Public Function SaveServidor(ByVal Servidor As clsServidor) As Boolean

    On Error GoTo TratarErro:

    If Servidor Is Nothing Then Exit Function
    If Not Servidor.Validate Then Exit Function

    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_HOSTNAME, Servidor.Hostname
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_FQDN, Servidor.FQDN
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_DESCRICAO, Servidor.Descricao
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_AMBIENTE, Servidor.Ambiente
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_SO, Servidor.SistemaOperacional
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_VERSAO_SO, Servidor.VersaoSO
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_ARQUITETURA, Servidor.Arquitetura
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_CPU, Servidor.CPU
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_MEMORIA, Servidor.MemoriaGB
    WriteField TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_DISCO, Servidor.DiscoGB

    SaveServidor = True
    Exit Function

TratarErro:
    DebugError "clsDataMapper.SaveServidor: " & Err.Description
    SaveServidor = False
End Function

'------------------------------------------------------------------------------
' Carrega uma entidade a partir da tabela.
'------------------------------------------------------------------------------
Public Function LoadServidor(ByVal ID As String) As clsServidor

    On Error GoTo TratarErro:

    Dim S As New clsServidor

    S.ID = ID
    S.Hostname = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_HOSTNAME))
    S.FQDN = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_FQDN))
    S.Descricao = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DESCRICAO))
    S.Ambiente = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_AMBIENTE))
    S.SistemaOperacional = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_SO))
    S.VersaoSO = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_VERSAO_SO))
    S.Arquitetura = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_ARQUITETURA))
    S.CPU = CStr(ReadField(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_CPU))

    S.AcceptChanges

    Set LoadServidor = S
    Exit Function

TratarErro:
    DebugError "clsDataMapper.LoadServidor: " & Err.Description
    Set LoadServidor = Nothing
End Function
