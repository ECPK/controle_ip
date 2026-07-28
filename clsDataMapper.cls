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

    On Error GoTo TratarErro

    If Servidor Is Nothing Then Exit Function
    If Not Servidor.Validate Then Exit Function

    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_HOSTNAME, Servidor.Hostname
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_FQDN, Servidor.FQDN
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_DESCRICAO, Servidor.Descricao
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_AMBIENTE, Servidor.Ambiente
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_SO, Servidor.SistemaOperacional
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_VERSAO_SO, Servidor.VersaoSO
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_ARQUITETURA, Servidor.Arquitetura
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_CPU, Servidor.CPU
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_MEMORIA, Servidor.MemoriaGB
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_OBSERVACOES, Servidor.Observacoes
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_DATA_CADASTRO, Servidor.DataCadastro
    WriteFieldByKey TB_SERVIDORES, COL_SERV_ID, Servidor.ID, COL_SERV_DATA_ATUALIZACAO, Servidor.DataAtualizacao

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

    On Error GoTo TratarErro

    Dim S As New clsServidor

    S.ID = ID
    S.Hostname = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_HOSTNAME))
    S.FQDN = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_FQDN))
    S.Descricao = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DESCRICAO))
    S.Ambiente = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_AMBIENTE))
    S.SistemaOperacional = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_SO))
    S.VersaoSO = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_VERSAO_SO))
    S.Arquitetura = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_ARQUITETURA))
    S.CPU = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_CPU))
    S.MemoriaGB = Val(CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_MEMORIA)))
    S.Observacoes = CStr(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_OBSERVACOES))
    If IsDate(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DATA_CADASTRO)) Then S.DataCadastro = CDate(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DATA_CADASTRO))
    If IsDate(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DATA_ATUALIZACAO)) Then S.DataAtualizacao = CDate(ReadFieldByKey(TB_SERVIDORES, COL_SERV_ID, ID, COL_SERV_DATA_ATUALIZACAO))

    S.AcceptChanges

    Set LoadServidor = S
    Exit Function

TratarErro:
    DebugError "clsDataMapper.LoadServidor: " & Err.Description
    Set LoadServidor = Nothing
End Function
