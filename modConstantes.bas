Attribute VB_Name = "modConstantes"
Option Explicit

'==============================================================================
' InfraControl SES
' Módulo.......: modConstantes
' Descrição....: Constantes globais da aplicação
' Compatibilidade: Microsoft Excel 2016+
'==============================================================================

Public Const APP_NOME As String = "InfraControl SES"
Public Const APP_SIGLA As String = "ICS"
Public Const APP_EMPRESA As String = "SES-DF"
Public Const APP_AUTOR As String = "Equipe InfraControl"
Public Const APP_VERSAO As String = "1.0.0"
Public Const APP_BUILD As String = "2026.07.22.001"

Public Const FRAMEWORK_NOME As String = "ICS Framework"
Public Const FRAMEWORK_VERSAO As String = "2.4.0"

Public Const APP_DEBUG As Boolean = True
Public Const APP_TIMEOUT As Long = 30
Public Const APP_MAX_LOG As Long = 10000
Public Const APP_MAX_IMPORTACAO As Long = 100000

Public Const APP_DATA_FORMATO As String = "dd/mm/yyyy"
Public Const APP_HORA_FORMATO As String = "hh:mm:ss"
Public Const APP_DATAHORA_FORMATO As String = "dd/mm/yyyy hh:mm:ss"

Public Const APP_SEPARADOR As String = "============================================================"

Public Const APP_EXTENSAO_XLSM As String = ".xlsm"
Public Const APP_EXTENSAO_XLSX As String = ".xlsx"
Public Const APP_EXTENSAO_CSV As String = ".csv"

Public Const WS_INICIO As String = "Inicio"
Public Const WS_BASE_IPS As String = "Base_IPs"
Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_CONFIG As String = "Config"
Public Const WS_LOG As String = "LOG"

Public Const TB_IPS As String = "tbIPs"
Public Const TB_CONFIG As String = "tbConfig"

Public Const UF_PRINCIPAL As String = "frmPrincipal"
Public Const UF_SERVIDOR As String = "frmServidor"
Public Const UF_IP As String = "frmIP"

Public Const STATUS_LIVRE As String = "LIVRE"
Public Const STATUS_UTILIZADO As String = "UTILIZADO"
Public Const STATUS_RESERVADO As String = "RESERVADO"
Public Const STATUS_INATIVO As String = "INATIVO"

Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARNING As String = "WARNING"
Public Const LOG_ERROR As String = "ERROR"
Public Const LOG_DEBUG As String = "DEBUG"
Public Const LOG_SUCCESS As String = "SUCCESS"

Public Const ACTION_INSERT As String = "INSERT"
Public Const ACTION_UPDATE As String = "UPDATE"
Public Const ACTION_DELETE As String = "DELETE"
Public Const ACTION_IMPORT As String = "IMPORT"
Public Const ACTION_EXPORT As String = "EXPORT"

Public Const APP_STATUS_INICIALIZANDO As String = "Inicializando"
Public Const APP_STATUS_PRONTO As String = "Pronto"
Public Const APP_STATUS_PROCESSANDO As String = "Processando"
Public Const APP_STATUS_FINALIZADO As String = "Finalizado"
Public Const APP_STATUS_ERRO As String = "Erro"

Public Const MSG_SUCESSO As String = "Operação realizada com sucesso."
Public Const MSG_ERRO As String = "Ocorreu um erro durante a operação."
Public Const MSG_CONFIRMACAO As String = "Deseja continuar?"

' Build 003 - Importação
' Build 004 - Dashboard
' Build 005 - Cadastro
' Build 006 - Pesquisa
' Build 007 - Relatórios
