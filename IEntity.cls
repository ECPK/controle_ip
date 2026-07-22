VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: IEntity
' Projeto....: InfraControl SES
' Build......: 2026.07.22.001
' Tipo.......: Interface
' Compatível.: Excel 2016+
'
' Esta classe define o contrato que todas as entidades do domínio deverão
' implementar utilizando:
'
'     Implements IEntity
'
'==============================================================================

'==================== IDENTIFICAÇÃO ===========================================

Public Property Get ID() As String
End Property

Public Property Let ID(ByVal Value As String)
End Property

'==================== METADADOS ===============================================

Public Property Get DataCadastro() As Date
End Property

Public Property Let DataCadastro(ByVal Value As Date)
End Property

Public Property Get DataAtualizacao() As Date
End Property

Public Property Let DataAtualizacao(ByVal Value As Date)
End Property

'==================== ESTADO ==================================================

Public Property Get IsNew() As Boolean
End Property

Public Property Get IsLoaded() As Boolean
End Property

Public Property Get IsDirty() As Boolean
End Property

'==================== MAPEAMENTO GENÉRICO =====================================

Public Function GetValue(ByVal PropertyName As String) As Variant
End Function

Public Sub SetValue(ByVal PropertyName As String, ByVal Value As Variant)
End Sub

'==================== CICLO DE VIDA ===========================================

Public Sub Clear()
End Sub

Public Sub Touch()
End Sub

Public Sub MarkDirty()
End Sub

Public Sub AcceptChanges()
End Sub

'==================== NEGÓCIO =================================================

Public Function Validate() As Boolean
End Function

Public Function Clone() As IEntity
End Function

Public Function Equals(ByVal Entity As IEntity) As Boolean
End Function

Public Function ToString() As String
End Function

