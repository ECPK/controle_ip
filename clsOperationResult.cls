VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Projeto.....: InfraControl SES Framework
' Classe......: clsOperationResult
' Camada......: Application.Results
' Versão......: 1.1.0
' Build.......: 2026.07.22.001
' Compatível..: Excel 2016+
'------------------------------------------------------------------------------
' Responsabilidade:
'   Representa o resultado padrão de qualquer operação do Framework.
'==============================================================================

Private mSuccess As Boolean
Private mCode As Long
Private mMessage As String
Private mData As Variant
Private mAffectedRecords As Long
Private mExecutionTime As Double

Private mErrors As Collection
Private mWarnings As Collection
Private mInformations As Collection

Private Sub Class_Initialize()
    Set mErrors = New Collection
    Set mWarnings = New Collection
    Set mInformations = New Collection
End Sub

Public Property Get Success() As Boolean
    Success = mSuccess
End Property

Public Property Let Success(ByVal Value As Boolean)
    mSuccess = Value
End Property

Public Property Get Code() As Long
    Code = mCode
End Property

Public Property Let Code(ByVal Value As Long)
    mCode = Value
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let Message(ByVal Value As String)
    mMessage = Value
End Property

Public Property Get Data() As Variant
    Data = mData
End Property

Public Property Let Data(ByVal Value As Variant)
    mData = Value
End Property

Public Property Get AffectedRecords() As Long
    AffectedRecords = mAffectedRecords
End Property

Public Property Let AffectedRecords(ByVal Value As Long)
    mAffectedRecords = Value
End Property

Public Property Get ExecutionTime() As Double
    ExecutionTime = mExecutionTime
End Property

Public Property Let ExecutionTime(ByVal Value As Double)
    mExecutionTime = Value
End Property

Public Property Get Errors() As Collection
    Set Errors = mErrors
End Property

Public Property Get Warnings() As Collection
    Set Warnings = mWarnings
End Property

Public Property Get Informations() As Collection
    Set Informations = mInformations
End Property

Public Sub AddError(ByVal Text As String)
    mErrors.Add Text
    mSuccess = False
End Sub

Public Sub AddWarning(ByVal Text As String)
    mWarnings.Add Text
End Sub

Public Sub AddInformation(ByVal Text As String)
    mInformations.Add Text
End Sub

Public Function HasErrors() As Boolean
    HasErrors = (mErrors.Count > 0)
End Function

Public Function HasWarnings() As Boolean
    HasWarnings = (mWarnings.Count > 0)
End Function

Public Function HasInformations() As Boolean
    HasInformations = (mInformations.Count > 0)
End Function

Public Sub Clear()
    Set mErrors = New Collection
    Set mWarnings = New Collection
    Set mInformations = New Collection
    mSuccess = False
    mCode = 0
    mMessage = vbNullString
    mAffectedRecords = 0
    mExecutionTime = 0
    mData = Empty
End Sub

Public Sub Merge(ByVal Source As clsOperationResult)
    Dim Item As Variant
    If Source Is Nothing Then Exit Sub

    For Each Item In Source.Errors
        mErrors.Add Item
    Next

    For Each Item In Source.Warnings
        mWarnings.Add Item
    Next

    For Each Item In Source.Informations
        mInformations.Add Item
    Next

    If Not Source.Success Then mSuccess = False
End Sub

Public Function Clone() As clsOperationResult
    Dim r As New clsOperationResult
    Dim V As Variant

    r.Success = mSuccess
    r.Code = mCode
    r.Message = mMessage
    r.AffectedRecords = mAffectedRecords
    r.ExecutionTime = mExecutionTime
    r.Data = mData

    For Each V In mErrors
        r.AddError CStr(V)
    Next

    For Each V In mWarnings
        r.AddWarning CStr(V)
    Next

    For Each V In mInformations
        r.AddInformation CStr(V)
    Next

    Set Clone = r
End Function

Public Function ToString() As String
    ToString = IIf(mSuccess, "SUCCESS", "FAILURE") & _
               " | Code=" & mCode & _
               " | Message=" & mMessage
End Function

Public Function CreateSuccess(Optional ByVal Msg As String = "") As clsOperationResult
    Clear
    mSuccess = True
    mMessage = Msg
    Set CreateSuccess = Me
End Function

Public Function CreateFailure(ByVal Msg As String, Optional ByVal ErrCode As Long = 0) As clsOperationResult
    Clear
    mSuccess = False
    mCode = ErrCode
    mMessage = Msg
    AddError Msg
    Set CreateFailure = Me
End Function

Public Function CreateWarning(ByVal Msg As String) As clsOperationResult
    Clear
    mSuccess = True
    AddWarning Msg
    mMessage = Msg
    Set CreateWarning = Me
End Function


