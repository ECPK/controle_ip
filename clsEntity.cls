VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: clsEntity
' Projeto....: InfraControl SES
' Build......: 2026.07.22.001
' Descrição..: Classe base para entidades do domínio
' Compatível.: Excel 2016+
'==============================================================================

'==================== CAMPOS PRIVADOS =========================================

Private mID As String
Private mDataCadastro As Date
Private mDataAtualizacao As Date

Private mIsNew As Boolean
Private mIsLoaded As Boolean
Private mIsDirty As Boolean

'==================== EVENTOS =================================================

Private Sub Class_Initialize()
    Clear
End Sub

'==================== PROPRIEDADES ============================================

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(ByVal Value As String)

    Value = Trim$(Value)

    If mID <> Value Then
        mID = Value
        MarkDirty
    End If

End Property

Public Property Get DataCadastro() As Date
    DataCadastro = mDataCadastro
End Property

Public Property Let DataCadastro(ByVal Value As Date)

    If mDataCadastro <> Value Then
        mDataCadastro = Value
        MarkDirty
    End If

End Property

Public Property Get DataAtualizacao() As Date
    DataAtualizacao = mDataAtualizacao
End Property

Public Property Let DataAtualizacao(ByVal Value As Date)

    If mDataAtualizacao <> Value Then
        mDataAtualizacao = Value
        MarkDirty
    End If

End Property

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = mIsLoaded
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

'==================== MÉTODOS PÚBLICOS ========================================

Public Sub Clear()

    mID = vbNullString
    mDataCadastro = 0
    mDataAtualizacao = 0

    mIsNew = True
    mIsLoaded = False
    mIsDirty = False

End Sub

Public Sub Touch()

    mDataAtualizacao = Now

    If mDataCadastro = 0 Then
        mDataCadastro = Now
    End If

    MarkDirty

End Sub

Public Sub MarkDirty()

    mIsDirty = True

End Sub

Public Sub AcceptChanges()

    mIsDirty = False
    mIsLoaded = True
    mIsNew = False

End Sub

Public Function Validate() As Boolean

    Validate = (Len(mID) > 0)

End Function

Public Function Equals(ByVal Entity As clsEntity) As Boolean

    If Entity Is Nothing Then Exit Function

    Equals = (Me.ID = Entity.ID)

End Function

Public Function ToString() As String

    ToString = TypeName(Me) & _
               " [ID=" & mID & "]"

End Function

Public Function Clone() As clsEntity

    Dim obj As New clsEntity

    obj.ID = Me.ID
    obj.DataCadastro = Me.DataCadastro
    obj.DataAtualizacao = Me.DataAtualizacao

    If Me.IsDirty Then obj.MarkDirty
    If Not Me.IsNew Then obj.AcceptChanges

    Set Clone = obj

End Function

'==================== FRIEND ==================================================

Friend Sub SetLoaded()

    mIsLoaded = True
    mIsNew = False
    mIsDirty = False

End Sub

Friend Sub SetNew()

    mIsNew = True
    mIsLoaded = False
    mIsDirty = False

End Sub

