VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntityMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: clsEntityMap
' Descrição..: Coleção de metadados de persistência da entidade
'==============================================================================

Private mTableName As String
Private mKeyColumn As String
Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal Value As String)
    mTableName = Trim$(Value)
End Property

Public Property Get KeyColumn() As String
    KeyColumn = mKeyColumn
End Property

Public Property Let KeyColumn(ByVal Value As String)
    mKeyColumn = Trim$(Value)
End Property

Public Function AddProperty() As clsPropertyMap
    Dim P As New clsPropertyMap
    mItems.Add P
    Set AddProperty = P
End Function

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Property Get Item(ByVal Index As Long) As clsPropertyMap
    Set Item = mItems(Index)
End Property

Public Sub Clear()
    Set mItems = New Collection
End Sub

Public Sub ConfigureServidor()

    Dim P As clsPropertyMap

    Clear

    TableName = TB_SERVIDORES
    KeyColumn = COL_SERV_ID

    Set P = AddProperty()
    P.ColumnName = COL_SERV_HOSTNAME
    P.PropertyName = "Hostname"
    P.Required = True
    P.MaxLength = 60
    P.DataType = ptString

    Set P = AddProperty()
    P.ColumnName = COL_SERV_FQDN
    P.PropertyName = "FQDN"
    P.DataType = ptString

    Set P = AddProperty()
    P.ColumnName = COL_SERV_SO
    P.PropertyName = "SistemaOperacional"
    P.Required = True
    P.DataType = ptString

End Sub

Public Function ToString() As String
    ToString = mTableName & " (" & mItems.Count & " propriedades)"
End Function

