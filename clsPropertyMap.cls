VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPropertyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Classe.....: clsPropertyMap
' Descrição..: Metadados de uma propriedade persistida
'==============================================================================

Public Enum ePropertyType
    ptString = 1
    ptLong
    ptDouble
    ptBoolean
    ptDate
End Enum

Private mColumnName As String
Private mPropertyName As String
Private mDataType As ePropertyType
Private mRequired As Boolean
Private mDefaultValue As Variant
Private mMaxLength As Long
Private mReadOnly As Boolean

Public Property Get ColumnName() As String: ColumnName = mColumnName: End Property
Public Property Let ColumnName(ByVal Value As String): mColumnName = Trim$(Value): End Property

Public Property Get PropertyName() As String: PropertyName = mPropertyName: End Property
Public Property Let PropertyName(ByVal Value As String): mPropertyName = Trim$(Value): End Property

Public Property Get DataType() As ePropertyType: DataType = mDataType: End Property
Public Property Let DataType(ByVal Value As ePropertyType): mDataType = Value: End Property

Public Property Get Required() As Boolean: Required = mRequired: End Property
Public Property Let Required(ByVal Value As Boolean): mRequired = Value: End Property

Public Property Get DefaultValue() As Variant: DefaultValue = mDefaultValue: End Property
Public Property Let DefaultValue(ByVal Value As Variant): mDefaultValue = Value: End Property

Public Property Get MaxLength() As Long: MaxLength = mMaxLength: End Property
Public Property Let MaxLength(ByVal Value As Long): mMaxLength = Value: End Property

Public Property Get ReadOnly() As Boolean: ReadOnly = mReadOnly: End Property
Public Property Let ReadOnly(ByVal Value As Boolean): mReadOnly = Value: End Property

Public Function ToString() As String
    ToString = mPropertyName & " -> " & mColumnName
End Function

