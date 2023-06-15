VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NamedPropertiesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private NamedProperties() As Variant
Private PropertyNames() As String


Public Function GetPropertyIndex(ByVal propName As String) As Integer
    Dim i As Integer
    For i = LBound(PropertyNames) To UBound(PropertyNames)
        If PropertyNames(i) = propName Then
            GetPropertyIndex = i
            Exit Function
        End If
    Next i
    GetPropertyIndex = -1
End Function

Public Property Let ConfigPropertyNames(PROPERTY_NAMES As String)
    PropertyNames = Split(PROPERTY_NAMES, ",")
    ReDim NamedProperties(UBound(PropertyNames)) As Variant
End Property


Public Property Get PropertyByName(ByVal propName As String) As Variant
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    PropertyByName = NamedProperties(propIndex)
End Property

Public Property Let PropertyByName(ByVal propName As String, ByVal propValue As Variant)
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    NamedProperties(propIndex) = propValue
End Property

Public Property Get Count() As Integer
    Count = UBound(PropertyNames)
End Property

