VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORMProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Properties() As ORMProperty
Private propNumber As Long

Public Function GetPropertyIndex(ByVal propName As String) As Integer
    Dim i As Integer
    For i = LBound(Properties) To UBound(Properties)
        If LCase(Properties(i).Name) = LCase(propName) Then
            GetPropertyIndex = i
            Exit Function
        End If
    Next i
    GetPropertyIndex = -1
End Function

Public Function CheckErrors() As String
    Dim i As Integer
    Dim result As String
    
    result = ""
    For i = LBound(Properties) To UBound(Properties)
        result = result + Properties(i).CheckErrors
    Next i
    CheckErrors = result
End Function

Public Property Get Value(ByVal propName As String) As Variant
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    Value = Properties(propIndex).Value
End Property

Public Property Let Value(ByVal propName As String, ByVal propValue As Variant)
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    Properties(propIndex).Value = propValue
End Property

Public Property Get Prop(ByVal propName As String) As ORMProperty
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    Prop = Properties(propIndex)
End Property

Public Sub AddWithValue(ByVal propName As String, ByVal propValue As Variant, Optional ByVal annotationDef As ORMAnnotationType = None, Optional ByVal annotationValue As Variant)
    Add propName, annotationDef, annotationValue
    Properties(propNumber - 1).Value = propValue
End Sub

Public Function Add(ByVal propName As String, Optional ByVal annotationDef As ORMAnnotationType = None, Optional ByVal annotationValue As Variant) As ORMProperty
    ReDim Preserve Properties(propNumber) As ORMProperty
    Dim Prop As New ORMProperty
    Prop.Name = propName
    
    If annotationDef <> None Then
        Prop.AddAnnotation annotationDef, annotationValue
    End If
    Set Properties(propNumber) = Prop
    propNumber = propNumber + 1
    Set Add = Prop
End Function
 
Public Function GetProperties() As ORMProperty()
    GetProperties = Properties
End Function

Public Function GetPropertyValues() As Variant()
    Dim values() As Variant
    Dim i As Integer
    
    ReDim values(LBound(Properties) To UBound(Properties)) As Variant
    For i = LBound(Properties) To UBound(Properties)
        values(i) = Properties(i).Value
    Next i
    
    GetPropertyValues = values
End Function


Public Property Get count() As Integer
    count = propNumber
End Property

Private Sub Class_Initialize()
    propNumber = 0
End Sub
