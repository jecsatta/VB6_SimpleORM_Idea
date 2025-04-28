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
Attribute VB_Exposed = False
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

Public Property Get PropertyByName(ByVal propName As String) As Variant
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    PropertyByName = Properties(propIndex).value
End Property

Public Property Let PropertyByName(ByVal propName As String, ByVal propValue As Variant)
    Dim propIndex As Integer
    propIndex = GetPropertyIndex(propName)
    Properties(propIndex).value = propValue
End Property

Public Sub AddWithValue(ByVal propName As String, ByVal propValue As Variant, Optional ByVal annotationDef As AnnotationDefinition = None, Optional ByVal annotationValue As Variant)
    Add propName, annotationDef, annotationValue
    Properties(propNumber - 1).value = propValue
End Sub

Public Function Add(ByVal propName As String, Optional ByVal annotationDef As AnnotationDefinition = None, Optional ByVal annotationValue As Variant) As ORMProperty
    ReDim Preserve Properties(propNumber) As ORMProperty
    Dim prop As New ORMProperty
    prop.Name = propName
    
    If annotationDef <> None Then
        prop.AddAnnotation annotationDef, annotationValue
    End If
    Set Properties(propNumber) = prop
    propNumber = propNumber + 1
    Set Add = prop
End Function
 
Public Function GetProperties() As ORMProperty()
    GetProperties = Properties
End Function

Public Function GetPropertyValues() As Variant()
    Dim values() As Variant
    Dim i As Integer
    
    ReDim values(LBound(Properties) To UBound(Properties)) As Variant
    For i = LBound(Properties) To UBound(Properties)
        values(i) = Properties(i).value
    Next i
    
    GetPropertyValues = values
End Function


Public Property Get count() As Integer
    count = propNumber
End Property

Private Sub Class_Initialize()
    propNumber = 0
End Sub
