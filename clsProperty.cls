VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mName As String
Private mValue As Variant
Private mAnnotations() As clsAnnotation
Private annotationsCount As Long
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get value() As Variant
    value = mValue
End Property

Public Property Let value(ByVal value As Variant)
    mValue = value
End Property

Public Property Get Annotations() As clsAnnotation()
    Annotations = mAnnotations
End Property

Public Property Let Annotations(values() As clsAnnotation)
    mAnnotations = values
End Property
Public Sub AddAnnotation(ByVal annotationDef As AnnotationDefinition, ByVal annotationValue As Variant)
    Dim annotation As New clsAnnotation
    annotation.Definition = annotationDef
    annotation.value = annotationDef
    
    ReDim Preserve mAnnotations(annotationsCount) As clsAnnotation
    
    Set mAnnotations(annotationsCount) = annotation

    annotationsCount = annotationsCount + 1
End Sub

Private Sub Class_Initialize()
    annotationsCount = 0
End Sub
