VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORMProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mValue As Variant
Private mAnnotations() As ORMAnnotation
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

Public Property Get Annotations() As ORMAnnotation()
    Annotations = mAnnotations
End Property

Public Property Let Annotations(values() As ORMAnnotation)
    mAnnotations = values
End Property
Public Sub AddAnnotation(ByVal annotationDef As AnnotationDefinition, ByVal annotationValue As Variant)
    Dim annotation As New ORMAnnotation
    annotation.Definition = annotationDef
    annotation.value = annotationValue
    
    ReDim Preserve mAnnotations(annotationsCount) As ORMAnnotation
    
    Set mAnnotations(annotationsCount) = annotation

    annotationsCount = annotationsCount + 1
End Sub
Public Function isRequired() As Boolean
    isRequired = CheckAnnotationValue(Required, True)
End Function

Public Function isPrimaryKey() As Boolean
    isPrimaryKey = CheckAnnotationValue(PrimaryKey, True)
End Function

Public Function isComputed() As Boolean
    isComputed = CheckAnnotationValue(Computed, True)
End Function

Public Function CheckErrors() As String
    Dim i As Integer
    Dim result As String
    result = ""
    
    For i = 0 To annotationsCount - 1
        If mAnnotations(i).Definition = AnnotationDefinition.validator Then
            If Not CallByName(objValidator, "" & mAnnotations(i).value, VbMethod, mValue) Then
                result = result & mName & " value is invalid" & vbNewLine
            End If
        End If
    Next i
    CheckErrors = result
 
End Function

Private Function CheckAnnotationValue(anotation As AnnotationDefinition, value As Variant) As Boolean
    CheckAnnotationValue = False
    Dim i As Integer
    For i = 0 To annotationsCount - 1
        If mAnnotations(i).Definition = anotation Then
            If mAnnotations(i).value = True Then
                CheckAnnotationValue = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetAnnotationValue(anotation As AnnotationDefinition) As Variant
    Dim i As Integer
    GetAnnotationValue = Nothing
    For i = 0 To annotationsCount
        If mAnnotations(i).Definition = anotation Then
            GetAnnotationValue = mAnnotations(i).value
        End If
    Next i
End Function
Private Sub Class_Initialize()
    annotationsCount = 0
End Sub
