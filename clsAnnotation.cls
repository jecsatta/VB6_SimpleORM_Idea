VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnnotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDefinition As AnnotationDefinition
Private mValue As Variant

Public Property Get Definition() As AnnotationDefinition
    Definition = mDefinition
End Property

Public Property Let Definition(ByVal value As AnnotationDefinition)
    mDefinition = value
End Property

Public Property Get value() As Variant
    value = mValue
End Property

Public Property Let value(ByVal value As Variant)
    mValue = value
End Property

