VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORMParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mValue As Variant
Private mOperator As String

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let value(ByVal value As Variant)
    mValue = value
End Property

Public Property Get value() As Variant
    value = mValue
End Property

Public Property Let Operator(ByVal value As String)
    mOperator = value
End Property

Public Property Get Operator() As String
    Operator = mOperator
End Property
