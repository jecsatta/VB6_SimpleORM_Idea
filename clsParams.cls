VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private params() As clsParam
Private currentPosition As Long
Private count As Long

Public Sub First()
    currentPosition = LBound(params)
End Sub

Public Sub MoveNext()
    currentPosition = currentPosition + 1
End Sub

Public Function EOF() As Boolean
    EOF = (currentPosition > UBound(params))
End Function

Public Property Get Name() As String
    Name = params(currentPosition).Name
End Property

Public Property Get value() As Variant
    value = params(currentPosition).value
End Property

Public Property Get Operator() As String
    Operator = params(currentPosition).Operator
End Property

Public Sub Add(ByVal paramName As String, ByVal paramValue As Variant, Optional ByVal paramOperator As String = "=")
    Dim param As New clsParam
    param.Name = paramName
    param.value = paramValue
    param.Operator = paramOperator
    
     
 
    ReDim Preserve params(count) As clsParam
    Set params(count) = param
    count = count + 1
End Sub

Public Function GetParamNames() As String()
    Dim paramNames() As String
    Dim paramCount As Integer
    paramCount = UBound(params)
    
    ReDim paramNames(paramCount) As String
    
    Dim i As Integer
    For i = 1 To paramCount
        paramNames(i) = params(i).Name
    Next i
    
    GetParamNames = paramNames
End Function

Public Function GetParamValues() As Variant()
    Dim paramValues() As Variant
    Dim paramCount As Integer
    paramCount = UBound(params)
    
    ReDim paramValues(paramCount) As Variant
    
    Dim i As Integer
    For i = 1 To paramCount
        paramValues(i) = params(i).value
    Next i
    
    GetParamValues = paramValues
End Function

Public Function GetParamOperators() As String()
    Dim paramOperators() As String
    Dim paramCount As Integer
    paramCount = UBound(params)
    
    ReDim paramOperators(paramCount) As String
    
    Dim i As Integer
    For i = 1 To paramCount
        paramOperators(i) = params(i).Operator
    Next i
    
    GetParamOperators = paramOperators
End Function

Private Sub Class_Initialize()
    currentPosition = -1
    count = 0
End Sub
