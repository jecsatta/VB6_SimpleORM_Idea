Attribute VB_Name = "ORMReflectionSimulator"
Option Explicit

Public Function HasProperty(obj As Object, propName As String) As Boolean
    On Error Resume Next
    Dim propValue As Variant
    Err.Clear
    
     propValue = CallByName(obj, propName, VbGet)
     
    If Err.Number = 0 Then
        HasProperty = True
    Else
        HasProperty = False
        Err.Clear
    End If
    On Error GoTo 0
End Function


Public Function HasMethod(ByVal obj As Object, ByVal methodName As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    Err.Clear
    Set temp = CallByName(obj, methodName, VbMethod)
    HasMethod = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub ORMPropertyLet(obj As Object, propName As String, value As Variant)
    On Error GoTo quit
    Dim propValue As Variant
    propValue = CallByName(obj, propName, VbGet)
    CallByName obj, propName, VbLet, value
    On Error GoTo 0
    Exit Sub
quit:
    On Error GoTo 0
End Sub
