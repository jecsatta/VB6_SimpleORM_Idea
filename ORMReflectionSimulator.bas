Attribute VB_Name = "ORMReflectionSimulator"
Option Explicit

Public Function GetCreate(ClassName As String) As Object
    Dim factory As New ORMBaseClassFactory
    If HasMethod(factory, ClassName & "_Create") Then
        Set GetCreate = CallByName(factory, ClassName & "_Create", VbMethod)
    End If
    Set factory = Nothing
End Function

Public Function HasProperty(Obj As Object, propName As String) As Boolean
    On Error Resume Next
    Dim propValue As Variant
    Err.Clear
    
     propValue = CallByName(Obj, propName, VbGet)
     
    If Err.Number = 0 Then
        HasProperty = True
    Else
        HasProperty = False
        Err.Clear
    End If
    On Error GoTo 0
End Function


Public Function HasMethod(ByVal Obj As Object, ByVal methodName As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    Err.Clear
    Set temp = CallByName(Obj, methodName, VbMethod)
    HasMethod = (Err.Number = 0)
    On Error GoTo 0
End Function


Public Sub ORMPropertyLet(Obj As Object, propName As String, value As Variant)
    On Error GoTo quit
    Dim propValue As Variant
    propValue = CallByName(Obj, propName, VbGet)
    CallByName Obj, propName, VbLet, value
    On Error GoTo 0
    Exit Sub
quit:
    On Error GoTo 0
End Sub
