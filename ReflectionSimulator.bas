Attribute VB_Name = "ReflectionSimulator"
Public Function GetCreate(ClassName As String) As Object
    Dim factory As New BaseClassFactory
    If HasMethod(factory, ClassName & "_Create") Then
        Set GetCreate = CallByName(factory, ClassName & "_Create", VbMethod)
    End If
    Set factory = Nothing
End Function

Public Function HasValues(arr As Variant) As Boolean
    Dim n As Long
    On Error Resume Next
    Err.Clear
    n = LBound(arr)
    If Err.Number = 0 Then
        HasValues = True
    Else
        HasValues = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

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
