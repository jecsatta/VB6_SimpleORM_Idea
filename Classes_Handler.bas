Attribute VB_Name = "Classes_Handler"
Public Function GetCreate(ClassName As String) As Object
    Dim factory As New BaseClassFactory
    If HasMethod(factory, ClassName & "_Create") Then
        Set GetCreate = CallByName(factory, ClassName & "_Create", VbMethod)
    End If
    Set factory = Nothing
End Function


