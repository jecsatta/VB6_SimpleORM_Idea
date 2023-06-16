Attribute VB_Name = "DatabaseMethods"
Option Explicit

Public Function DataBaseSelect(ClassName As String, params As clsParams, Optional intLimit As Integer = 0, Optional strOrderBy As String = "") As Variant
    Dim strSQL As String
    Dim tableName As String
    Dim i As Integer
    Dim lastParamValue As Variant
    Dim objReturn() As Variant
    Dim objClass As IBaseClass
    
    Set objClass = GetCreate(ClassName)
    
     tableName = objClass.GetTableName()
    strSQL = "SELECT * FROM " & tableName & " WHERE "

    params.First
    Do Until params.EOF
        strSQL = strSQL & params.Name & " " & params.Operator & " '" & params.value & "' AND "
        params.MoveNext
    Loop
    strSQL = Mid(strSQL, 1, Len(strSQL) - 4)


    If Not strOrderBy = "" And InStr(1, UCase(strOrderBy), "ORDER BY") > 1 Then
        strSQL = strSQL & strOrderBy
    End If

    
    If intLimit > 0 Then
        strSQL = strSQL & " LIMIT " & intLimit
    End If

    
   DataBaseSelect = DataBaseSelectSQL(ClassName, strSQL)
End Function
Public Function DataBaseSelectSQL(ClassName As String, strSQL As String) As Variant
    Dim rs As Object
    Dim objClass As Object
    Dim objReturn() As Variant
    Dim field As Object
    Dim i As Long
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cn
    
    If rs.EOF Then
        DataBaseSelectSQL = objReturn
        Exit Function
    End If
    
    
    i = 1
    Do Until rs.EOF
        Set objClass = GetCreate(ClassName)
        ReDim Preserve objReturn(i) As Variant
        For Each field In rs.Fields
            Dim propName As String
            propName = field.Name
            If HasProperty(objClass, UCase(propName)) Then
                CallByName objClass, propName, VbLet, field.value
            End If
        Next field
        
        Set objReturn(i) = objClass
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    Set rs = Nothing
    
    DataBaseSelectSQL = objReturn
End Function

