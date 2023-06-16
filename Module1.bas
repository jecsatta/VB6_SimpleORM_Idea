Attribute VB_Name = "Module1"
Option Explicit
Dim cn                As New ADODB.Connection
Dim vConnectionString As String

Public Sub Main()
    Dim objArray As Variant
    Dim objClient As clsClient
    Dim objEmployee As clsEmployee
    Dim dictParams As New Dictionary
    Dim i As Long
    
    vConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostConnection"
    cn.ConnectionString = vConnectionString
    cn.Open
    cn.Execute "create table if not exists client(id serial primary key, name text,age int,email text);"
    cn.Execute "create table if not exists employee(id serial primary key, name text,position text,email text);"
    'cn.Execute "insert into client values(default,'Jhon',28,'jhon@example.com')"
    'cn.Execute "insert into client values(default,'Mary',31,'mary@example.com')"
    
    'cn.Execute "insert into employee values(default,'Jhon','CTO','mary@example.com')"
    'cn.Execute "insert into employee values(default,'Mary','CFO','mary@example.com')"
    
'    objArray = DataBaseSelectSQL(C_clsClient, "select * from client")
'    Debug.Print "Count:" & UBound(objArray)
'    For i = 1 To UBound(objArray)
'      Set objClient = objArray(i)
'       Debug.Print "ID:" & objClient.ID
'       Debug.Print "Name:" & objClient.Name
'       Debug.Print "Age:" & objClient.Age
'       Debug.Print "Email:" & objClient.Email
'       Debug.Print ""
'       Debug.Print ""
'    Next i
     
    Dim params As New clsParams
    params.Add "name", "%mary%", "ilike"
    params.Add "id", 6
    objArray = DataBaseSelect(C_clsClient, params)
    If HasValues(objArray) Then
        Debug.Print "Count:" & UBound(objArray)
        For i = 1 To UBound(objArray)
          Set objClient = objArray(i)
           Debug.Print "ID:" & objClient.ID
           Debug.Print "Name:" & objClient.Name
           Debug.Print "Age:" & objClient.Age
           Debug.Print "Email:" & objClient.Email
           Debug.Print ""
           Debug.Print ""
        Next i
    End If
    
    objArray = DataBaseSelectSQL(C_clsEmployee, "select * from employee")
    If HasValues(objArray) Then
        Debug.Print "Count:" & UBound(objArray)
        For i = 1 To UBound(objArray)
          Set objEmployee = objArray(i)
           Debug.Print "ID:" & objEmployee.ID
           Debug.Print "Name:" & objEmployee.Name
           Debug.Print "Position:" & objEmployee.Position
           Debug.Print "Email:" & objEmployee.Email
           Debug.Print ""
           Debug.Print ""
        Next i
     End If

End Sub
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
