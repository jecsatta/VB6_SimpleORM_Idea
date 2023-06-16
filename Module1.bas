Attribute VB_Name = "Module1"
Option Explicit
Public cn                As New ADODB.Connection
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
           Debug.Print "Errors:"; objClient.AsIBaseClass.CheckErrors
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
