Attribute VB_Name = "Module1"
Option Explicit
Public cn                As New ADODB.Connection
Dim vConnectionString As String

Sub test()
    Dim startTime As Single
    Dim endTime As Single
    Dim elapsedTime As Single
    Dim objArray As Variant

    startTime = Timer
    objArray = ORMSelectSQL(clsClienteType, "select * from client order by id")
    endTime = Timer
    elapsedTime = endTime - startTime
    Debug.Print "Execution time: " & Format(elapsedTime, "0.000") & " seconds"
    
    startTime = Timer
    objArray = ORMSelectSQLWithProps(clsClienteType, "select * from client order by id")
    endTime = Timer
    elapsedTime = endTime - startTime
    Debug.Print "Execution time: " & Format(elapsedTime, "0.000") & " seconds"
End Sub
Public Sub Main()
    Dim objArray As Variant
    Dim objClient As clsClient
    Dim objEmployee As clsEmployee
    Dim i As Long
    
    vConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=Conexao"
    cn.ConnectionString = vConnectionString
    cn.Open
    
    cn.Execute "create table if not exists client(id serial primary key, name text,age int,email text);"
    cn.Execute "create table if not exists employee(id serial primary key, name text,position text,email text);"

'    cn.Execute "insert into client values(default,'Jhon',28,'jhon@example.com')"
'    cn.Execute "insert into client values(default,'Mary',31,'mary@example.com')"

'    cn.Execute "insert into employee values(default,'Jhon','CTO','mary@example.com')"
'    cn.Execute "insert into employee values(default,'Mary','CFO','mary@example.com')"
    Set objClient = GetCreate(clsClienteType)
    
    objClient.ID = 1
    objClient.Name = "Updated client"
    objClient.Age = 35
    objClient.Email = "client@example.com"
     'Save the client (insert or update based on primary key)
    If objClient.AsORMBaseClass.Save Then
        Debug.Print "Client saved successfully!"
    Else
        Debug.Print "Error saving client."
    End If
    
    objArray = ORMSelectSQL(clsClienteType, "select * from client order by id")
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
 
 
 
     
    Dim params As New ORMParams
    params.Add "name", "%New%", "ilike"
  '  params.Add "id", 6
    objArray = ORMSelect(clsClienteType, params)
    
    Debug.Print "Count:" & UBound(objArray)
    For i = 1 To UBound(objArray)
       Set objClient = objArray(i)
       Debug.Print "ID:" & objClient.ID
       Debug.Print "Name:" & objClient.Name
       Debug.Print "Age:" & objClient.Age
       Debug.Print "Email:" & objClient.Email
       Debug.Print "Errors:"; objClient.AsORMBaseClass.CheckErrors
       Debug.Print ""
    Next i

    
    objArray = ORMSelectSQL(clsEmployeeType, "select * from employee")
    
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


End Sub
