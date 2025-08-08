Attribute VB_Name = "Module1"
Option Explicit

Public gFactory As New MyFactory
Public gcore As New ormcore
Public gValidator As New MyValidator

Sub Main()
    Dim objArray As Variant
    Dim objClient As clsClient
    Dim objEmployee As clsEmployee
    Dim i As Long
    Dim connStr As String
    connStr = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=Conexao"
    gcore.Initialize connStr, "POSTGRES", gFactory, gValidator
    gcore.ExecuteNonQuery "create table if not exists client(id serial primary key, name text,age int,email text);"
    gcore.ExecuteNonQuery "create table if not exists employee(id serial primary key, name text,position text,email text);"
    

    'gcore.ExecuteNonQuery "insert into client values(default,'Jhon',28,'jhon@example.com')"
    'gcore.ExecuteNonQuery "insert into client values(default,'Mary',31,'mary@example.com')"

    'gcore.ExecuteNonQuery "insert into employee values(default,'Jhon','CTO','mary@example.com')"
    'gcore.ExecuteNonQuery "insert into employee values(default,'Mary','CFO','mary@example.com')"
    
    Set objClient = gFactory.clsClient_Create
    

    objClient.ID = 1
    objClient.Name = "Updated client"
    objClient.Age = 45
    objClient.Email = "client@example.com"
     'Save the client (insert or update based on primary key)
    If objClient.AsIORMEntity.Save Then
        Debug.Print "Client saved successfully!"
    Else
        Debug.Print "Error saving client."
    End If
    
    objArray = gcore.QuerySQL(gFactory.clsClientType, "select * from client order by id")
    Debug.Print "Count:" & UBound(objArray)
    Dim te As Object
    For i = 0 To UBound(objArray)
      Set objClient = objArray(i)
      
       Debug.Print "ID:" & objClient.ID
       Debug.Print "Name:" & objClient.Name
       Debug.Print "Age:" & objClient.Age
       Debug.Print "Email:" & objClient.Email
       Debug.Print "Errors:" & objClient.AsIORMEntity.CheckErrors
       Debug.Print ""
    Next i
 
 
 
     
    Dim params As New ORMParams
    params.Add "name", "%New%", "ilike"
  '  params.Add "id", 6
    objArray = gcore.QueryParams(gFactory.clsClientType, params)
    
    Debug.Print "Count:" & UBound(objArray)
    For i = 1 To UBound(objArray)
       Set objClient = objArray(i)
       Debug.Print "ID:" & objClient.ID
       Debug.Print "Name:" & objClient.Name
       Debug.Print "Age:" & objClient.Age
       Debug.Print "Email:" & objClient.Email
       Debug.Print "Errors:"; objClient.AsIORMEntity.CheckErrors
       Debug.Print ""
    Next i

    
    objArray = gcore.QuerySQL(gFactory.clsEmployeeType, "select * from employee")
    
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
