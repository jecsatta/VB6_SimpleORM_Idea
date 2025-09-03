# VB6 Simple ORM

A powerful yet simple idea of Object-Relational Mapping library for Visual Basic 6, designed for modern development practices with DLL compilation support, dependency injection, and advanced ORM features.

## Usage in 
`Module1.bas`
```vb6
Option Explicit

Public gFactory As New MyFactory
Public gcore As New ORMCore
Public gValidator As New MyValidator

Sub Main()
  Dim objArray As Variant
  Dim objClient As clsClient
  Dim objEmployee As clsEmployee
  Dim i As Long
  Dim params As New ORMParams
  Dim connStr As String
  
  connStr = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=Conn"
  gcore.Initialize connStr, "POSTGRES", gFactory, gValidator

  Set objClient = gFactory.clsClient_Create

  objClient.ID = 1
  objClient.Name = "Updated client"
  objClient.Age = 45
  objClient.Email = "client@example.com"
   'Save the client (insert or update based on primary key)
  If objClient.AsEntity.Save Then
      Debug.Print "Client saved successfully!"
  Else
      Debug.Print "Error saving client."
  End If

  params.Add "name", "%mary%", "ilike"
  params.Add "id", 6
  objArray = gcore.QueryParams(gFactory.clsClientType, params)
  
  Debug.Print "Count:" & UBound(objArray)
  For i = 1 To UBound(objArray)
     Set objClient = objArray(i)
     Debug.Print "ID:" & objClient.ID
     Debug.Print "Name:" & objClient.Name
     Debug.Print "Age:" & objClient.Age
     Debug.Print "Email:" & objClient.Email
     Debug.Print "Errors:" & objClient.AsEntity.CheckErrors
     Debug.Print ""
  Next i
End Sub

'Example of execution
Count:1
ID:6
Name:Mary
Age:31
Email:mary@example.com
Errors:Age value is invalid
```

#### Classes setup
Entity classes must implement `IORMEntity`

In `Class_Initialize` all properties and annotations must be added
`File:clsClient.cls`
```vb6
Implements IORMEntity
Private mProperties As ORMProperties
Public Sub Class_Initialize()
    Set mProperties = New ORMProperties
    mProperties.Add "id", PrimaryKey, True
    mProperties.Add "Name"
    mProperties.Add "Age", Validator, gValidator.ValidatorAge
    mProperties.Add "Email"   
End Sub
Private Function IORMEntity_GetTableName() As String
        IORMEntity_GetTableName = "client"
End Function
Private Function IORMEntity_Props() As ORMProperties
      Set IORMEntity_Props = mProperties
End Function
```
You must have a concrete Factory class of IORMFactory with all yours entity classes e.g.: `MyFactory.bas`

```vb6
Implements IORMFactory

Public Property Get clsClientType() As String: clsClientType = "clsClient": End Property
Public Function clsClient_Create() As Object: Set clsClient_Create = New clsClient: End Function

Public Property Get clsEmployeeType() As String: clsEmployeeType = "clsEmployee": End Property
Public Function clsEmployee_Create() As Object: Set clsEmployee_Create = New clsEmployee: End Function
```

#### Properties implementation
```vb6
Public Property Let ID(ByVal value As Variant):    mProperties.value("ID") = value: End Property
Public Property Get ID() As Variant:    ID = mProperties.value("ID"): End Property

Public Property Let Name(ByVal value As Variant): mProperties.value("Name") = value: End Property
Public Property Get Name() As Variant: Name = mProperties.value("Name"): End Property
```
#### Validators 
Your concrete validator class must implement IORMValidator

`File:MyValidator.cls`
```vb6
Option Explicit

Implements IORMValidator

'Validator Age
Public Property Get ValidatorAge(): ValidatorAge = "ValidatorAgeFunction": End Property
Public Function ValidatorAgeFunction(value As Variant) As Boolean
    ValidatorAgeFunction = (value >= 40)
End Function
```

| To-do | Status |
| --- | :---: |
| Dynamic properties  | :white_check_mark: |
| Select with simple text  | :white_check_mark: |
| Select with parametrized query|  :white_check_mark: |
| annotations support| :white_check_mark: |
| create routine| in future|
| update routine|:white_check_mark:|
| insert routine|:white_check_mark:|
| transaction support||
| external validators support| :white_check_mark: |

License: [GNU General Public License v3.0](LICENSE)
