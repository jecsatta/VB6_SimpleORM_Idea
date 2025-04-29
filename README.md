# VB6 Simple ORM Idea
Just a simple idea of how to create dynamic classes using vb6 to perform access and fetch data stored in a database and little more.

## Usage
```vb6
Dim objArray As Variant
Dim objClient As clsClient
Dim objEmployee As clsEmployee
Dim i As Long
Dim params As New ORMParams

Set objClient = GetCreate(clsClienteType)

'objClient.ID = 1
objClient.Name = "Updated client"
objClient.Age = 35
objClient.Email = "client@example.com"
 'Save the client (insert or update based on primary key)
If objClient.AsORMBaseClass.Save Then
    Debug.Print "Client saved successfully!"
Else
    Debug.Print "Error saving client."
End If

params.Add "name", "%mary%", "ilike"
params.Add "id", 6
objArray = DataBaseSelect(clsClienteType, params)
If HasValues(objArray) Then
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
End If

'Example of execution
Count:1
ID:6
Name:Mary
Age:31
Email:mary@example.com
Errors:Age value is invalid
```

#### Classes setup
Databases classes must implement `ORMBaseClass`

In `Class_Initialize` all properties and annotations must be added
`File:clsClient.bas`
```vb6

Implements ORMBaseClass
Private mProperties As ORMProperties

Public Sub Class_Initialize()
    Set mProperties = New ORMProperties
    mProperties.Add "id", PrimaryKey, True
    mProperties.Add "Name"
    mProperties.Add "Age", validator, ValidatorAge
    mProperties.Add "Email"
End Sub

Private Function ORMBaseClass_GetTableName() As String
        ORMBaseClass_GetTableName = "client"
End Function

Private Function ORMBaseClass_Props() As ORMProperties
      Set ORMBaseClass_Props = mProperties
End Function
```

#### Properties implementation
```vb6
Public Property Let ID(ByVal value As Variant)
    mProperties.ValueByName("ID") = value
End Property

Public Property Get ID() As Variant
    ID = mProperties.ValueByName("ID")
End Property
```
#### Validators 
All Validator must be in `ORMValidator.cls` and in `ORMValidators.bas`

`File:ORMValidator.cls`
```vb6
Public Function ValidatorAge(value As Variant) As Boolean
    ValidatorAge = (value >= 40)
End Function
```
`File:ORMValidators.bas`
```vb6
Public Const ValidatorAge As String = "ValidatorAge"
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
