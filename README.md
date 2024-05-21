# VB6 Simple ORM Idea
Just a simple idea of how to create dynamic classes using vb6 to perform access and fetch data stored in a database and little more.

## Usage
```vb6
Dim objArray As Variant
Dim objClient As clsClient
Dim i As Long
Dim params As New clsParams

Set objClient = GetCreate(C_clsClient)

'objClient.ID = 1
objClient.Name = "Updated client"
objClient.Age = 35
objClient.Email = "client@example.com"
 'Save the client (insert or update based on primary key)
If objClient.AsIBaseClass.Save Then
    Debug.Print "Client saved successfully!"
Else
    Debug.Print "Error saving client."
End If

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
       Debug.Print "Errors:" &  objClient.AsIBaseClass.CheckErrors
       Debug.Print ""
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
Databases classes must implement `IBaseClass`

In `Class_Initialize` all properties and annotations must be added
`File:clsClient.bas`
```vb6

Implements IBaseClass
Private namedClass As NamedPropertiesClass

Public Sub Class_Initialize()
    Set namedClass = New NamedPropertiesClass
    namedClass.Add("id", PrimaryKey, True).AddAnnotation AutoIncrement, True
    namedClass.Add "Name"
    namedClass.Add "Age", Validator, V_AgeValidator
    namedClass.Add "Email"
End Sub

Private Function IBaseClass_GetTableName() As String
        IBaseClass_GetTableName = "client"
End Function

Private Function IBaseClass_Props() As NamedPropertiesClass
      Set IBaseClass_Props = namedClass
End Function
```

#### Properties implementation
```vb6
Public Property Let ID(ByVal value As Variant)
    namedClass.PropertyByName("ID") = value
End Property

Public Property Get ID() As Variant
    ID = namedClass.PropertyByName("ID")
End Property
```
#### Validators 
All Validator must be in `clsValidator` and in `Validator_Constants`

`File:clsValidator.bas`
```vb6
Public Function Validator_Age(value As Variant) As Boolean
    Validator_Age = (value >= 40)
End Function
```
`File:Validator_Constants.bas`
```vb6
Public Const V_AgeValidator As String = "Validator_Age"
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
