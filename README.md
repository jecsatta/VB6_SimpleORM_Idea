# VB6 Simple ORM Idea
Just a simple idea of how to create dynamic classes using vb6 to perform access and fetch data stored in a database.

:warning: *This is a working in progress*

## Usage
```vb6
Dim objArray As Variant
Dim objClient As clsClient
Dim i As Long
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
```
Databases classes must implement `IBaseClass`

In `Class_Initialize` all properties and annotations must be added
```vb6
'Class:clsClient

Implements IBaseClass
Private namedClass As NamedPropertiesClass

Public Sub Class_Initialize()
    Set namedClass = New NamedPropertiesClass
    namedClass.Add("id", PrimaryKey, True).AddAnnotation AutoIncrement, True
    namedClass.Add "Name"
    namedClass.Add "Age"
    namedClass.Add "Email"
End Sub

Private Function IBaseClass_GetTableName() As String
        IBaseClass_GetTableName = "client"
End Function

Private Function IBaseClass_Props() As NamedPropertiesClass
      Set IBaseClass_Props = namedClass
End Function
```

Property implementation
```vb6
Public Property Let ID(ByVal value As Variant)
    namedClass.PropertyByName("ID") = value
End Property

Public Property Get ID() As Variant
    ID = namedClass.PropertyByName("ID")
End Property
```


| To-do | Status |
| --- | :---: |
| Dynamic properties  | :white_check_mark: |
| Select with simple text  | :white_check_mark: |
| Select with parametrized query|  :white_check_mark: |
| annotations support| :white_check_mark: |
| create routine||
| update routine||
| insert routine|working|
| transaction support||

License: [GNU General Public License v3.0](LICENSE)
