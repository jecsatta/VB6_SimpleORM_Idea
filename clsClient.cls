VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseClass

Private namedClass As NamedPropertiesClass
 

Public Sub Class_Initialize()
    Set namedClass = New NamedPropertiesClass
    namedClass.Add "id", PrimaryKey, True
    namedClass.Add "Name"
    namedClass.Add "Age", validator, V_AgeValidator
    namedClass.Add "Email"
    
End Sub
Public Function AsIBaseClass() As IBaseClass
    Set AsIBaseClass = Me
End Function

Private Function IBaseClass_CheckErrors() As String
    IBaseClass_CheckErrors = namedClass.CheckErrors
End Function

Private Function IBaseClass_GetTableName() As String
        IBaseClass_GetTableName = "client"
End Function

Private Function IBaseClass_Props() As NamedPropertiesClass
      Set IBaseClass_Props = namedClass
End Function

Private Function IBaseClass_Save() As Boolean
    IBaseClass_Save = DataBaseSave(C_clsClient, Me)
End Function


Public Property Let ID(ByVal value As Variant)
    namedClass.PropertyByName("ID") = value
End Property

Public Property Get ID() As Variant
    ID = namedClass.PropertyByName("ID")
End Property

Public Property Let Name(ByVal value As Variant)
    namedClass.PropertyByName("Name") = value
End Property

Public Property Get Name() As Variant
   Name = namedClass.PropertyByName("Name")
   
End Property

Public Property Let Age(ByVal value As Variant)
    namedClass.PropertyByName("Age") = value
End Property

Public Property Get Age() As Variant
    Age = namedClass.PropertyByName("Age")
End Property

Public Property Let Email(ByVal value As Variant)
    namedClass.PropertyByName("Email") = value
End Property

Public Property Get Email() As Variant
    Email = namedClass.PropertyByName("Email")
End Property


