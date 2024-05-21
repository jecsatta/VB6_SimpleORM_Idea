VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseClass


Private namedClass As NamedPropertiesClass
 
    
Public Sub Class_Initialize()
    Set namedClass = New NamedPropertiesClass
    namedClass.Add("id").AddAnnotation PrimaryKey, True
    
    namedClass.Add "Name"
    namedClass.Add "Position"
    namedClass.Add "Email"
   
   
End Sub
Public Function AsIBaseClass() As IBaseClass
    Set AsIBaseClass = Me
End Function

Private Function IBaseClass_CheckErrors() As String
    IBaseClass_CheckErrors = namedClass.CheckErrors
End Function

Private Function IBaseClass_GetTableName() As String
        IBaseClass_GetTableName = "employee"
End Function

Private Function IBaseClass_Props() As NamedPropertiesClass
        IBaseClass_Props = namedClass
End Function

Private Function IBaseClass_Save() As Boolean
    IBaseClass_Save = DataBaseSave(C_clsEmployee, Me)
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

Public Property Let Position(ByVal value As Variant)
    namedClass.PropertyByName("Position") = value
End Property

Public Property Get Position() As Variant
    Position = namedClass.PropertyByName("Position")
End Property

Public Property Let Email(ByVal value As Variant)
    namedClass.PropertyByName("Email") = value
End Property

Public Property Get Email() As Variant
    Email = namedClass.PropertyByName("Email")
End Property

