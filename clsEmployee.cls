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

Implements ORMBaseClass


Private mProperties As ORMProperties
 
    
Public Sub Class_Initialize()
    Set mProperties = New ORMProperties
    mProperties.Add("id").AddAnnotation PrimaryKey, True
    
    mProperties.Add "Name"
    mProperties.Add "Position"
    mProperties.Add "Email"
   
   
End Sub
Public Function AsORMBaseClass() As ORMBaseClass
    Set AsORMBaseClass = Me
End Function

Private Function ORMBaseClass_CheckErrors() As String
    ORMBaseClass_CheckErrors = mProperties.CheckErrors
End Function

Private Function ORMBaseClass_GetTableName() As String
        ORMBaseClass_GetTableName = "employee"
End Function

Private Function ORMBaseClass_Props() As ORMProperties
        ORMBaseClass_Props = mProperties
End Function

Private Function ORMBaseClass_Save() As Boolean
    ORMBaseClass_Save = DataBaseSave(clsEmployeeType, Me)
End Function

Public Property Let ID(ByVal value As Variant)
    mProperties.PropertyByName("ID") = value
End Property

Public Property Get ID() As Variant
    ID = mProperties.PropertyByName("ID")
End Property

Public Property Let Name(ByVal value As Variant)
    mProperties.PropertyByName("Name") = value
End Property

Public Property Get Name() As Variant
    Name = mProperties.PropertyByName("Name")
End Property

Public Property Let Position(ByVal value As Variant)
    mProperties.PropertyByName("Position") = value
End Property

Public Property Get Position() As Variant
    Position = mProperties.PropertyByName("Position")
End Property

Public Property Let Email(ByVal value As Variant)
    mProperties.PropertyByName("Email") = value
End Property

Public Property Get Email() As Variant
    Email = mProperties.PropertyByName("Email")
End Property

