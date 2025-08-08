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

Implements IORMEntity

Private mProperties As ORMProperties
 

Public Sub Class_Initialize()
    Set mProperties = New ORMProperties
    mProperties.Add "id", PrimaryKey, True
    mProperties.Add "Name"
    mProperties.Add "Age", Validator, gValidator.ValidatorAge
    mProperties.Add "Email"
    
End Sub
Public Function AsIORMEntity() As IORMEntity
    Set AsIORMEntity = Me
End Function

Private Function IORMEntity_CheckErrors() As String
    IORMEntity_CheckErrors = mProperties.CheckErrors
End Function

Private Function IORMEntity_GetTableName() As String
        IORMEntity_GetTableName = "client"
End Function

Private Function IORMEntity_Props() As ORMProperties
      Set IORMEntity_Props = mProperties
End Function

Private Function IORMEntity_Save() As Boolean
        IORMEntity_Save = gcore.ORMSave(gFactory.clsClientType, Me)
End Function

Public Property Let ID(ByVal value As Variant):    mProperties.value("ID") = value: End Property
Public Property Get ID() As Variant:    ID = mProperties.value("ID"): End Property

Public Property Let Name(ByVal value As Variant): mProperties.value("Name") = value: End Property
Public Property Get Name() As Variant: Name = mProperties.value("Name"): End Property

Public Property Let Age(ByVal value As Variant): mProperties.value("Age") = value: End Property
Public Property Get Age() As Variant: Age = mProperties.value("Age"): End Property

Public Property Let Email(ByVal value As Variant): mProperties.value("Email") = value: End Property
Public Property Get Email() As Variant: Email = mProperties.value("Email"): End Property


