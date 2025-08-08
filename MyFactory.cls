VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IORMFactory

Public Function IORMFactory_GetCreate(ClassName As String) As Object
    Set IORMFactory_GetCreate = CallByName(Me, ClassName & "_Create", VbMethod)
End Function

Public Property Get clsClientType() As String: clsClientType = "clsClient": End Property
Public Function clsClient_Create() As Object:    Set clsClient_Create = New clsClient: End Function

Public Property Get clsEmployeeType() As String: clsEmployeeType = "clsEmployee": End Property
Public Function clsEmployee_Create() As Object:    Set clsEmployee_Create = New clsEmployee: End Function



