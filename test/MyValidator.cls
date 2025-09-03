VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IORMValidator

Public Function IORMValidator_Validate(ValidateFunction As String, value As Variant) As Boolean
    IORMValidator_Validate = CallByName(Me, ValidateFunction, VbMethod, value)
End Function


Public Property Get ValidatorAge(): ValidatorAge = "ValidatorAgeFunction": End Property

Public Function ValidatorAgeFunction(value As Variant) As Boolean
    ValidatorAgeFunction = (value >= 40)
End Function


