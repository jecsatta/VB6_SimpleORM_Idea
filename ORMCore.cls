VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORMCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize(connString As String, dbType As String, factory As IORMFactory, Optional validator As IORMValidator)
    g_ConnectionString = connString
    g_DatabaseType = dbType

    Set g_Connection = CreateObject("ADODB.Connection")
    g_Connection.Open connString
    If Not validator Is Nothing Then Set g_Validator = validator
    Set g_factory = factory
    
End Sub


Public Function QueryParams(ClassName As String, params As ORMParams, Optional intLimit As Integer = 0, Optional strOrderBy As String = "") As Variant
    Dim strSQL As String
    Dim tableName As String
    Dim i As Integer
    Dim lastParamValue As Variant
    Dim objReturn() As Variant
    Dim objClass As IORMEntity
    
    Set objClass = g_factory.GetCreate(ClassName)
    
     tableName = objClass.GetTableName()
    strSQL = "SELECT * FROM " & tableName & " WHERE "

    params.First
    Do Until params.EOF
        strSQL = strSQL & params.Name & " " & params.Operator & " '" & params.Value & "' AND "
        params.MoveNext
    Loop
    strSQL = Mid(strSQL, 1, Len(strSQL) - 4)


    If Not strOrderBy = "" And InStr(1, UCase(strOrderBy), "ORDER BY") > 1 Then
        strSQL = strSQL & strOrderBy
    End If

    
    If intLimit > 0 Then
        strSQL = strSQL & " LIMIT " & intLimit
    End If

    
   QueryParams = QuerySQL(ClassName, strSQL)
End Function

Public Function QueryProps(ClassName As String, strSQL As String) As Variant
    Dim rs As Object
 
    Dim objORM As IORMEntity
    Dim objReturn() As Variant
    Dim field As Object
    Dim i As Long
    Dim prop2 As Variant
    Dim prop As ORMProperty
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, g_Connection
 
    objReturn = Array()
    i = 0
    Do Until rs.EOF
        Set objORM = g_factory.GetCreate(ClassName)
        ReDim Preserve objReturn(i) As Variant
        
        For Each prop2 In objORM.Props.GetProperties
            Set prop = prop2
            If Not prop.isComputed Then
                prop.Value = rs(prop.Name)
            End If
        Next prop2
        
        Set objReturn(i) = objORM
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    Set rs = Nothing
    
    QueryProps = objReturn
End Function

Public Function QuerySQL(ClassName As String, strSQL As String) As Variant
    Dim rs As Object
    Dim objClass As Object
    Dim objReturn() As Variant
    Dim field As Object
    Dim i As Long
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, g_Connection
 
    objReturn = Array()
     
    i = 0
    Do Until rs.EOF
        Set objClass = g_factory.GetCreate(ClassName)
        
        ReDim Preserve objReturn(i) As Variant
        For Each field In rs.Fields
            Dim propName As String
            propName = field.Name
            If HasProperty(objClass, UCase(propName)) Then
                CallByName objClass, propName, VbLet, field.Value
            End If
        Next field
        
        Set objReturn(i) = objClass
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    QuerySQL = objReturn
End Function

Public Function ORMInsert(ClassName As String, objData As IORMEntity) As Boolean
    Dim strSQL As String
    Dim tableName As String
    Dim prop As ORMProperty
    Dim strFields As String
    Dim strValues As String
    Dim prop2 As Variant
    
    tableName = objData.GetTableName()
    
    For Each prop2 In objData.Props.GetProperties
        Set prop = prop2
        If Not prop.isPrimaryKey Then
            strFields = strFields & prop.Name & ","
            strValues = strValues & "'" & prop.Value & "',"
        End If
    Next prop2
    
    strFields = Left(strFields, Len(strFields) - 1)
    strValues = Left(strValues, Len(strValues) - 1)
    
    strSQL = "INSERT INTO " & tableName & " (" & strFields & ") VALUES (" & strValues & ")"

    On Error Resume Next
    g_Connection.execute strSQL
    If Err.Number = 0 Then
        ORMInsert = True
    Else
        Debug.Print "Error inserting data: " & Err.Description
        ORMInsert = False
    End If
    On Error GoTo 0
End Function


Public Function ORMUpdate(ClassName As String, objData As IORMEntity) As Boolean
    Dim strSQL As String
    Dim tableName As String
    Dim prop As ORMProperty
    Dim strUpdateFields As String
    Dim primaryKeyField As String
    Dim primaryKeyValue As Variant
    Dim prop2 As Variant
    
    tableName = objData.GetTableName()
    
    For Each prop2 In objData.Props.GetProperties
        Set prop = prop2
        If prop.isPrimaryKey Then
            primaryKeyField = prop.Name
            primaryKeyValue = prop.Value
        Else
            strUpdateFields = strUpdateFields & prop.Name & " = '" & prop.Value & "',"
        End If
    Next prop2
    
    strUpdateFields = Left(strUpdateFields, Len(strUpdateFields) - 1)

    strSQL = "UPDATE " & tableName & " SET " & strUpdateFields & " WHERE " & primaryKeyField & " = '" & primaryKeyValue & "'"

    On Error Resume Next
    g_Connection.execute strSQL
    If Err.Number = 0 Then
        ORMUpdate = True
    Else
        Debug.Print "Error updating data: " & Err.Description
        ORMUpdate = False
    End If
    On Error GoTo 0
End Function


Public Function ORMSave(ClassName As String, objData As IORMEntity) As Boolean
    Dim primaryKeyField As String
    Dim primaryKeyValue As Variant
    Dim prop As ORMProperty
    Dim prop2 As Variant

    For Each prop2 In objData.Props.GetProperties
        Set prop = prop2
        
        If prop.isPrimaryKey Then
            primaryKeyField = prop.Name
            primaryKeyValue = prop.Value
            Exit For
        End If
    Next prop2

    If IsEmpty(primaryKeyValue) Or IsNull(primaryKeyValue) Then
        ORMSave = ORMInsert(ClassName, objData)
    Else
        ORMSave = ORMUpdate(ClassName, objData)
    End If
End Function

Public Function ExecuteNonQuery(sql As String) As Boolean
    g_Connection.execute sql
End Function
