VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Checks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

Dim lTableOID As Long
Dim szTable As String

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(lOID As Long, szName As String, szDefinition As String) As pgCheck
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgCheck

  'Create a new object
  Set objNewMember = New pgCheck
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Table = szTable
  objNewMember.Definition = szDefinition
  If szName <> "" Then
    objNewMember.Identifier = szName
  Else
    objNewMember.Identifier = "Unamed Check - " & lOID
  End If
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, objNewMember.Identifier
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal Definition As String) As pgCheck
Attribute Add.VB_Description = "Add a new Check."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Definition & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "ALTER TABLE " & QUOTE & szTable & QUOTE
  szSQL = szSQL & " ADD CONSTRAINT " & QUOTE & Name & QUOTE
  szSQL = szSQL & " CHECK (" & Definition & ")"
  objServer.ExecSQL szSQL, cnDatabase
    
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_CHECKS & " WHERE rcrelid = " & lTableOID & " AND rcname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set Add = iAdd(rs!OID, rs!rcname & "", rs!rcsrc & "")
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
  
End Function

'Return the specified Item
Public Property Get Item(IndexOrKey As Variant) As pgCheck
Attribute Item.VB_Description = "Returns a Check object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property Get Item(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

  Set Item = mCol(IndexOrKey)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Checks on the current Table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property Get Count()", etFullDebug

Dim objCheck As pgCheck
Dim lCount As Long

  If IgnoreSys Then
    For Each objCheck In mCol
      If Not objCheck.SystemObject Then lCount = lCount + 1
    Next objCheck
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Checks."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_CHECKS & " WHERE rcrelid = " & lTableOID & " ORDER BY rcname", cnDatabase)
    While Not rs.EOF
      iAdd rs!OID, rs!rcname & "", rs!rcsrc & ""
      rs.MoveNext
    Wend
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Checks.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
