VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tables"
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

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(lOID As Long, szName As String, szOwner As String, szACL As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim objNewMember As pgTable
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgTable
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.iName = szName
  objNewMember.iOwner = szOwner
  objNewMember.ACL = szACL
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Language
  mCol.Add objNewMember, objNewMember.Identifier
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, ByVal Columns As String, Optional ByVal PrimaryKeys As String, Optional ByVal Checks As String, Optional ByVal ForeignKeys As String, Optional ByVal Inherits As String, Optional ByVal Comment As String) As pgTable
Attribute Add.VB_Description = "Add a new Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Columns & QUOTE & ", " & QUOTE & PrimaryKeys & QUOTE & ", " & QUOTE & Checks & QUOTE & ", " & QUOTE & Checks & QUOTE & ", " & QUOTE & ForeignKeys & QUOTE & ", " & QUOTE & Inherits & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE TABLE " & QUOTE & Name & QUOTE & " (" & Columns
  If PrimaryKeys <> "" Then szSQL = szSQL & ", PRIMARY KEY (" & PrimaryKeys & ")"
  If Checks <> "" Then szSQL = szSQL & ", " & Checks
  If ForeignKeys <> "" Then szSQL = szSQL & ", " & ForeignKeys
  szSQL = szSQL & ")"
  If Inherits <> "" Then szSQL = szSQL & " INHERITS (" & Inherits & ")"
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    szSQL = "COMMENT ON TABLE " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If

  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_TABLES & " AND relname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    iAdd rs!OID, rs!relname & "", rs!tableowner & "", rs!relacl & ""
  End If
  
  'Return the object created
  Set Add = mCol(Name)
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
  
End Function

'Rename
Public Sub Rename(ByVal OldName As String, ByVal NewName As String)
'On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objTable As New pgTable
  
  Set objTable = mCol(OldName)
  objTable.Name = NewName
  objTable.Identifier = NewName
  mCol.Add objTable, NewName, OldName
  mCol.Remove OldName
  
  Exit Sub
Err_Handler:
  objServer.iLogError Err
  
End Sub

'Return the specified Item
Public Property Get Item(IndexOrKey As Variant) As pgTable
Attribute Item.VB_Description = "Returns a Table object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Property Get Item(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

  Set Item = mCol(IndexOrKey)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Tables in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objTable As pgTable
Dim lCount As Long

  If IgnoreSys Then
    For Each objTable In mCol
      If Not objTable.SystemObject Then lCount = lCount + 1
    Next objTable
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal IndexOrKey As Variant)
Attribute Remove.VB_Description = "Remove the specified Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Remove(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP TABLE " & QUOTE & dbSZ(mCol(IndexOrKey).Name) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove IndexOrKey
  
  Exit Sub
Err_Handler:
  'Error 5 is thrown if the item is not in the collection...
  If Err.Number = 5 Then Exit Sub
  objServer.iLogError Err
End Sub

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_TABLES & " ORDER BY relname", cnDatabase)
    While Not rs.EOF
      iAdd rs!OID, rs!relname & "", rs!tableowner, rs!relacl & ""
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

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Tables."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Tables.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objTable As pgTable
  
  szSQL = "-- Tables" & vbCrLf & vbCrLf

  For Each objTable In mCol
    If IncludeSystem Or (Not objTable.SystemObject) Then szSQL = szSQL & objTable.SQL & vbCrLf
  Next objTable
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
