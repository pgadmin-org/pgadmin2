VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Views"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

'Rollback to a prior version
Public Sub Rollback(Identifier As String, Version As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

Dim szDefinition As String
Dim objItem As rcItem

  'Get the definition to restore to first as it will vanish as soon as we
  'drop the current version
  For Each objItem In mCol(ULEncode(Identifier)).History
    If objItem.Version = Version Then
      szDefinition = objItem.Definition
      Exit For
    End If
  Next objItem
  If szDefinition = "" Then Exit Sub
  
  'Remove the current version
  Remove Identifier
  
  'Recreate the new object.
  objServer.Databases(Database).Execute szDefinition
  
  'Kill the Comment Cache and log the change
  objServer.Databases(Database).KillCommentCache
  Refresh
  mCol(ULEncode(Identifier)).Commit rcAdd, "Object rolled back to version " & Version & "."
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objView As pgView

  For Each objView In mCol
    If objView.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objView
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(dTimeStamp As Date, lOID As Long, szName As String, szOwner As String, szACL As String, szDefinition As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgView
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgView
  If lOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.iName = szName
  objNewMember.Owner = szOwner
  objNewMember.ACL = szACL
  objNewMember.iDefinition = szDefinition
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Language
  mCol.Add objNewMember, ULEncode(szName)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, ByVal Definition As String, Optional ByVal Comment As String) As pgView
Attribute Add.VB_Description = "Add a new View."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Definition & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE VIEW " & QUOTE & Name & QUOTE & " AS " & Definition
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    szSQL = "COMMENT ON VIEW " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If

  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_VIEWS & " AND relname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    iAdd rs!ts, rs!OID, rs!relname & "", rs!viewowner & "", rs!relacl & "", rs!Definition & ""
  End If
  
  'Return the object created
  If Not mCol(ULEncode(Name)).SystemObject Then mCol(ULEncode(Name)).Commit rcAdd, "Created object."
  Set Add = mCol(ULEncode(Name))
  
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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objView As New pgView
    
  If Not mCol(ULEncode(OldName)).SystemObject Then mCol(ULEncode(OldName)).Commit rcDelete, "Deleted object (prior to rename)."
  Set objView = mCol(ULEncode(OldName))
  objView.Name = NewName
  objView.Identifier = NewName
  mCol.Add objView, ULEncode(NewName), ULEncode(OldName)
  mCol.Remove ULEncode(OldName)
  If Not mCol(ULEncode(NewName)).SystemObject Then mCol(ULEncode(NewName)).Commit rcAdd, "Created object (following rename)."
  
Exit Sub
Err_Handler:
objServer.iLogError Err
End Sub
'Return the specified Item
Public Property Get Item(Key As String) As pgView
Attribute Item.VB_Description = "Returns a View object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Views in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objView As pgView
Dim lCount As Long

  If IgnoreSys Then
    For Each objView In mCol
      If Not objView.SystemObject Then lCount = lCount + 1
    Next objView
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified View."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the View.
  szSQL = "DROP VIEW " & QUOTE & dbSZ(mCol(ULEncode(Key)).Name) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rcDelete, "Deleted object."
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
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
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_VIEWS & " ORDER BY relname", cnDatabase)
    While Not rs.EOF
      iAdd rs!ts, rs!OID, rs!relname & "", rs!viewowner, rs!relacl & "", rs!Definition & ""
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
Attribute Refresh.VB_Description = "Refresh the collection of Views."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objView As pgView
  
  szSQL = "-- Views" & vbCrLf & vbCrLf

  For Each objView In mCol
    If IncludeSystem Or (Not objView.SystemObject) Then szSQL = szSQL & objView.SQL & vbCrLf
  Next objView
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all changes to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Views.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objView As pgView

  For Each objView In mCol
    If Not objView.SystemObject Then objView.Commit UpdateType, UpdateComment
  Next objView
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
