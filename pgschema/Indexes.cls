VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Indexes"
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

Dim lTableOID As Long
Dim szTable As String

'Rollback to a prior version
Public Sub Rollback(Identifier As String, Version As Long)
Attribute Rollback.VB_Description = "Rollback the specified object to the specified version."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objIndex As pgIndex

  For Each objIndex In mCol
    If objIndex.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objIndex
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dTimeStamp As Date, lOID As Long, szName As String, bUnique As Boolean, bPrimary As Boolean, szDefinition As String) As pgIndex
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & bUnique & ", " & bPrimary & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgIndex

  'Create a new object
  Set objNewMember = New pgIndex
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.iName = szName
  objNewMember.Table = szTable
  objNewMember.Unique = bUnique
  objNewMember.Primary = bPrimary
  objNewMember.Definition = szDefinition
  objNewMember.Identifier = szName
  objNewMember.TableOID = lTableOID
  Set objNewMember.Connection = cnDatabase
  
  'Indexes are a pain in the butt because of the wierd way in which we must
  'define whether or not they are a system object. Baiscally we must create
  'pgIndex, and then ask it if it is a system object before adding it to
  'the collection
  
  If objServer.IncludeSys Then 'Show whatever...
    mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Else 'Only show non system...
    If Not objNewMember.SystemObject Then mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  End If
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal Unique As Boolean, ByVal ColumnList As String, Optional ByVal AccessMethod As String, Optional ByVal Comment As String, Optional ByVal Constraint As String) As pgIndex
Attribute Add.VB_Description = "Add a new Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Add(" & QUOTE & Name & QUOTE & ", " & Unique & ", " & QUOTE & ColumnList & QUOTE & ", " & QUOTE & Comment & QUOTE & ", " & QUOTE & Constraint & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objIndex As pgIndex
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE"
  If Unique Then szSQL = szSQL & " UNIQUE"
  szSQL = szSQL & " INDEX " & QUOTE & Name & QUOTE
  szSQL = szSQL & " ON " & QUOTE & szTable & QUOTE
  If AccessMethod <> "" Then szSQL = szSQL & " USING " & dbSZ(AccessMethod)
  szSQL = szSQL & " (" & dbSZ(ColumnList) & ")"
  'Constraint is only valid for PostgreSQL 7.2+
  If (Constraint <> "") And (objServer.dbVersion.VersionNum >= 7.2) Then
    szSQL = szSQL & " WHERE " & Constraint
  End If
  objServer.ExecSQL szSQL, cnDatabase
  
  If Comment <> "" Then
    szSQL = "COMMENT ON INDEX " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_INDEXES & " AND relname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objIndex = iAdd(rs!ts, rs!OID, rs!relname & "", ToBool(rs!indisunique), ToBool(rs!indisprimary), rs!Definition & "")
  End If
  
  If Not objIndex.SystemObject Then objIndex.Commit rcAdd, "Created object."
  Set Add = objIndex
  
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
Attribute Rename.VB_Description = "Rename the specified index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objIndex As New pgIndex
  Dim szSQL As String
  Dim rs As New Recordset
    
  If Not mCol(ULEncode(OldName)).SystemObject Then mCol(ULEncode(OldName)).Commit rcDelete, "Deleted object (prior to rename)."
  Set objIndex = mCol(ULEncode(OldName))
  objIndex.Name = NewName
  objIndex.Identifier = NewName

  'Select definition back from the DB...
  szSQL = "SELECT pg_get_indexdef(" & objIndex.OID & ") AS definition"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    objIndex.Definition = rs!Definition & ""
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  mCol.Add objIndex, ULEncode(NewName), ULEncode(OldName)
  mCol.Remove ULEncode(OldName)
  If Not mCol(ULEncode(NewName)).SystemObject Then mCol(ULEncode(NewName)).Commit rcAdd, "Created object (following rename)."
  
  Exit Sub
Err_Handler:
  objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgIndex
Attribute Item.VB_Description = "Returns an Index object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Indexes on the current Table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objIndex As pgIndex
Dim lCount As Long

  If IgnoreSys Then
    For Each objIndex In mCol
      If Not objIndex.SystemObject Then lCount = lCount + 1
    Next objIndex
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Rule.
  szSQL = "DROP INDEX " & QUOTE & mCol(ULEncode(Key)).Name & QUOTE
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
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Indexes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Refresh()", etFullDebug

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

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_INDEXES & " AND indrelid = " & lTableOID & " ORDER BY relname", cnDatabase)
    While Not rs.EOF
      iAdd rs!ts, rs!OID, rs!relname & "", ToBool(rs!indisunique), ToBool(rs!indisprimary), rs!Definition & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objIndex As pgIndex
  
  szSQL = "-- Indexes" & vbCrLf & vbCrLf

  For Each objIndex In mCol
    If IncludeSystem Or (Not objIndex.SystemObject) Then szSQL = szSQL & objIndex.SQL & vbCrLf
  Next objIndex
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all changes to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objIndex As pgIndex

  For Each objIndex In mCol
    objIndex.Commit UpdateType, UpdateComment
  Next objIndex
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

