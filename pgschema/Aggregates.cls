VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Aggregates"
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
Attribute Rollback.VB_Description = "Rollback the specified object to the specified version."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objAggregate As pgAggregate

  For Each objAggregate In mCol
    If objAggregate.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objAggregate
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dTimeStamp As Date, lOID As Long, szName As String, szOwner As String, lInputType As Long, szStateFunction As String, lStateType As Long, szFinalFunction As String, lFinalType As Long, szInitialCondition As String) As pgAggregate
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & lInputType & ", " & QUOTE & szStateFunction & QUOTE & ", " & lStateType & ", " & QUOTE & szFinalFunction & QUOTE & ", " & lFinalType & ", " & QUOTE & szInitialCondition & QUOTE & ")", etFullDebug

Dim objNewMember As pgAggregate
Dim rs As New Recordset
Dim szSQL As String

  'Create a new object
  Set objNewMember = New pgAggregate
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  
  'Get the cached Type name. If an error occurs it means the
  'OID specified doesn't exist.
  On Error Resume Next
  objNewMember.InputType = objServer.Databases(Database).LookupType(lInputType)
  objNewMember.StateType = objServer.Databases(Database).LookupType(lStateType)
  objNewMember.FinalType = objServer.Databases(Database).LookupType(lFinalType)
  On Error GoTo Err_Handler

  objNewMember.StateFunction = szStateFunction
  objNewMember.FinalFunction = szFinalFunction
  objNewMember.InitialCondition = szInitialCondition
  
  Set objNewMember.Connection = cnDatabase
  objNewMember.Identifier = szName & " " & objNewMember.InputType
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal InputType As String, ByVal StateFunction As String, ByVal StateType As String, Optional ByVal FinalFunction As String, Optional ByVal InitialCondition As String, Optional ByVal Comment As String) As pgAggregate
Attribute Add.VB_Description = "Add a new Aggregate."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & InputType & QUOTE & ", " & QUOTE & StateFunction & QUOTE & ", " & QUOTE & StateType & QUOTE & ", " & QUOTE & FinalFunction & QUOTE & ", " & QUOTE & InitialCondition & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objAggregate As pgAggregate
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE AGGREGATE " & QUOTE & Name & QUOTE
  szSQL = szSQL & " (BASETYPE = " & QUOTE & InputType & QUOTE & ", "
  szSQL = szSQL & "SFUNC = " & QUOTE & StateFunction & QUOTE & ", "
  szSQL = szSQL & "STYPE = " & QUOTE & StateType & QUOTE
  If FinalFunction <> "" Then szSQL = szSQL & ", FINALFUNC = " & QUOTE & FinalFunction & QUOTE
  If InitialCondition <> "" Then szSQL = szSQL & ", INITCOND = " & QUOTE & InitialCondition & QUOTE
  szSQL = szSQL & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If InputType = "ANY" Then
    szSQL = "SELECT now() AS ts, pg_aggregate.oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate WHERE aggname = '" & dbSZ(Name) & "' AND aggbasetype = 0"
  Else
    szSQL = "SELECT now() AS ts, pg_aggregate.oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate, pg_type WHERE aggbasetype = pg_type.oid AND aggname = '" & dbSZ(Name) & "' AND typname = '" & dbSZ(InputType) & "'"
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Comment <> "" Then
    If InputType = "ANY" Then
      szSQL = "COMMENT ON AGGREGATE " & QUOTE & Name & QUOTE & " * IS '" & dbSZ(Comment) & "'"
      objServer.ExecSQL szSQL, cnDatabase
    Else
      szSQL = "COMMENT ON AGGREGATE " & QUOTE & Name & QUOTE & " " & QUOTE & InputType & QUOTE & " IS '" & dbSZ(Comment) & "'"
      objServer.ExecSQL szSQL, cnDatabase
    End If
    objServer.Databases(Database).KillCommentCache
  End If
  
  If Not rs.EOF Then
    Set objAggregate = iAdd(rs!ts, rs!OID, rs!aggname & "", rs!Owner & "", Val(rs!aggbasetype & ""), rs!aggtransfn & "", Val(rs!aggtranstype & ""), rs!aggfinalfn & "", Val(rs!aggfinaltype & ""), rs!agginitval & "")
  End If
  
  If Not objAggregate.SystemObject Then objAggregate.Commit rcAdd, "Created object."
  Set Add = objAggregate
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  'If rs.State <> adStateClosed Then rs.Close
  'Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgAggregate
Attribute Item.VB_Description = "Returns a Aggregate object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Aggregates in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objAggregate As pgAggregate
Dim lCount As Long

  If IgnoreSys Then
    For Each objAggregate In mCol
      If Not objAggregate.SystemObject Then lCount = lCount + 1
    Next objAggregate
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Aggregate."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Aggregate.
  If mCol(ULEncode(Key)).InputType = "opaque" Then
    szSQL = "DROP AGGREGATE " & QUOTE & mCol(ULEncode(Key)).Name & QUOTE & " *"
  Else
    szSQL = "DROP AGGREGATE " & QUOTE & mCol(ULEncode(Key)).Name & QUOTE & " " & QUOTE & mCol(ULEncode(Key)).InputType & QUOTE
  End If
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
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Aggregates."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_AGGREGATES & " ORDER BY aggname", cnDatabase)
    While Not rs.EOF
      iAdd rs!ts, rs!OID, rs!aggname & "", rs!Owner & "", Val(rs!aggbasetype & ""), rs!aggtransfn & "", Val(rs!aggtranstype & ""), rs!aggfinalfn & "", Val(rs!aggfinaltype & ""), rs!agginitval & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objAggregate As pgAggregate
  
  szSQL = "-- Aggregates" & vbCrLf & vbCrLf

  For Each objAggregate In mCol
    If IncludeSystem Or (Not objAggregate.SystemObject) Then szSQL = szSQL & objAggregate.SQL & vbCrLf
  Next objAggregate
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all changes to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objAggregate As pgAggregate

  For Each objAggregate In mCol
    If Not objAggregate.SystemObject Then objAggregate.Commit UpdateType, UpdateComment
  Next objAggregate
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
