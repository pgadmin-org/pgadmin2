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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objAggregate As pgAggregate

  For Each objAggregate In mCol
    If objAggregate.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objAggregate
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szOwner As String, dInputType As Double, szStateFunction As String, dStateType As Double, szFinalFunction As String, dFinalType As Double, szInitialCondition As String) As pgAggregate
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & dInputType & ", " & QUOTE & szStateFunction & QUOTE & ", " & dStateType & ", " & QUOTE & szFinalFunction & QUOTE & ", " & dFinalType & ", " & QUOTE & szInitialCondition & QUOTE & ")", etFullDebug

Dim objNewMember As pgAggregate
Dim rs As New Recordset
Dim szSQL As String

  'Create a new object
  Set objNewMember = New pgAggregate
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.Oid = dOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  
  'Get the cached Type name. If an error occurs it means the
  'OID specified doesn't exist.
  On Error Resume Next
  objNewMember.InputType = objServer.Databases(Database).LookupType(dInputType)
  objNewMember.StateType = objServer.Databases(Database).LookupType(dStateType)
  objNewMember.FinalType = objServer.Databases(Database).LookupType(dFinalType)
  If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler

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
  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal InputType As String, ByVal StateFunction As String, ByVal StateType As String, Optional ByVal FinalFunction As String, Optional ByVal InitialCondition As String, Optional ByVal Comment As String) As pgAggregate
Attribute Add.VB_Description = "Add a new Aggregate."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & InputType & QUOTE & ", " & QUOTE & StateFunction & QUOTE & ", " & QUOTE & StateType & QUOTE & ", " & QUOTE & FinalFunction & QUOTE & ", " & QUOTE & InitialCondition & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objAggregate As pgAggregate
  
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "CREATE AGGREGATE " & fmtID(szNamespace) & "." & fmtID(Name)
  Else
    szSQL = "CREATE AGGREGATE " & fmtID(Name)
  End If
  szSQL = szSQL & " (BASETYPE = " & InputType & ", "
  szSQL = szSQL & "SFUNC = " & StateFunction & ", "
  szSQL = szSQL & "STYPE = " & StateType
  If FinalFunction <> "" Then szSQL = szSQL & ", FINALFUNC = " & FinalFunction
  If InitialCondition <> "" Then szSQL = szSQL & ", INITCOND = " & InitialCondition
  szSQL = szSQL & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If InputType = "ANY" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "SELECT oid, proname AS aggname, pg_get_userbyid(proowner) AS owner, aggtransfn, aggfinalfn, proargtypes[0] AS aggbasetype, aggtranstype, prorettype AS aggfinaltype, agginitval FROM pg_aggregate, pg_proc WHERE pg_proc.oid = pg_aggregate.aggfnoid AND aggname = '" & Name & "' AND aggbasetype = 0 AND aggnamespace = " & dNamespaceOID & "::oid"
    Else
      szSQL = "SELECT pg_aggregate.oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate WHERE aggname = '" & dbSZ(Name) & "' AND aggbasetype = 0"
    End If
  Else
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "SELECT oid, proname AS aggname, pg_get_userbyid(proowner) AS owner, aggtransfn, aggfinalfn, proargtypes[0] AS aggbasetype, aggtranstype, prorettype AS aggfinaltype, agginitval FROM pg_aggregate, pg_proc, pg_type WHERE pg_proc.oid = pg_aggregate.aggfnoid AND proargtypes[0] = pg_type.oid AND aggname = '" & Name & "' AND typname = '" & InputType & "' AND aggnamespace = " & dNamespaceOID & "::oid"
    Else
      szSQL = "SELECT pg_aggregate.oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate, pg_type WHERE aggbasetype = pg_type.oid AND aggname = '" & Name & "' AND typname = '" & InputType & "'"
    End If
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Comment <> "" Then
    If InputType = "ANY" Then
      If objVersion.VersionNum >= 7.3 Then
        szSQL = "COMMENT ON AGGREGATE " & fmtID(szNamespace) & "." & fmtID(Name) & " * IS '" & dbSZ(Comment) & "'"
      Else
        szSQL = "COMMENT ON AGGREGATE " & fmtID(Name) & " * IS '" & dbSZ(Comment) & "'"
      End If
      objServer.ExecSQL szSQL, cnDatabase
    Else
      If objVersion.VersionNum >= 7.3 Then
        szSQL = "COMMENT ON AGGREGATE " & fmtID(szNamespace) & "." & fmtID(Name) & " " & InputType & " IS '" & dbSZ(Comment) & "'"
      Else
        szSQL = "COMMENT ON AGGREGATE " & fmtID(Name) & " " & InputType & " IS '" & dbSZ(Comment) & "'"
      End If
      objServer.ExecSQL szSQL, cnDatabase
    End If
    objServer.Databases(Database).KillCommentCache
  End If
  
  If Not rs.EOF Then
    Set objAggregate = iAdd(rs.Fields("oid"), rs!aggname & "", rs!Owner & "", Val(rs!aggbasetype & ""), rs!aggtransfn & "", Val(rs!aggtranstype & ""), rs!aggfinalfn & "", Val(rs!aggfinaltype & ""), rs!agginitval & "")
  End If
  
  Set Add = objAggregate
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgAggregate
Attribute Item.VB_Description = "Returns a Aggregate object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Aggregates in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
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
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Aggregate."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Aggregate.
  szSQL = "DROP AGGREGATE " & mCol(ULEncode(Key)).FormattedID

  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
  Exit Sub
Err_Handler:
  'Error 5 is thrown if the item is not in the collection...
  If Err.Number = 5 Then Exit Sub
  objServer.iLogError Err.Number, Err.Description
End Sub

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Aggregates."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_AGGREGATES7_3 & " AND pronamespace = " & dNamespaceOID & "::oid ORDER BY aggname", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_AGGREGATES7_1 & " ORDER BY aggname", cnDatabase)
    End If
    While Not rs.EOF
      iAdd rs.Fields("oid"), rs!aggname & "", rs!Owner & "", Val(rs!aggbasetype & ""), rs!aggtransfn & "", Val(rs!aggtranstype & ""), rs!aggfinalfn & "", Val(rs!aggfinaltype & ""), rs!agginitval & ""
      rs.MoveNext
    Wend
  End If
  
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Aggregates.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objAggregate As pgAggregate
  
  szSQL = "-- Aggregates" & vbCrLf & vbCrLf

  For Each objAggregate In mCol
    If IncludeSystem Or (Not objAggregate.SystemObject) Then szSQL = szSQL & objAggregate.SQL & vbCrLf
  Next objAggregate
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
