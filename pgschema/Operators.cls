VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Operators"
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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objOperator As pgOperator

  For Each objOperator In mCol
    If objOperator.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objOperator
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dTimeStamp As Date, lOID As Long, szName As String, szOwner As String, szKind As String, bHashJoins As Boolean, lLeftOperandType As Long, lRightOperandType As Long, lResultType As Long, lCommutator As Long, lNegator As Long, lLeftTypeSortOperator As Long, lRightTypeSortOperator As Long, szOperatorFunction As String, szRestrictFunction As String, szJoinFunction As String) As pgOperator
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szKind & QUOTE & ", " & bHashJoins & ", " & lLeftOperandType & ", " & lRightOperandType & ", " & lResultType & ", " & lCommutator & ", " & lNegator & ", " & lLeftTypeSortOperator & ", " & lRightTypeSortOperator & ", " & QUOTE & szOperatorFunction & QUOTE & ", " & QUOTE & szRestrictFunction & QUOTE & ", " & QUOTE & szJoinFunction & QUOTE & ")", etFullDebug

Dim objNewMember As pgOperator
Dim rs As New Recordset
Dim szSQL As String
Dim szLOT As String
Dim szROT As String
Dim szTemp As String

  'Create a new object
  Set objNewMember = New pgOperator
  If lOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  
  Select Case UCase(szKind)
    Case "B"
      objNewMember.Kind = "Infix"
    Case "L"
      objNewMember.Kind = "Prefix"
    Case "R"
      objNewMember.Kind = "Postfix"
  End Select

  objNewMember.HashJoins = bHashJoins
  szLOT = objServer.Databases(Database).LookupType(lLeftOperandType)
  If szLOT = "opaque" Then szLOT = "NONE"
  objNewMember.LeftOperandType = szLOT
  szROT = objServer.Databases(Database).LookupType(lRightOperandType)
  If szROT = "opaque" Then szROT = "NONE"
  objNewMember.RightOperandType = szROT
  szTemp = objServer.Databases(Database).LookupType(lResultType)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.ResultType = szTemp
  objNewMember.Commutator = objServer.Databases(Database).LookupOperator(lCommutator)
  objNewMember.Negator = objServer.Databases(Database).LookupOperator(lNegator)
  szTemp = objServer.Databases(Database).LookupOperator(lLeftTypeSortOperator)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.LeftTypeSortOperator = szTemp
  szTemp = objServer.Databases(Database).LookupOperator(lRightTypeSortOperator)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.RightTypeSortOperator = szTemp
  objNewMember.OperatorFunction = szOperatorFunction
  objNewMember.RestrictFunction = szRestrictFunction
  objNewMember.JoinFunction = szJoinFunction
  Set objNewMember.Connection = cnDatabase
  
  'Add the Operator
  objNewMember.Identifier = szName & " (" & szLOT & ", " & szROT & ")"
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
  
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal OperatorFunction As String, Optional ByVal LeftOperandType As String, Optional ByVal RightOperandType As String, Optional ByVal Commutator As String, Optional ByVal Negator As String, Optional ByVal RestrictFunction As String, Optional ByVal JoinFunction As String, Optional ByVal Hashes As Boolean, Optional ByVal LeftTypeSortOperator As String, Optional ByVal RightTypeSortOperator As String, Optional ByVal Comment As String) As pgOperator
Attribute Add.VB_Description = "Add a new Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & OperatorFunction & QUOTE & ", " & QUOTE & LeftOperandType & QUOTE & ", " & QUOTE & RightOperandType & QUOTE & ", " & QUOTE & Commutator & QUOTE & ", " & QUOTE & Negator & QUOTE & ", " & QUOTE & RestrictFunction & QUOTE & ", " & QUOTE & JoinFunction & QUOTE & ", " & Hashes & ", " & QUOTE & LeftTypeSortOperator & QUOTE & ", " & QUOTE & RightTypeSortOperator & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim szLOT As String
Dim szROT As String
Dim objOperator As pgOperator
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE OPERATOR " & Name
  szSQL = szSQL & " (PROCEDURE = " & QUOTE & OperatorFunction & QUOTE
  If LeftOperandType <> "" Then szSQL = szSQL & ", LEFTARG = " & QUOTE & LeftOperandType & QUOTE
  If RightOperandType <> "" Then szSQL = szSQL & ", RIGHTARG = " & QUOTE & RightOperandType & QUOTE
  If Commutator <> "" Then szSQL = szSQL & ", COMMUTATOR = " & QUOTE & Commutator & QUOTE
  If Negator <> "" Then szSQL = szSQL & ", NEGATOR = " & QUOTE & Negator & QUOTE
  If RestrictFunction <> "" Then szSQL = szSQL & ", RESTRICT = " & QUOTE & RestrictFunction & QUOTE
  If JoinFunction <> "" Then szSQL = szSQL & ", JOIN = " & QUOTE & JoinFunction & QUOTE
  If Hashes Then szSQL = szSQL & ", HASHES"
  If LeftTypeSortOperator <> "" Then szSQL = szSQL & ", SORT1 = " & QUOTE & LeftTypeSortOperator & QUOTE
  If RightTypeSortOperator <> "" Then szSQL = szSQL & ", SORT2 = " & QUOTE & RightTypeSortOperator & QUOTE
  szSQL = szSQL & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  
  szLOT = LeftOperandType
  If szLOT = "" Then szLOT = "NONE"
  szROT = RightOperandType
  If szROT = "" Then szROT = "NONE"
  If Comment <> "" Then
    szSQL = "COMMENT ON OPERATOR " & Name & "(" & szLOT & ", " & szROT & ") IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_OPERATORS & " AND oprname = '" & dbSZ(Name) & "' AND oprleft = '" & dbSZ(LeftOperandType) & "' AND oprright = '" & dbSZ(RightOperandType) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objOperator = iAdd(rs!ts, rs!OID, rs!oprname & "", rs!opowner & "", rs!oprkind & "", ToBool(rs!oprcanhash), rs!oprleft, rs!oprright, rs!oprresult, rs!oprcom, rs!oprnegate, rs!oprlsortop, rs!oprrsortop, rs!oprcode & "", rs!oprrest & "", rs!oprjoin & "")
  End If
    
  'Invalidate the Operator Cache...
  objServer.Databases(Database).KillOperatorCache
  
  If Not objOperator.SystemObject Then objOperator.Commit rcAdd, "Created object."
  Set Add = objOperator
  
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
Public Property Get Item(Key As String) As pgOperator
Attribute Item.VB_Description = "Returns a Operator object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Operators in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objOperator As pgOperator
Dim lCount As Long

  If IgnoreSys Then
    For Each objOperator In mCol
      If Not objOperator.SystemObject Then lCount = lCount + 1
    Next objOperator
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim szROT As String
Dim szLOT As String

  szLOT = mCol(ULEncode(Key)).LeftOperandType
  If szLOT = "" Then szLOT = "NONE"
  szROT = mCol(ULEncode(Key)).RightOperandType
  If szROT = "" Then szROT = "NONE"
  
  'Generate/Execute SQL to drop the Operator.
  szSQL = "DROP OPERATOR " & mCol(ULEncode(Key)).Name & "(" & szLOT & ", " & szROT & ")"
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rcAdd, "Deleted object."
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
  'Invalidate the Operator Cache...
  objServer.Databases(Database).KillOperatorCache
  
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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Operators."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_OPERATORS & " ORDER BY oprname", cnDatabase)
    While Not rs.EOF
      iAdd rs!ts, rs!OID, rs!oprname & "", rs!opowner & "", rs!oprkind & "", ToBool(rs!oprcanhash), rs!oprleft, rs!oprright, rs!oprresult, rs!oprcom, rs!oprnegate, rs!oprlsortop, rs!oprrsortop, rs!oprcode & "", rs!oprrest & "", rs!oprjoin & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objOperator As pgOperator
  
  szSQL = "-- Operators" & vbCrLf & vbCrLf

  For Each objOperator In mCol
    If IncludeSystem Or (Not objOperator.SystemObject) Then szSQL = szSQL & objOperator.SQL & vbCrLf
  Next objOperator
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all changes to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objOperator As pgOperator

  For Each objOperator In mCol
    If Not objOperator.SystemObject Then objOperator.Commit UpdateType, UpdateComment
  Next objOperator
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
