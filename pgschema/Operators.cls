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
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objOperator As pgOperator

  For Each objOperator In mCol
    If objOperator.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objOperator
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szOwner As String, szKind As String, bHashJoins As Boolean, dLeftOperandType As Double, dRightOperandType As Double, dResultType As Double, dCommutator As Double, dNegator As Double, dLeftTypeSortOperator As Double, dRightTypeSortOperator As Double, szOperatorFunction As String, szRestrictFunction As String, szJoinFunction As String) As pgOperator
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szKind & QUOTE & ", " & bHashJoins & ", " & dLeftOperandType & ", " & dRightOperandType & ", " & dResultType & ", " & dCommutator & ", " & dNegator & ", " & dLeftTypeSortOperator & ", " & dRightTypeSortOperator & ", " & QUOTE & szOperatorFunction & QUOTE & ", " & QUOTE & szRestrictFunction & QUOTE & ", " & QUOTE & szJoinFunction & QUOTE & ")", etFullDebug

Dim objNewMember As pgOperator
Dim rs As New Recordset
Dim szSQL As String
Dim szLOT As String
Dim szROT As String
Dim szTemp As String


  'Create a new object
  Set objNewMember = New pgOperator
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
  
  Select Case UCase(szKind)
    Case "B"
      objNewMember.Kind = "Infix"
    Case "L"
      objNewMember.Kind = "Prefix"
    Case "R"
      objNewMember.Kind = "Postfix"
  End Select

  objNewMember.HashJoins = bHashJoins
  szLOT = objServer.Databases(Database).LookupType(dLeftOperandType)
  If szLOT = "opaque" Then szLOT = "NONE"
  objNewMember.LeftOperandType = szLOT
  szROT = objServer.Databases(Database).LookupType(dRightOperandType)
  If szROT = "opaque" Then szROT = "NONE"
  objNewMember.RightOperandType = szROT
  szTemp = objServer.Databases(Database).LookupType(dResultType)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.ResultType = szTemp
  objNewMember.Commutator = objServer.Databases(Database).LookupOperator(dCommutator)
  objNewMember.Negator = objServer.Databases(Database).LookupOperator(dNegator)
  szTemp = objServer.Databases(Database).LookupOperator(dLeftTypeSortOperator)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.LeftTypeSortOperator = szTemp
  szTemp = objServer.Databases(Database).LookupOperator(dRightTypeSortOperator)
  If szTemp = "opaque" Then szTemp = "NONE"
  objNewMember.RightTypeSortOperator = szTemp
  objNewMember.OperatorFunction = szOperatorFunction
  If szRestrictFunction = "-" Then szRestrictFunction = ""
  If szJoinFunction = "-" Then szJoinFunction = ""
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
  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal OperatorFunction As String, Optional ByVal LeftOperandType As String, Optional ByVal RightOperandType As String, Optional ByVal Commutator As String, Optional ByVal Negator As String, Optional ByVal RestrictFunction As String, Optional ByVal JoinFunction As String, Optional ByVal Hashes As Boolean, Optional ByVal LeftTypeSortOperator As String, Optional ByVal RightTypeSortOperator As String, Optional ByVal Comment As String) As pgOperator
Attribute Add.VB_Description = "Add a new Operator."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & OperatorFunction & QUOTE & ", " & QUOTE & LeftOperandType & QUOTE & ", " & QUOTE & RightOperandType & QUOTE & ", " & QUOTE & Commutator & QUOTE & ", " & QUOTE & Negator & QUOTE & ", " & QUOTE & RestrictFunction & QUOTE & ", " & QUOTE & JoinFunction & QUOTE & ", " & Hashes & ", " & QUOTE & LeftTypeSortOperator & QUOTE & ", " & QUOTE & RightTypeSortOperator & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim szLOT As String
Dim szROT As String
Dim objOperator As pgOperator
Dim lOidOpLeft As Long
Dim lOidOpRigth As Long
  
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "CREATE OPERATOR " & fmtID(szNamespace) & "." & Name
  Else
    szSQL = "CREATE OPERATOR " & Name
  End If
  szSQL = szSQL & " (PROCEDURE = " & OperatorFunction
  If LeftOperandType <> "" Then szSQL = szSQL & ", LEFTARG = " & LeftOperandType
  If RightOperandType <> "" Then szSQL = szSQL & ", RIGHTARG = " & RightOperandType
  If Commutator <> "" Then szSQL = szSQL & ", COMMUTATOR = " & Commutator
  If Negator <> "" Then szSQL = szSQL & ", NEGATOR = " & Negator
  If RestrictFunction <> "" Then szSQL = szSQL & ", RESTRICT = " & RestrictFunction
  If JoinFunction <> "" Then szSQL = szSQL & ", JOIN = " & JoinFunction
  If Hashes Then szSQL = szSQL & ", HASHES"
  If LeftTypeSortOperator <> "" Then szSQL = szSQL & ", SORT1 = " & LeftTypeSortOperator
  If RightTypeSortOperator <> "" Then szSQL = szSQL & ", SORT2 = " & RightTypeSortOperator
  szSQL = szSQL & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  szLOT = fmtID(LeftOperandType)
  If szLOT = "" Then szLOT = "NONE"
  szROT = fmtID(RightOperandType)
  If szROT = "" Then szROT = "NONE"
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON OPERATOR " & fmtID(szNamespace) & "." & Name & "(" & szLOT & ", " & szROT & ") IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON OPERATOR " & Name & "(" & szLOT & ", " & szROT & ") IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'get oid type
  With objServer.Databases(Database)
    If .Namespaces(szNamespace).Types.Exists(LeftOperandType) Then
      lOidOpLeft = .Namespaces(szNamespace).Types(LeftOperandType).Oid
    ElseIf .Namespaces("pg_catalog").Types.Exists(LeftOperandType) Then
      lOidOpLeft = .Namespaces("pg_catalog").Types(LeftOperandType).Oid
    End If
      
    If .Namespaces(szNamespace).Types.Exists(RightOperandType) Then
      lOidOpRigth = .Namespaces(szNamespace).Types(RightOperandType).Oid
    ElseIf .Namespaces("pg_catalog").Types.Exists(RightOperandType) Then
      lOidOpRigth = .Namespaces("pg_catalog").Types(RightOperandType).Oid
    End If
  End With
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_OPERATORS & " WHERE oprname = '" & dbSZ(Name) & "' AND oprleft = " & lOidOpLeft & "::oid AND oprright = " & lOidOpRigth & "::oid"
  If objVersion.VersionNum >= 7.3 Then szSQL = szSQL & " AND oprnamespace = " & dNamespaceOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objOperator = iAdd(rs.Fields("oid"), rs!oprname & "", rs!opowner & "", rs!oprkind & "", ToBool(rs!oprcanhash), rs!oprleft, rs!oprright, rs!oprresult, rs!oprcom, rs!oprnegate, rs!oprlsortop, rs!oprrsortop, rs!oprcode & "", rs!oprrest & "", rs!oprjoin & "")
  End If
    
  'Invalidate the Operator Cache...
  objServer.Databases(Database).KillOperatorCache
  
  Set Add = objOperator
  
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
Public Property Get Item(Key As String) As pgOperator
Attribute Item.VB_Description = "Returns a Operator object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Operators in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
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
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Operator."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  szSQL = "DROP OPERATOR " & mCol(ULEncode(Key)).FormattedID
  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
  'Invalidate the Operator Cache...
  objServer.Databases(Database).KillOperatorCache
  
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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Operators."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_OPERATORS & " WHERE oprnamespace = " & dNamespaceOID & "::oid ORDER BY oprname", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_OPERATORS & " ORDER BY oprname", cnDatabase)
    End If
    While Not rs.EOF
      iAdd rs.Fields("oid"), rs!oprname & "", rs!opowner & "", rs!oprkind & "", ToBool(rs!oprcanhash), rs!oprleft, rs!oprright, rs!oprresult, rs!oprcom, rs!oprnegate, rs!oprlsortop, rs!oprrsortop, rs!oprcode & "", rs!oprrest & "", rs!oprjoin & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Operators.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Operators.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objOperator As pgOperator
  
  szSQL = "-- Operators" & vbCrLf & vbCrLf

  For Each objOperator In mCol
    If IncludeSystem Or (Not objOperator.SystemObject) Then szSQL = szSQL & objOperator.SQL & vbCrLf
  Next objOperator
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
