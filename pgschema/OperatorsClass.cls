VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatorsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

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
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let NameSpace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NameSpace() As String
Attribute NameSpace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get Namespace()", etFullDebug

  NameSpace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objOpClass As pgOperatorClass

  For Each objOpClass In mCol
    If objOpClass.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objOpClass
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szOwner As String, szAccessMethod As String, szInputType As String, bDefault As Boolean) As pgOperatorClass
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szAccessMethod & QUOTE & ", " & QUOTE & szInputType & QUOTE & ", " & bDefault & ")", etFullDebug

Dim objNewMember As pgOperatorClass
Dim rs As New Recordset
Dim szSQL As String

  'Create a new object
  Set objNewMember = New pgOperatorClass
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  With objNewMember
    .Oid = dOID
    .Name = szName
    .Owner = szOwner
    .NamespaceOID = dNamespaceOID
    .NameSpace = szNamespace
    .AccessMethod = szAccessMethod
    .InputType = szInputType
    .Default = bDefault
    Set .Connection = cnDatabase
    .Identifier = szName & "(" & szAccessMethod & ")"
  End With
  
  'Add the Operator
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal AccessMethod As String, ByVal InputType As String, ByVal Default As Boolean, ByVal DataOps As OpClassOps, ByVal DataFncs As OpClassFncs) As pgOperatorClass
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & AccessMethod & QUOTE & ", " & QUOTE & InputType & QUOTE & ", " & Default & "," & DataOps.Count & "," & DataFncs.Count & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objOperatorClass As pgOperatorClass
Dim objOpClassOp As OpClassOp
Dim objOpClassFnc As OpClassFnc
Dim ii As Integer

  'Generate the szSQL query to create the new object and execute
  szSQL = szSQL & "CREATE OPERATOR CLASS " & fmtID(szNamespace) & "." & Name
  If Default Then szSQL = szSQL & "    DEFAULT "
  szSQL = szSQL & " FOR TYPE " & InputType
  szSQL = szSQL & " USING " & AccessMethod
  szSQL = szSQL & " AS " & vbCrLf

  'operator
  If DataOps.Count > 0 Then
    ii = 0
    For Each objOpClassOp In DataOps
      ii = ii + 1
      szSQL = szSQL & "        OPERATOR  "
      szSQL = szSQL & objOpClassOp.StrategyNumber
      szSQL = szSQL & "  " & objOpClassOp.Operator
      If objOpClassOp.Rechecked Then szSQL = szSQL & " RECHECK "
      If ii < DataOps.Count Then szSQL = szSQL & ","
    Next
  End If
    
  'function
  If DataFncs.Count > 0 Then
    If DataOps.Count > 0 Then szSQL = szSQL & ","
    ii = 0
    For Each objOpClassFnc In DataFncs
      ii = ii + 1
      szSQL = szSQL & "        FUNCTION  "
      szSQL = szSQL & objOpClassFnc.ProcedureIndex
      szSQL = szSQL & "  " & objOpClassFnc.Procedure
      If ii < DataFncs.Count Then szSQL = szSQL & ","
    Next
  End If
  objServer.ExecSQL szSQL, cnDatabase

  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_OPERATORCLASS & ",pg_am AS am WHERE opc.opcname = '" & dbSZ(Name) & "'"
  szSQL = szSQL & " AND am.oid=opc.opcamid AND opc.opcnamespace = " & dNamespaceOID & "::oid"
  szSQL = szSQL & " AND am.amname='" & AccessMethod & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    Set objOperatorClass = iAdd(rs.Fields("oid"), rs!opcname & "", rs!opclassowner & "", rs!opcindexaccessmethod, rs!opcinputtype, ToBool(rs!opcdefault))
  End If

  Set Add = objOperatorClass

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
Public Property Get Item(Key As String) As pgOperatorClass
Attribute Item.VB_Description = "Returns a Operator object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the OperatorsClass in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objOpClass As pgOperatorClass
Dim lCount As Long

  If IgnoreSys Then
    For Each objOpClass In mCol
      If Not objOpClass.SystemObject Then lCount = lCount + 1
    Next objOpClass
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
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  szSQL = "DROP OPERATOR CLASS " & mCol(ULEncode(Key)).FormattedID
  szSQL = szSQL & " USING " & mCol(ULEncode(Key)).AccessMethod
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
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of OperatorsClass."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Refresh()", etFullDebug

  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_OPERATORCLASS & " WHERE opcnamespace = " & dNamespaceOID & "::oid ORDER BY opcname", cnDatabase)
    While Not rs.EOF
      iAdd rs.Fields("oid"), rs!opcname & "", rs!opclassowner & "", rs!opcindexaccessmethod, rs!opcinputtype, ToBool(rs!opcdefault)
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
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":OperatorsClass.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objOperator As pgOperator
  
  szSQL = "-- OperatorsClass" & vbCrLf & vbCrLf

  For Each objOperator In mCol
    If IncludeSystem Or (Not objOperator.SystemObject) Then szSQL = szSQL & objOperator.SQL & vbCrLf
  Next objOperator
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
