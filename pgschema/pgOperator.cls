VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szKind As String
Private bHashJoins As Boolean
Private szLeftOperandType As String
Private szRightOperandType As String
Private szResultType As String
Private szCommutator As String
Private szNegator As String
Private szLeftTypeSortOperator As String
Private szRightTypeSortOperator As String
Private szOperatorFunction As String
Private szRestrictFunction As String
Private szJoinFunction As String
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get History()", etFullDebug

  If objServer.Databases(Database).RevisionControl Then
    If colHistory Is Nothing Then
      Set colHistory = New History
     
      'Pass the options to the class
      Set colHistory.Connection = cnDatabase
      colHistory.ObjIdentifier = szIdentifier
      colHistory.objType = ObjectType
      colHistory.Initialize
    End If
    Set History = colHistory
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TimeStamp
Friend Property Let TimeStamp(ByVal vData As Date)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Operator OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Operator name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get ObjectType()", etFullDebug

  ObjectType = "Operator"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Operator's Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Kind
Friend Property Let Kind(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Kind(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szKind = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Kind() As String
Attribute Kind.VB_Description = "Returns the Operators Kind (infix, prefix, postfix)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Kind()", etFullDebug

  Kind = szKind
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return HashJoins?
Friend Property Let HashJoins(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let HashJoins(" & vData & ")", etFullDebug

  bHashJoins = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get HashJoins() As Boolean
Attribute HashJoins.VB_Description = "Returns a value indicating whether the Operator supports Hash Joins."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get HashJoins()", etFullDebug

  HashJoins = bHashJoins
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the LeftOperandType
Friend Property Let LeftOperandType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let LeftOperandType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLeftOperandType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get LeftOperandType() As String
Attribute LeftOperandType.VB_Description = "Returns the Operator's left operand Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get LeftOperandType()", etFullDebug

  LeftOperandType = szLeftOperandType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the RightOperandType
Friend Property Let RightOperandType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RightOperandType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRightOperandType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get RightOperandType() As String
Attribute RightOperandType.VB_Description = "Returns the Operator's right operand Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RightOperandType()", etFullDebug

  RightOperandType = szRightOperandType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ResultType
Friend Property Let ResultType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let ResultType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szResultType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ResultType() As String
Attribute ResultType.VB_Description = "Returns the Operator's result Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get ResultType()", etFullDebug

  ResultType = szResultType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Commutator
Friend Property Let Commutator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Commutator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szCommutator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Commutator() As String
Attribute Commutator.VB_Description = "Returns the Operator's Commutator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Commutator()", etFullDebug

  Commutator = szCommutator
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Negator
Friend Property Let Negator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Negator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szNegator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Negator() As String
Attribute Negator.VB_Description = "Returns the Operator's negator Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Negator()", etFullDebug

  Negator = szNegator
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the LeftTypeSortOperator
Friend Property Let LeftTypeSortOperator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let LeftTypeSortOperator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLeftTypeSortOperator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get LeftTypeSortOperator() As String
Attribute LeftTypeSortOperator.VB_Description = "Returns the Operator's left type-sort Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get LeftTypeSortOperator()", etFullDebug

  LeftTypeSortOperator = szLeftTypeSortOperator
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the RightTypeSortOperator
Friend Property Let RightTypeSortOperator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RightTypeSortOperator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRightTypeSortOperator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get RightTypeSortOperator() As String
Attribute RightTypeSortOperator.VB_Description = "Returns the Operator's right type-sort Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RightTypeSortOperator()", etFullDebug

  RightTypeSortOperator = szRightTypeSortOperator
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OperatorFunction
Friend Property Let OperatorFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let OperatorFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOperatorFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OperatorFunction() As String
Attribute OperatorFunction.VB_Description = "Returns the Operator Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get OperatorFunction()", etFullDebug

  OperatorFunction = szOperatorFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the RestrictFunction
Friend Property Let RestrictFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RestrictFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRestrictFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get RestrictFunction() As String
Attribute RestrictFunction.VB_Description = "Returns the Operator's restrict function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RestrictFunction()", etFullDebug

  RestrictFunction = szRestrictFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the JoinFunction
Friend Property Let JoinFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let JoinFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szJoinFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get JoinFunction() As String
Attribute JoinFunction.VB_Description = "Returns the Operator's Join Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get JoinFunction()", etFullDebug

  JoinFunction = szJoinFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String
Dim szROT As String
Dim szLOT As String

  szLOT = szLeftOperandType
  If szLOT = "" Then szLOT = "NONE"
  szROT = szRightOperandType
  If szROT = "" Then szROT = "NONE"
  szSQL = "COMMENT ON OPERATOR " & szName & "(" & szLOT & ", " & szROT & ") IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Operator comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim szROT As String
Dim szLOT As String

  szLOT = szLeftOperandType
  If szLOT = "" Then szLOT = "NONE"
  szROT = szRightOperandType
  If szROT = "" Then szROT = "NONE"
  
  szSQL = "-- Operator: " & szIdentifier & vbCrLf & _
          "CREATE OPERATOR " & szName & " (" & _
          "PROCEDURE = " & QUOTE & szOperatorFunction & QUOTE
  
  If szLeftOperandType <> "" Then szSQL = szSQL & ", LEFTARG = " & QUOTE & szLeftOperandType & QUOTE
  If szRightOperandType <> "" Then szSQL = szSQL & ", RIGHTARG = " & QUOTE & szRightOperandType & QUOTE
  If szCommutator <> "" Then szSQL = szSQL & ", COMMUTATOR = " & szCommutator
  If szNegator <> "" Then szSQL = szSQL & ", NEGATOR = " & szNegator
  If szRestrictFunction <> "" Then szSQL = szSQL & ", RESTRICT = " & QUOTE & szRestrictFunction & QUOTE
  If szJoinFunction <> "" Then szSQL = szSQL & ", JOIN = " & QUOTE & szJoinFunction & QUOTE
  If bHashJoins Then szSQL = szSQL & ", HASHES"
  If szLeftTypeSortOperator <> "" Then szSQL = szSQL & ", SORT1 = " & szLeftTypeSortOperator
  If szRightTypeSortOperator <> "" Then szSQL = szSQL & ", SORT2 = " & szRightTypeSortOperator
  
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON OPERATOR " & szName & "(" & szLOT & ", " & szROT & ") IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.RCStatus()", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If (Not objServer.Databases(Database).RevisionControl) Or SystemObject Then
    RCStatus = rcNotApplicable
    Exit Property
  End If
  
  szSQL = "SELECT rc_timestamp, rc_definition FROM pgadmin_rclog WHERE rc_type = '" & dbSZ(ObjectType) & "' AND rc_identifier = '" & dbSZ(szIdentifier) & "' ORDER BY rc_version DESC LIMIT 1"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If rs.EOF Then
    RCStatus = rcNotInRC
  Else
    If SQL = rs!rc_definition & "" Then
      RCStatus = rcUpToDate
    Else
      If rs!rc_timestamp > dTimeStamp Then
        RCStatus = rcNeedsRefresh
      Else
        RCStatus = rcOutOfDate
      End If
    End If
  End If
  
  Set rs = Nothing
  
  Exit Property
Err_Handler:
  Set rs = Nothing
  objServer.iLogError Err
End Property


