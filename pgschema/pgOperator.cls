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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private dOID As Double
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
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get FormattedID()", etFullDebug
 
Dim szROT As String
Dim szLOT As String

  szLOT = fmtID(szLeftOperandType)
  If szLOT = "" Then szLOT = "NONE"
  szROT = fmtID(szRightOperandType)
  If szROT = "" Then szROT = "NONE"
  If objVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szNamespace) & "." & szName & "(" & szLOT & ", " & szROT & ")"
  Else
    FormattedID = szName & "(" & szLOT & ", " & szROT & ")"
  End If

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Tag()", etFullDebug

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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the Operator OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Operator name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get ObjectType()", etFullDebug

  ObjectType = "Operator"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Operator's Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Kind
Friend Property Let Kind(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Kind(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szKind = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Kind() As String
Attribute Kind.VB_Description = "Returns the Operators Kind (infix, prefix, postfix)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Kind()", etFullDebug

  Kind = szKind
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return HashJoins?
Friend Property Let HashJoins(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let HashJoins(" & vData & ")", etFullDebug

  bHashJoins = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get HashJoins() As Boolean
Attribute HashJoins.VB_Description = "Returns a value indicating whether the Operator supports Hash Joins."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get HashJoins()", etFullDebug

  HashJoins = bHashJoins
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the LeftOperandType
Friend Property Let LeftOperandType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let LeftOperandType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLeftOperandType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get LeftOperandType() As String
Attribute LeftOperandType.VB_Description = "Returns the Operator's left operand Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get LeftOperandType()", etFullDebug

  LeftOperandType = szLeftOperandType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the RightOperandType
Friend Property Let RightOperandType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RightOperandType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRightOperandType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get RightOperandType() As String
Attribute RightOperandType.VB_Description = "Returns the Operator's right operand Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RightOperandType()", etFullDebug

  RightOperandType = szRightOperandType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the ResultType
Friend Property Let ResultType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let ResultType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szResultType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ResultType() As String
Attribute ResultType.VB_Description = "Returns the Operator's result Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get ResultType()", etFullDebug

  ResultType = szResultType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Commutator
Friend Property Let Commutator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Commutator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szCommutator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Commutator() As String
Attribute Commutator.VB_Description = "Returns the Operator's Commutator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Commutator()", etFullDebug

  Commutator = szCommutator
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Negator
Friend Property Let Negator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Negator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szNegator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Negator() As String
Attribute Negator.VB_Description = "Returns the Operator's negator Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Negator()", etFullDebug

  Negator = szNegator
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the LeftTypeSortOperator
Friend Property Let LeftTypeSortOperator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let LeftTypeSortOperator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLeftTypeSortOperator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get LeftTypeSortOperator() As String
Attribute LeftTypeSortOperator.VB_Description = "Returns the Operator's left type-sort Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get LeftTypeSortOperator()", etFullDebug

  LeftTypeSortOperator = szLeftTypeSortOperator
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the RightTypeSortOperator
Friend Property Let RightTypeSortOperator(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RightTypeSortOperator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRightTypeSortOperator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get RightTypeSortOperator() As String
Attribute RightTypeSortOperator.VB_Description = "Returns the Operator's right type-sort Operator."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RightTypeSortOperator()", etFullDebug

  RightTypeSortOperator = szRightTypeSortOperator
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OperatorFunction
Friend Property Let OperatorFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let OperatorFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOperatorFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OperatorFunction() As String
Attribute OperatorFunction.VB_Description = "Returns the Operator Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get OperatorFunction()", etFullDebug

  OperatorFunction = szOperatorFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the RestrictFunction
Friend Property Let RestrictFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let RestrictFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szRestrictFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get RestrictFunction() As String
Attribute RestrictFunction.VB_Description = "Returns the Operator's restrict function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get RestrictFunction()", etFullDebug

  RestrictFunction = szRestrictFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the JoinFunction
Friend Property Let JoinFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let JoinFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szJoinFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get JoinFunction() As String
Attribute JoinFunction.VB_Description = "Returns the Operator's Join Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get JoinFunction()", etFullDebug

  JoinFunction = szJoinFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON OPERATOR " & FormattedID & " IS '" & dbSZ(vData) & "'"

  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Operator comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgOperator.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(dOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
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
  
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "-- Operator: " & FormattedID & vbCrLf & _
            "CREATE OPERATOR " & fmtID(szNamespace) & "." & szName & " (" & _
            "PROCEDURE = " & fmtID(szOperatorFunction)
  Else
    szSQL = "-- Operator: " & FormattedID & vbCrLf & _
            "CREATE OPERATOR " & szName & " (" & _
            "PROCEDURE = " & fmtID(szOperatorFunction)
  End If
  
  If szLeftOperandType <> "" Then szSQL = szSQL & ", LEFTARG = " & szLeftOperandType
  If szRightOperandType <> "" Then szSQL = szSQL & ", RIGHTARG = " & szRightOperandType
  If szCommutator <> "" Then szSQL = szSQL & ", COMMUTATOR = " & szCommutator
  If szNegator <> "" Then szSQL = szSQL & ", NEGATOR = " & szNegator
  If szRestrictFunction <> "" Then szSQL = szSQL & ", RESTRICT = " & szRestrictFunction
  If szJoinFunction <> "" Then szSQL = szSQL & ", JOIN = " & szJoinFunction
  If bHashJoins Then szSQL = szSQL & ", HASHES"
  If szLeftTypeSortOperator <> "" Then szSQL = szSQL & ", SORT1 = " & szLeftTypeSortOperator
  If szRightTypeSortOperator <> "" Then szSQL = szSQL & ", SORT2 = " & szRightTypeSortOperator
  
  szSQL = szSQL & ");" & vbCrLf

  If Comment <> "" Then szSQL = szSQL & "COMMENT ON OPERATOR " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf
  
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
