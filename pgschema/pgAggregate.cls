VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Properties
Private bSystemObject As Boolean
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private szOwner As String
Private szInputType As String
Private szStateFunction As String
Private szStateType As String
Private szFinalFunction As String
Private szFinalType As String
Private szInitialCondition As String
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant
Private colDependent As Collection
Private colReferenced As Collection

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get FormattedID()", etFullDebug

  If objVersion.VersionNum >= 7.3 Then
    If szInputType = "opaque" Then
      FormattedID = fmtID(szNamespace) & "." & fmtID(szName) & " *"
    Else
      FormattedID = fmtID(szNamespace) & "." & fmtID(szName) & " " & fmtID(szInputType)
    End If
  Else
    If szInputType = "opaque" Then
      FormattedID = fmtID(szName) & " *"
    Else
      FormattedID = fmtID(szName) & " " & fmtID(szInputType)
    End If
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let NameSpace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NameSpace() As String
Attribute NameSpace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Namespace()", etFullDebug

  NameSpace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let Oid(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Oid() As Double
Attribute Oid.VB_Description = "Returns the Aggregate OID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get OID()", etFullDebug

  Oid = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Aggregate name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object Identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get ObjectType()", etFullDebug

  ObjectType = "Aggregate"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Returns a boolean indicating whether this is a system object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Aggregate's Owner."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the InputType
Friend Property Let InputType(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let InputType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInputType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get InputType() As String
Attribute InputType.VB_Description = "Returns the Aggregates Input Type (Base Type)"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get InputType()", etFullDebug

  InputType = szInputType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the StateFunction
Friend Property Let StateFunction(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let StateFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStateFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get StateFunction() As String
Attribute StateFunction.VB_Description = "Returns the Aggregate's State Function."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get StateFunction()", etFullDebug

  StateFunction = szStateFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the StateType
Friend Property Let StateType(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let StateType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStateType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get StateType() As String
Attribute StateType.VB_Description = "Returns the Aggregate's State Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get StateType()", etFullDebug

  StateType = szStateType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the FinalFunction
Friend Property Let FinalFunction(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let FinalFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szFinalFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get FinalFunction() As String
Attribute FinalFunction.VB_Description = "Returns the Aggregate's Final Function."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get FinalFunction()", etFullDebug

  FinalFunction = szFinalFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the FinalType
Friend Property Let FinalType(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let FinalType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szFinalType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get FinalType() As String
Attribute FinalType.VB_Description = "Returns the Aggregate's Final Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get FinalType()", etFullDebug

  FinalType = szFinalType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the InitialCondition
Friend Property Let InitialCondition(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let InitialCondition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInitialCondition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get InitialCondition() As String
Attribute InitialCondition.VB_Description = "Returns the Aggregate's Initial Condition."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get InitialCondition()", etFullDebug

  InitialCondition = szInitialCondition
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON AGGREGATE " & FormattedID & " IS '" & dbSZ(vData) & "'"

  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Aggregate comment."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Comment()", etFullDebug
  
  Comment = objServer.Databases(Database).LookupComment(dOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "-- Aggregate: " & szNamespace & "." & szIdentifier & vbCrLf & _
            "CREATE AGGREGATE " & fmtID(szNamespace) & "." & fmtID(szName) & " ("
  Else
    szSQL = "-- Aggregate: " & szIdentifier & vbCrLf & _
            "CREATE AGGREGATE " & fmtID(szName) & " ("
  End If
          
  If szInputType = "opaque" Then
    szSQL = szSQL & "BASETYPE = " & QUOTE & "ANY" & QUOTE & ", "
  Else
    szSQL = szSQL & "BASETYPE = " & fmtID(szInputType) & ", "
  End If
  
  szSQL = szSQL & "SFUNC = " & fmtID(szStateFunction) & ", " & _
                  "STYPE = " & fmtID(szStateType)
          
  If (szFinalFunction <> "") And (szFinalFunction <> "-") Then szSQL = szSQL & ", FINALFUNC = " & fmtID(szFinalFunction)
  If szInitialCondition <> "" Then szSQL = szSQL & ", INITCOND = '" & szInitialCondition & "'"
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON AGGREGATE " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the dependent object is in
Public Function Dependent() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Function Dependent()", etFullDebug
  
  If colDependent Is Nothing Then Set colDependent = DepRef(dOID, cnDatabase, Me.Database, EDR_Depend)
  Set Dependent = colDependent
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Return the referenced object is in
Public Function Referenced() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Function Dependent()", etFullDebug
  
  If colReferenced Is Nothing Then Set colReferenced = DepRef(dOID, cnDatabase, Me.Database, EDR_Reference)
  Set Referenced = colReferenced
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Class_Terminate()", etFullDebug
  
  Set colDependent = Nothing
  Set colReferenced = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

