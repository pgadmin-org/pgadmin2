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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private bSystemObject As Boolean
Private lOID As Long
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

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Aggregate OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Aggregate name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get ObjectType()", etFullDebug

  ObjectType = "Aggregate"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get SystemObject()", etFullDebug

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
Attribute Owner.VB_Description = "Returns the Aggregate's Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the InputType
Friend Property Let InputType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let InputType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInputType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get InputType() As String
Attribute InputType.VB_Description = "Returns the Aggregates Input Type (Base Type)"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get InputType()", etFullDebug

  InputType = szInputType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the StateFunction
Friend Property Let StateFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let StateFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStateFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get StateFunction() As String
Attribute StateFunction.VB_Description = "Returns the Aggregate's State Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get StateFunction()", etFullDebug

  StateFunction = szStateFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the StateType
Friend Property Let StateType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let StateType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStateType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get StateType() As String
Attribute StateType.VB_Description = "Returns the Aggregate's State Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get StateType()", etFullDebug

  StateType = szStateType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the FinalFunction
Friend Property Let FinalFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let FinalFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szFinalFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get FinalFunction() As String
Attribute FinalFunction.VB_Description = "Returns the Aggregate's Final Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get FinalFunction()", etFullDebug

  FinalFunction = szFinalFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the FinalType
Friend Property Let FinalType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let FinalType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szFinalType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get FinalType() As String
Attribute FinalType.VB_Description = "Returns the Aggregate's Final Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get FinalType()", etFullDebug

  FinalType = szFinalType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the InitialCondition
Friend Property Let InitialCondition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let InitialCondition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInitialCondition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get InitialCondition() As String
Attribute InitialCondition.VB_Description = "Returns the Aggregate's Initial Condition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get InitialCondition()", etFullDebug

  InitialCondition = szInitialCondition
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " " & QUOTE & szInputType & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Aggregate comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Comment()", etFullDebug
  
  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Aggregate: " & szIdentifier & vbCrLf & _
          "CREATE AGGREGATE " & QUOTE & szName & QUOTE & "(" & _
          "BASETYPE = " & QUOTE & szInputType & QUOTE & ", " & _
          "SFUNC = " & QUOTE & szStateFunction & QUOTE & ", " & _
          "STYPE = " & QUOTE & szStateType & QUOTE
          
  If szFinalFunction <> "" Then szSQL = szSQL & ", FINALFUNC = " & QUOTE & szFinalFunction & QUOTE
  If szInitialCondition <> "" Then szSQL = szSQL & ", INITCOND = '" & szInitialCondition & "'"
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " " & QUOTE & dbSZ(szInputType) & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

