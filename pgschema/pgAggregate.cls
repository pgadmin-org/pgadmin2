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
Private dTimeStamp As Date

'Set/Return the TimeStamp
Friend Property Let TimeStamp(ByVal vData As Date)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

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
Attribute Database.VB_Description = "Returns the parent database name."
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
Attribute Identifier.VB_Description = "Returns the object Identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
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
'On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  If szInputType = "opaque" Then
    szSQL = "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " * IS '" & dbSZ(vData) & "'"
  Else
    szSQL = "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " " & QUOTE & szInputType & QUOTE & " IS '" & dbSZ(vData) & "'"
  End If
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Aggregate comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get Comment()", etFullDebug
  
  Comment = objServer.Databases(Database).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Aggregate: " & szIdentifier & vbCrLf & _
          "CREATE AGGREGATE " & QUOTE & szName & QUOTE & "("
          
  If szInputType = "opaque" Then
    szSQL = szSQL & "BASETYPE = " & QUOTE & "ANY" & QUOTE & ", "
  Else
    szSQL = szSQL & "BASETYPE = " & QUOTE & szInputType & QUOTE & ", "
  End If
  
  szSQL = szSQL & "SFUNC = " & QUOTE & szStateFunction & QUOTE & ", " & _
                  "STYPE = " & QUOTE & szStateType & QUOTE
          
  If szFinalFunction <> "" Then szSQL = szSQL & ", FINALFUNC = " & QUOTE & szFinalFunction & QUOTE
  If szInitialCondition <> "" Then szSQL = szSQL & ", INITCOND = '" & szInitialCondition & "'"
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then
    If szInputType = "opaque" Then
      szSQL = szSQL & "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " * IS '" & dbSZ(Comment) & "';" & vbCrLf
    Else
      szSQL = szSQL & "COMMENT ON AGGREGATE " & QUOTE & szName & QUOTE & " " & QUOTE & dbSZ(szInputType) & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    End If
  End If
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgAggregate.RCStatus()", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If Not objServer.Databases(Database).RevisionControl Then
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

