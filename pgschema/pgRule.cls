VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private szTable As String
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As String
Private szDefinition As String
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get History()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Rule OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Rule name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get ObjectType()", etFullDebug

  ObjectType = "Rule"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Definition
Friend Property Let Definition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Definition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefinition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Definition() As String
Attribute Definition.VB_Description = "Returns the Rule Definition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Definition()", etFullDebug

  Definition = szDefinition
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Event
Public Property Get RuleEvent() As String
Attribute RuleEvent.VB_Description = "Returns the event that will fire the rule (INSERT, UPDATE, SELECT, DELETE)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get RuleEvent()", etFullDebug

  'Extract the Event
  RuleEvent = Mid(szDefinition, InStr(1, szDefinition, " AS ON ") + 7, 6)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Condition
Public Property Get Condition() As String
Attribute Condition.VB_Description = "Returns the conditions (if any) under which the rule will apply."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Condition()", etFullDebug

Dim szTemp As String
Dim lStart As Long
Dim lLength As Long

  'Quoted Table name
  szTemp = "TO " & QUOTE & szTable & QUOTE & " WHERE "
  If InStr(1, szDefinition, szTemp) <> 0 Then
    lStart = InStr(1, szDefinition, szTemp) + Len(szTemp)
    lLength = InStr(lStart, szDefinition, " DO ") - lStart
    Condition = Mid(szDefinition, lStart, lLength)
    Exit Property
  End If
  
  'Unquoted tablename
  szTemp = "TO " & szTable & " WHERE "
  If InStr(1, szDefinition, szTemp) <> 0 Then
    lStart = InStr(1, szDefinition, szTemp) + Len(szTemp)
    lLength = InStr(lStart, szDefinition, " DO ") - lStart
    Condition = Mid(szDefinition, lStart, lLength)
    Exit Property
  End If
  
  Condition = ""
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return DoInstead
Public Property Get DoInstead() As Boolean
Attribute DoInstead.VB_Description = "Returns a boolean indicating whether the Action will be performed instead of the original query."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get DoInstead()", etFullDebug

  If InStr(1, szDefinition, " DO INSTEAD ") <> 0 Then
    DoInstead = True
  Else
    DoInstead = False
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Action
Public Property Get Action() As String
Attribute Action.VB_Description = "Returns the action the the rule will take."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Action()", etFullDebug

Dim szTemp As String

  If InStr(1, szDefinition, " DO INSTEAD ") <> 0 Then
    szTemp = Mid(szDefinition, InStr(1, szDefinition, " DO INSTEAD ") + 12)
  Else
    szTemp = Mid(szDefinition, InStr(1, szDefinition, " DO ") + 4)
  End If
  Action = Left(szTemp, Len(szTemp) - 1)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON RULE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Rule comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Rule: " & szName & vbCrLf & _
          szDefinition & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON RULE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType, szTable
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.RCStatus()", etFullDebug

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


