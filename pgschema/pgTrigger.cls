VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgTrigger"
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
Private szTable As String
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szTriggerFunction As String
Private szExecutes As String
Private szTriggerEvent As String
Private szForEach As String
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get History()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Trigger OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Trigger name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get ObjectType()", etFullDebug

  ObjectType = "Trigger"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether or not this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TriggerFunction
Friend Property Let TriggerFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let TriggerFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTriggerFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TriggerFunction() As String
Attribute TriggerFunction.VB_Description = "Returns the function the Trigger will execute."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get TriggerFunction()", etFullDebug

  TriggerFunction = szTriggerFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Executes
Friend Property Let Executes(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Executes(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szExecutes = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Executes() As String
Attribute Executes.VB_Description = "Returns a value indicating when the Trigger will fire - 'Before' or 'After' the firing statement is executed."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Executes()", etFullDebug

  Executes = szExecutes
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TriggerEvent
Friend Property Let TriggerEvent(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let TriggerEvent(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTriggerEvent = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TriggerEvent() As String
Attribute TriggerEvent.VB_Description = "Returns a string indicating the events that will fire the Trigger."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get TriggerEvent()", etFullDebug

  TriggerEvent = szTriggerEvent
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return ForEach
Friend Property Let ForEach(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let ForEach(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szForEach = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ForEach() As String
Attribute ForEach.VB_Description = "Returns a value indicating whether the Trigger fires for each Statement or affected Row."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get ForEach()", etFullDebug

  ForEach = szForEach
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON TRIGGER " & QUOTE & szName & QUOTE & " ON " & QUOTE & szTable & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Trigger comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL(Optional bSQLDrop As Boolean = False, Optional bSQLCreate As Boolean = True, Optional bSQLComment As Boolean = True, Optional szSQLTriggerName As String, Optional szSQLTriggerTable As String, Optional szSQLTriggerExecutes As String, Optional szSQLTriggerEvent As String, Optional szSQLTriggerForEach As String, Optional szSQLTriggerFunction As String, Optional szSQLTriggerComment As String) As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler

    'Set default values
    If (bSQLDrop = True) And (bSQLCreate = True) And (objServer.dbVersion.VersionNum < 7.2) Then bSQLDrop = False
    If (szSQLTriggerName = "") Then szSQLTriggerName = szName
    If (szSQLTriggerTable = "") Then szSQLTriggerTable = szTable
    If (szSQLTriggerExecutes = "") Then szSQLTriggerExecutes = UCase(szExecutes)
    If (szSQLTriggerEvent = "") Then szSQLTriggerEvent = UCase(szTriggerEvent)
    If (szSQLTriggerForEach = "") Then szSQLTriggerForEach = UCase(szForEach)
    If (szSQLTriggerFunction = "") Then szSQLTriggerFunction = szTriggerFunction
    If (szSQLTriggerComment = "") Then szSQLTriggerComment = dbSZ(Comment)
    
    'Log
    objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get SQL()", etFullDebug
    Dim szSQL As String
    Dim szMember As Variant
    
    szSQL = "-- Trigger: " & szIdentifier & vbCrLf
    
    If (bSQLDrop = True) Then
        szSQL = szSQL & "DROP TRIGGER " & QUOTE & szSQLTriggerName & QUOTE & _
        " ON " & QUOTE & szSQLTriggerTable & QUOTE & ";" & vbCrLf
    End If
    
    If (bSQLCreate = True) Then
        szSQL = szSQL & "CREATE TRIGGER " & QUOTE & szSQLTriggerName & QUOTE & _
        " " & szSQLTriggerExecutes & " " & szSQLTriggerEvent & _
        " ON " & QUOTE & szSQLTriggerTable & QUOTE & _
        " FOR EACH " & szSQLTriggerForEach & _
        " EXECUTE PROCEDURE " & szSQLTriggerFunction & _
        ";" & vbCrLf
    End If
     
    If (bSQLComment = True) Then
        szSQL = szSQL & "COMMENT ON TRIGGER " & QUOTE & szSQLTriggerName & QUOTE & " ON " & QUOTE & szSQLTriggerTable & QUOTE & " IS '" & szSQLTriggerComment & "';" & vbCrLf
    End If
    
    SQL = szSQL
  
Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType, szTable
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.RCStatus()", etFullDebug

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
    If SQL(False, True, False) = rs!rc_definition & "" Then
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

Public Sub Alter(szNewTriggerName As String, szNewTriggerExecutes As String, szNewTriggerEvent As String, szNewTriggerTable As String, szNewTriggerForEach As String, szNewTriggerFunction As String, szNewTriggerComment As String)
On Error GoTo Err_Handler
Dim szSQL As String
Dim rs As New Recordset

  objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Sub Alter(" & szNewTriggerName & ", " & szNewTriggerExecutes & ", " & szNewTriggerEvent & ", " & szNewTriggerTable & ", " & szNewTriggerForEach & ", " & szNewTriggerFunction & ", " & szNewTriggerComment & ")", etFullDebug
  
  szSQL = SQL(True, False, False)
  objServer.ExecSQL szSQL, cnDatabase

  szSQL = SQL(False, True, True, szNewTriggerName, szNewTriggerTable, szNewTriggerExecutes, szNewTriggerEvent, szNewTriggerForEach, szNewTriggerFunction, szNewTriggerComment)
  objServer.ExecSQL szSQL, cnDatabase

  'Select oid back from the DB...
  Set rs = objServer.ExecSQL("SELECT oid FROM pg_trigger WHERE tgname = '" & szName & "'", cnDatabase)
  If Not rs.EOF Then lOID = rs!OID & ""
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  If Not SystemObject Then Commit rcUpdate, "Altered trigger."
Exit Sub
  
Err_Handler:
Set rs = Nothing
objServer.iLogError Err
End Sub
