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
Private lNamespaceOID As Long
Private szNamespace As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get FormattedID()", etFullDebug
 
  If objServer.dbVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szName) & " ON " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    FormattedID = fmtID(szName) & " ON " & fmtID(szTable)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
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
Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Name(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER TRIGGER " & FormattedID & " RENAME TO " & fmtID(vData)
  objServer.ExecSQL szSQL, cnDatabase
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

  szSQL = "COMMENT ON TRIGGER " & FormattedID & " IS '" & dbSZ(vData) & "'"

  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  
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

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get SQL()", etFullDebug
    
Dim szSQL As String

  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "-- Trigger: " & FormattedID & vbCrLf & _
            "CREATE TRIGGER " & fmtID(szName) & _
            " " & szExecutes & " " & szTriggerEvent & _
            " ON " & fmtID(szNamespace) & "." & fmtID(szTable) & _
            " FOR EACH " & szForEach & _
            " EXECUTE PROCEDURE " & szTriggerFunction & _
            ";" & vbCrLf
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON TRIGGER " & FormattedID & " IS '" & Comment & "';" & vbCrLf
  Else
    szSQL = "-- Trigger: " & FormattedID & vbCrLf & _
            "CREATE TRIGGER " & fmtID(szName) & _
            " " & szExecutes & " " & szTriggerEvent & _
            " ON " & fmtID(szTable) & _
            " FOR EACH " & szForEach & _
            " EXECUTE PROCEDURE " & szTriggerFunction & _
            ";" & vbCrLf
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON TRIGGER " & FormattedID & " IS '" & Comment & "';" & vbCrLf
  End If

  SQL = szSQL
  
Exit Property
Err_Handler:  objServer.iLogError Err
End Property
