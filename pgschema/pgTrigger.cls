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

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Trigger OID."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Trigger name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get ObjectType()", etFullDebug

  ObjectType = "Trigger"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether or not this is a system object."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TriggerFunction
Friend Property Let TriggerFunction(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let TriggerFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTriggerFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TriggerFunction() As String
Attribute TriggerFunction.VB_Description = "Returns the function the Trigger will execute."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get TriggerFunction()", etFullDebug

  TriggerFunction = szTriggerFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Executes
Friend Property Let Executes(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Executes(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szExecutes = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Executes() As String
Attribute Executes.VB_Description = "Returns a value indicating when the Trigger will fire - 'Before' or 'After' the firing statement is executed."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Executes()", etFullDebug

  Executes = szExecutes
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TriggerEvent
Friend Property Let TriggerEvent(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let TriggerEvent(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTriggerEvent = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TriggerEvent() As String
Attribute TriggerEvent.VB_Description = "Returns a string indicating the events that will fire the Trigger."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get TriggerEvent()", etFullDebug

  TriggerEvent = szTriggerEvent
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return ForEach
Friend Property Let ForEach(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let ForEach(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szForEach = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ForEach() As String
Attribute ForEach.VB_Description = "Returns a value indicating whether the Trigger fires for each Statement or affected Row."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get ForEach()", etFullDebug

  ForEach = szForEach
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "COMMENT ON TRIGGER " & QUOTE & szName & QUOTE & " ON " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szTable & QUOTE & " IS '" & dbSZ(vData) & "'"
  Else
    szSQL = "COMMENT ON TRIGGER " & QUOTE & szName & QUOTE & " ON " & QUOTE & szTable & QUOTE & " IS '" & dbSZ(vData) & "'"
  End If
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Trigger comment."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTrigger.Property Get SQL()", etFullDebug
    
Dim szSQL As String

  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "-- Trigger: " & szNamespace & "." & szIdentifier & vbCrLf & _
            "CREATE TRIGGER " & QUOTE & szName & QUOTE & _
            " " & szExecutes & " " & szTriggerEvent & _
            " ON " & QUOTE & QUOTE & szNamespace & QUOTE & "." & QUOTE & szTable & QUOTE & _
            " FOR EACH " & szForEach & _
            " EXECUTE PROCEDURE " & szTriggerFunction & _
            ";" & vbCrLf
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON TRIGGER " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szName & QUOTE & " ON " & QUOTE & szTable & QUOTE & " IS '" & Comment & "';" & vbCrLf
  Else
    szSQL = "-- Trigger: " & szIdentifier & vbCrLf & _
            "CREATE TRIGGER " & QUOTE & szName & QUOTE & _
            " " & szExecutes & " " & szTriggerEvent & _
            " ON " & QUOTE & szTable & QUOTE & _
            " FOR EACH " & szForEach & _
            " EXECUTE PROCEDURE " & szTriggerFunction & _
            ";" & vbCrLf
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON TRIGGER " & QUOTE & szName & QUOTE & " ON " & QUOTE & szTable & QUOTE & " IS '" & Comment & "';" & vbCrLf
  End If

  SQL = szSQL
  
Exit Property
Err_Handler:  objServer.iLogError Err
End Property
