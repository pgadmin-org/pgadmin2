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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
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
Private lNamespaceOID As Long
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Friend Property Let Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Tag()", etFullDebug

  Tag = vTag
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
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

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
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

  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "COMMENT ON RULE " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  Else
    szSQL = "COMMENT ON RULE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  End If
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache

  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Rule comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgRule.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(lOID)
    
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

  If objServer.dbVersion.VersionNum >= 7.3 Then
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON RULE " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
  Else
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON RULE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
  End If
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property
