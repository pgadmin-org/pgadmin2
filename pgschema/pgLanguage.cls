VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgLanguage"
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
Private bTrusted As Boolean
Private szHandler As String
Private szComment As String
Private cnDatabase As Connection
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get FormattedID()", etFullDebug

  FormattedID = "'" & szName & "'"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Language OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Language name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get ObjectType()", etFullDebug

  ObjectType = "Language"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Trusted?
Friend Property Let Trusted(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Trusted(" & vData & ")", etFullDebug

  bTrusted = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Trusted() As Boolean
Attribute Trusted.VB_Description = "Returns a value indicating whether the Language is trusted."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Trusted()", etFullDebug

  Trusted = bTrusted
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Handler
Friend Property Let Handler(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Handler(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szHandler = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Handler() As String
Attribute Handler.VB_Description = "Returns the Language handler."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Handler()", etFullDebug

  Handler = szHandler
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Friend Property Let iComment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Comment(" & vData & ")", etFullDebug

  szComment = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let iComment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "UPDATE pg_language SET lancompiler = '" & dbSZ(vData) & "' WHERE oid = " & lOID
  objServer.ExecSQL szSQL, cnDatabase
  szComment = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Language comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Comment()", etFullDebug
  
  Comment = szComment
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim szMember As Variant
  
  szSQL = "-- Language: " & FormattedID & vbCrLf & _
          "CREATE"
          
  If bTrusted Then szSQL = szSQL & " TRUSTED"
  szSQL = szSQL & " PROCEDURAL LANGUAGE " & FormattedID
  szSQL = szSQL & " HANDLER " & QUOTE & szHandler & QUOTE
  szSQL = szSQL & " LANCOMPILER '" & dbSZ(Comment) & "'"
  szSQL = szSQL & ";" & vbCrLf

  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property
