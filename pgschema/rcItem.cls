VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rcItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private szParentTable As String
Private szObjectType As String
Private szIdentifier As String
Private dTimeStamp As Date
Private szUser As String
Private szAction As String
Private lOID As Long
Private lVersion As Long
Private szDefinition As String
Private szComment As String

'Set/Return the ParentTable
Friend Property Let ParentTable(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let ParentTable(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szParentTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ParentTable() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get ParentTable()", etFullDebug

  ParentTable = szParentTable
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ObjectType
Friend Property Let ObjectType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let ObjectType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szObjectType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get ObjectType()", etFullDebug

  ObjectType = szObjectType
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TimeStamp
Friend Property Let TimeStamp(ByVal vData As Date)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
Attribute TimeStamp.VB_Description = "Returns the entry timestamp."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the User
Friend Property Let User(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let User(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szUser = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get User() As String
Attribute User.VB_Description = "Returns the user associated with the entry."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get User()", etFullDebug

  User = szUser
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Action
Friend Property Let Action(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let Action(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szAction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Action() As String
Attribute Action.VB_Description = "Returns the entry action."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get Action()", etFullDebug

  Action = szAction
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the objects OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Version
Friend Property Let Version(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let Version(" & vData & ")", etFullDebug

  lVersion = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Version() As Long
Attribute Version.VB_Description = "Returns the entry version number."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get Version()", etFullDebug

  Version = lVersion
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Definition
Friend Property Let Definition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let Definition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefinition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Definition() As String
Attribute Definition.VB_Description = "Returns the object's definition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get Definition()", etFullDebug

  Definition = szDefinition
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Friend Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Let Comment(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szComment = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Returns the entry comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":rcItem.Property Get Comment()", etFullDebug

  Comment = szComment
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
