VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgRelationship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private szIdentifier As String
Private szLocalColumn As String
Private szReferencedColumn As String
Private cnDatabase As Connection

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Get ObjectType()", etFullDebug

  ObjectType = "Relationship"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the LocalColumn
Friend Property Let LocalColumn(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Let LocalColumn(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLocalColumn = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get LocalColumn() As String
Attribute LocalColumn.VB_Description = "Returns the Local Column in the Relationship."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Get LocalColumn()", etFullDebug

  LocalColumn = szLocalColumn
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the RemoteColumn
Friend Property Let ReferencedColumn(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Let ReferencedColumn(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReferencedColumn = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ReferencedColumn() As String
Attribute ReferencedColumn.VB_Description = "Returns the Referenced Column in the Relationship."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgRelationship.Property Get ReferencedColumn()", etFullDebug

  ReferencedColumn = szReferencedColumn
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property


