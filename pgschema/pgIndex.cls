VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colIndexedColumns As IndexedColumns

'Properties
Private szTable As String
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private bUnique As Boolean
Private bPrimary As Boolean
Private szDefinition As String
Private cnDatabase As Connection
Private lTableOID As Long

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Indexes.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the IndexedColumns Collection
Public Property Get IndexedColumns() As IndexedColumns
Attribute IndexedColumns.VB_Description = "Returns a collection of Indexed Columns."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgDatabase.Property Get IndexedColumns()", etFullDebug

  If colIndexedColumns Is Nothing Then
    Set colIndexedColumns = New IndexedColumns
    
    'Pass the connection to the class
    Set colIndexedColumns.Connection = cnDatabase
    colIndexedColumns.Initialize lOID
  End If
  Set IndexedColumns = colIndexedColumns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Index OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Index name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get ObjectType()", etFullDebug

  ObjectType = "Index"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return a boolean indicating whether this is a System Object
Public Property Get SystemObject() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get SystemObject()", etFullDebug

  'There are a 2 situations that define an index as being a System Index
  ' 1) Unique = True AND Name = TableName_FirstColName_key (i.e. created as a unique constraint)
  ' 2) Primary = True (i.e. created as a primary key on a table).
  
  If (((bUnique = True) And (szName = szTable & "_" & IndexedColumns(1) & "_key")) Or _
      (bPrimary = True)) Then
    SystemObject = True
  Else
    SystemObject = False
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Index Type
Public Property Get IndexType() As String
Attribute IndexType.VB_Description = "Returns the Index Type (btree, rtree or hash)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get IndexType()", etFullDebug

  If InStr(1, szDefinition, " USING hash ") <> 0 Then
    IndexType = "hash"
  ElseIf InStr(1, szDefinition, " USING btree ") <> 0 Then
    IndexType = "btree"
  ElseIf InStr(1, szDefinition, " USING rtree ") <> 0 Then
    IndexType = "rtree"
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Unique
Friend Property Let Unique(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Unique(" & vData & ")", etFullDebug

  bUnique = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Unique() As Boolean
Attribute Unique.VB_Description = "Returns a value indicating whether the Index is a Unique Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Unique()", etFullDebug

  Unique = bUnique
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Primary
Friend Property Let Primary(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Primary(" & vData & ")", etFullDebug

  bPrimary = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Primary() As Boolean
Attribute Primary.VB_Description = "Returns a value indicating whether the index is a Primary Key."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Primary()", etFullDebug

  Primary = bPrimary
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Definition
Friend Property Let Definition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Definition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefinition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Definition() As String
Attribute Definition.VB_Description = "Returns the Index Definition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Definition()", etFullDebug

  Definition = szDefinition
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON INDEX " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Index comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Class_Terminate()", etFullDebug
  
  Set colIndexedColumns = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering pgIndex.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Index: " & szName & vbCrLf & _
          szDefinition & ";" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON INDEX " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

