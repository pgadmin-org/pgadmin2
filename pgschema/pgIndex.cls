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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colIndexedColumns As IndexedColumns

'Properties
Private szTable As String
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private bUnique As Boolean
Private bPrimary As Boolean
Private szDefinition As String
Private cnDatabase As Connection
Private dTableOID As Double
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get FormattedID()", etFullDebug
 
  If objServer.dbVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szNamespace) & "." & fmtID(szName)
  Else
    FormattedID = fmtID(szName)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the IndexedColumns Collection
Public Property Get IndexedColumns() As IndexedColumns
Attribute IndexedColumns.VB_Description = "Returns a collection of Indexed Columns."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get IndexedColumns()", etFullDebug

  If colIndexedColumns Is Nothing Then
    Set colIndexedColumns = New IndexedColumns
    
    'Pass the connection to the class
    Set colIndexedColumns.Connection = cnDatabase
    colIndexedColumns.Initialize dOID
  End If
  Set IndexedColumns = colIndexedColumns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the Index OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER TABLE " & FormattedID & " RENAME TO " & QUOTE & vData & QUOTE

  objServer.ExecSQL szSQL, cnDatabase
  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Index name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get ObjectType()", etFullDebug

  ObjectType = "Index"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return a boolean indicating whether this is a System Object
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get SystemObject()", etFullDebug

  'There are a 2 situations that define an index as being a System Index
  ' 1) Unique = True AND Name = TableName_FirstColName_key (i.e. created as a unique constraint)
  ' 2) Primary = True (i.e. created as a primary key on a table).
  
  If (((bUnique = True) And (szName = szTable & "_" & IndexedColumns(1) & "_key")) Or _
      (bPrimary = True) Or Left(szName, 8) = "pgadmin_") Then
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
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get IndexType()", etFullDebug

  If InStr(1, szDefinition, " USING hash ") <> 0 Then
    IndexType = "hash"
  ElseIf InStr(1, szDefinition, " USING btree ") <> 0 Then
    IndexType = "btree"
  ElseIf InStr(1, szDefinition, " USING rtree ") <> 0 Then
    IndexType = "rtree"
  ElseIf InStr(1, szDefinition, " USING gist ") <> 0 Then
    IndexType = "gist"
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Unique
Friend Property Let Unique(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Unique(" & vData & ")", etFullDebug

  bUnique = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Unique() As Boolean
Attribute Unique.VB_Description = "Returns a value indicating whether the Index is a Unique Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Unique()", etFullDebug

  Unique = bUnique
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Primary
Friend Property Let Primary(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Primary(" & vData & ")", etFullDebug

  bPrimary = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Primary() As Boolean
Attribute Primary.VB_Description = "Returns a value indicating whether the index is a Primary Key."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Primary()", etFullDebug

  Primary = bPrimary
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Definition
Friend Property Let Definition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Definition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefinition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Definition() As String
Attribute Definition.VB_Description = "Returns the Index Definition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Definition()", etFullDebug

  Definition = szDefinition
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Index Constraint
Public Property Get Constraint() As String
Attribute Constraint.VB_Description = "Returns the Index Constraint. This should only be present for PostgreSQL 7.2+"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Constraint()", etFullDebug

Dim lPos As Long

  lPos = InStr(1, szDefinition, " WHERE ")
  If lPos <> 0 Then Constraint = Mid(szDefinition, lPos + 7)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON INDEX " & FormattedID & " IS '" & dbSZ(vData) & "'"

  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Index comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(dOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Class_Terminate()", etFullDebug
  
  Set colIndexedColumns = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgIndex.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Index: " & FormattedID & vbCrLf & _
          szDefinition & ";" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON INDEX " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property
