VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgForeignKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Dim colRelationships As Relationships

'Properties
Private lTableOID As Long
Private szTable As String
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szReferencedTable As String
Private szLocalColumns As String
Private szReferencedColumns As String
Private szOnDelete As String
Private szOnUpdate As String
Private bDeferrable As Boolean
Private szInitially As String
Private cnDatabase As Connection
Private lNamespaceOID As Long
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Friend Property Let Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Tag()", etFullDebug

  Tag = vTag
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Relationships Collection
Public Property Get Relationships() As Relationships
Attribute Relationships.VB_Description = "Returns a collection of column Relationships."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Relationships()", etFullDebug

  If colRelationships Is Nothing Then
    Set colRelationships = New Relationships
    Set colRelationships.Connection = cnDatabase
    colRelationships.Initialize szLocalColumns, szReferencedColumns
  End If
  Set Relationships = colRelationships
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Foreign Key OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Foreign Key name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get ObjectType()", etFullDebug

  ObjectType = "Foreign Key"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ReferencedTable
Friend Property Let ReferencedTable(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let ReferencedTable(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReferencedTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ReferencedTable() As String
Attribute ReferencedTable.VB_Description = "Returns the name of the table referenced by the Foreign Key."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get ReferencedTable()", etFullDebug

  ReferencedTable = szReferencedTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the LocalColumns
Friend Property Let LocalColumns(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let LocalColumns(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLocalColumns = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Get LocalColumns() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get LocalColumns()", etFullDebug

  LocalColumns = szLocalColumns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ReferencedColumns
Friend Property Let ReferencedColumns(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let ReferencedColumns(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReferencedColumns = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Get ReferencedColumns() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get ReferencedColumns()", etFullDebug

  ReferencedColumns = szReferencedColumns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return OnDelete
Public Property Get OnDelete() As String
Attribute OnDelete.VB_Description = "Returns the action to do when a referenced row in the referenced table is being deleted."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get OnDelete()", etFullDebug

Static bCached As Boolean
Dim szSQL As String
Dim rs As New Recordset

  If bCached = False Then
    szSQL = "SELECT proname FROM pg_proc p, pg_trigger t, pg_class c WHERE p.oid = t.tgfoid AND (t.tgconstrrelid = " & lTableOID & " OR t.tgconstrrelid = 0) AND t.tgtype = 9 AND t.tgrelid = c.oid AND c.relname = '" & szReferencedTable & "'"
    Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    If Not rs.EOF Then
      Select Case rs!proname
        Case "RI_FKey_noaction_del"
          szOnDelete = "No Action"
        Case "RI_FKey_restrict_del"
          szOnDelete = "Restrict"
        Case "RI_FKey_cascade_del"
          szOnDelete = "Cascade"
        Case "RI_FKey_setnull_del"
          szOnDelete = "Set Null"
        Case "RI_FKey_setdefault_del"
          szOnDelete = "Set Default"
      End Select
    End If
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    bCached = True
  End If
  
  OnDelete = szOnDelete
    
  Exit Property
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Property

'Return OnUpdate
Public Property Get OnUpdate() As String
Attribute OnUpdate.VB_Description = "Returns the action to do when a referenced row in the referenced table is being updated. "
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get OnUpdate()", etFullDebug

Static bCached As Boolean
Dim szSQL As String
Dim rs As New Recordset

  If bCached = False Then
    szSQL = "SELECT proname FROM pg_proc p, pg_trigger t, pg_class c WHERE p.oid = t.tgfoid AND (t.tgconstrrelid = " & lTableOID & " OR t.tgconstrrelid = 0) AND t.tgtype = 17 AND t.tgrelid = c.oid AND c.relname = '" & szReferencedTable & "'"
    Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    If Not rs.EOF Then
      Select Case rs!proname
        Case "RI_FKey_noaction_upd"
          szOnUpdate = "No Action"
        Case "RI_FKey_restrict_upd"
          szOnUpdate = "Restrict"
        Case "RI_FKey_cascade_upd"
          szOnUpdate = "Cascade"
        Case "RI_FKey_setnull_upd"
          szOnUpdate = "Set Null"
        Case "RI_FKey_setdefault_upd"
          szOnUpdate = "Set Default"
      End Select
    End If
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    bCached = True
  End If
  
  OnUpdate = szOnUpdate
    
  Exit Property
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Property

'Set/Return the Deferrable
Friend Property Let Deferrable(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Deferrable(" & vData & ")", etFullDebug

  bDeferrable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Deferrable() As Boolean
Attribute Deferrable.VB_Description = "Returns a boolean indicating whether the constraint can be deferred to the end of the transaction. If DEFERRABLE, SET CONSTRAINTS ALL DEFERRED will cause the foreign key to be checked only at the end of the transaction. "
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Deferrable()", etFullDebug

  Deferrable = bDeferrable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Initially
Friend Property Let Initially(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Let Initially(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInitially = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Initially() As String
Attribute Initially.VB_Description = "Returns two possible values which specify the default time to check the constraint. \r\n\r\nIMMEDIATE - Check constraint after each statement. This is the default. \r\nDEFERRED - Check constraint only at the end of the transaction. "
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Property Get Initially()", etFullDebug

  Initially = szInitially
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgForeignKey.Class_Terminate()", etFullDebug

  Set colRelationships = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
