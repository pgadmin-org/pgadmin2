VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colColumns As Columns
Private colChecks As Checks
Private colIndexes As Indexes
Private colRules As Rules
Private colTriggers As Triggers
Private colInheritedTables As InheritedTables
Private colForeignKeys As ForeignKeys

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As String
Private szOwner As String
Private szACL As String
Private bHasOIDs As Boolean
Private cnDatabase As Connection

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the ForeignKeys Collection
Public Property Get ForeignKeys() As ForeignKeys
Attribute ForeignKeys.VB_Description = "Returns a collection of Foreign Keys on the current Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ForeignKeys()", etFullDebug

  If colForeignKeys Is Nothing Then
    Set colForeignKeys = New ForeignKeys
    
    'Pass the connection to the class
    Set colForeignKeys.Connection = cnDatabase
    colForeignKeys.TableOID = lOID
    colForeignKeys.Table = szName
    colForeignKeys.Initialize
  End If
  Set ForeignKeys = colForeignKeys
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the InheritedTables Collection
Public Property Get InheritedTables() As InheritedTables
Attribute InheritedTables.VB_Description = "Returns a collection of Inherited Tables on the current Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get InheritedTables()", etFullDebug

  If colInheritedTables Is Nothing Then
    Set colInheritedTables = New InheritedTables
    
    'Pass the connection to the class
    Set colInheritedTables.Connection = cnDatabase
    colInheritedTables.Initialize lOID
  End If
  Set InheritedTables = colInheritedTables
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Columns Collection
Public Property Get Columns() As Columns
Attribute Columns.VB_Description = "Returns a collection of Columns."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Columns()", etFullDebug

  If colColumns Is Nothing Then
    Set colColumns = New Columns
    
    'Pass the connection & oid to the class
    Set colColumns.Connection = cnDatabase
    colColumns.TableOID = lOID
    colColumns.Table = szName
    colColumns.Initialize
  End If
  Set Columns = colColumns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Checks Collection
Public Property Get Checks() As Checks
Attribute Checks.VB_Description = "Returns a collection of Checks on the current Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Checks()", etFullDebug

  If colChecks Is Nothing Then
    Set colChecks = New Checks

    'Pass the connection to the class
    Set colChecks.Connection = cnDatabase
    colChecks.TableOID = lOID
    colChecks.Table = szName
    colChecks.Initialize
  End If
  Set Checks = colChecks
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Indexes Collection
Public Property Get Indexes() As Indexes
Attribute Indexes.VB_Description = "Returns a collection of Indexes on the current Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Indexes()", etFullDebug

  If colIndexes Is Nothing Then
    Set colIndexes = New Indexes

    'Pass the connection to the class
    Set colIndexes.Connection = cnDatabase
    colIndexes.TableOID = lOID
    colIndexes.Table = szName
    colIndexes.Initialize
  End If
  Set Indexes = colIndexes
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Triggers Collection
Public Property Get Triggers() As Triggers
Attribute Triggers.VB_Description = "Returns a collection of Triggers."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Triggers()", etFullDebug

  If colTriggers Is Nothing Then
    Set colTriggers = New Triggers
    
    'Pass the connection & oid to the class
    Set colTriggers.Connection = cnDatabase
    colTriggers.TableOID = lOID
    colTriggers.Table = szName
    colTriggers.Initialize
  End If
  Set Triggers = colTriggers
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Rules Collection
Public Property Get Rules() As Rules
Attribute Rules.VB_Description = "Returns a collection of Rules."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Rules()", etFullDebug

  If colRules Is Nothing Then
    Set colRules = New Rules

    'Pass the connection to the class
    Set colRules.Connection = cnDatabase
    colRules.TableOID = lOID
    colRules.Table = szName
    colRules.Initialize
  End If
  Set Rules = colRules
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Table OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Name(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER TABLE " & QUOTE & dbSZ(szName) & QUOTE & " RENAME TO " & QUOTE & dbSZ(vData) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  szName = vData
  
  'Kill all the collections to force them to be re-initialised
  Set colForeignKeys = Nothing
  Set colInheritedTables = Nothing
  Set colColumns = Nothing
  Set colChecks = Nothing
  Set colIndexes = Nothing
  Set colRules = Nothing
  Set colTriggers = Nothing
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get ObjectType()", etFullDebug

  ObjectType = "Table"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let iOwner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let iOwner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Owner(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER TABLE " & QUOTE & szName & QUOTE & " OWNER TO " & QUOTE & vData & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  szOwner = vData
  If Not SystemObject Then Commit rlUpdate, "Updated object owner."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Sets/Returns the Table Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the Table ACL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON TABLE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  If Not SystemObject Then Commit rlUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Table comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the row count
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns the number of rows in the table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get Comment()", etFullDebug

Dim rs As New Recordset

  Set rs = objServer.ExecSQL("SELECT count(*) AS rows FROM " & QUOTE & dbSZ(szName) & QUOTE, cnDatabase)
  If Not rs.EOF Then Rows = rs!Rows
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
 
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this table has OIDs
Friend Property Let HasOIDs(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let HasOIDs(" & vData & ")", etFullDebug

  bHasOIDs = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get HasOIDs() As Boolean
Attribute HasOIDs.VB_Description = "Indicates whether or not the table has OIDs."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Get HasOIDs()", etFullDebug

  HasOIDs = bHasOIDs
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant privileges on the Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclSelect) = aclSelect Then szSQL = szSQL & "SELECT, "
    If (Privilege And aclInsert) = aclInsert Then szSQL = szSQL & "INSERT, "
    If (Privilege And aclUpdate) = aclUpdate Then szSQL = szSQL & "UPDATE, "
    If (Privilege And aclDelete) = aclDelete Then szSQL = szSQL & "DELETE, "
    If (Privilege And aclRule) = aclRule Then szSQL = szSQL & "RULE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON " & QUOTE & szName & QUOTE & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  If Not SystemObject Then Commit rlUpdate, "Updated object ACL."
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke privileges from the Table."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclSelect) = aclSelect Then szSQL = szSQL & "SELECT, "
    If (Privilege And aclInsert) = aclInsert Then szSQL = szSQL & "INSERT, "
    If (Privilege And aclUpdate) = aclUpdate Then szSQL = szSQL & "UPDATE, "
    If (Privilege And aclDelete) = aclDelete Then szSQL = szSQL & "DELETE, "
    If (Privilege And aclRule) = aclRule Then szSQL = szSQL & "RULE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & " ON " & QUOTE & szName & QUOTE & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  If Not SystemObject Then Commit rlUpdate, "Updated object ACL."
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Class_Terminate()", etFullDebug

  Set colForeignKeys = Nothing
  Set colInheritedTables = Nothing
  Set colColumns = Nothing
  Set colChecks = Nothing
  Set colIndexes = Nothing
  Set colRules = Nothing
  Set colTriggers = Nothing

  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim objColumn As pgColumn
Dim objCheck As pgCheck
Dim objIndex As pgIndex
Dim objForeignKey As pgForeignKey
Dim objRelationship As pgRelationship
Dim vTable As Variant
Dim vColumn As Variant
  
  szSQL = "-- Table: " & szName & vbCrLf & _
          "CREATE TABLE " & QUOTE & szName & QUOTE & " (" & vbCrLf

  For Each objColumn In Columns
    If Not objColumn.SystemObject Then
      szSQL = szSQL & "  " & QUOTE & objColumn.Name & QUOTE 'Name
      szSQL = szSQL & " " & objColumn.DataType 'Type
      
      'Get the length of fixed, definable columns
      If objColumn.DataType = "char" Or _
         objColumn.DataType = "varchar" Then
        szSQL = szSQL & "(" & objColumn.Length & ")"
      End If
      
      'Get the length of the numberic column
      If objColumn.DataType = "numeric" Then
        szSQL = szSQL & "(" & objColumn.Length & ", " & objColumn.NumericScale & ")"
      End If
      
      If objColumn.Default <> "" Then szSQL = szSQL & " DEFAULT " & objColumn.Default 'Default
      If objColumn.NotNull Then szSQL = szSQL & " NOT NULL" 'Not Null
      szSQL = szSQL & ", " & vbCrLf
    End If
  Next objColumn

  'Unique Constraints
  If Indexes.Count > 0 Then
    For Each objIndex In Indexes
      If objIndex.Primary = False And objIndex.Unique = True Then
        If objIndex.Name <> "" Then
          szSQL = szSQL & "  CONSTRAINT " & QUOTE & objIndex.Name & QUOTE & " "
        Else
          szSQL = szSQL & "  "
        End If
        szSQL = szSQL & "UNIQUE ("
        For Each vColumn In objIndex.IndexedColumns
          szSQL = szSQL & QUOTE & vColumn & QUOTE & ", "
        Next vColumn
        szSQL = Left(szSQL, Len(szSQL) - 2) & "), " & vbCrLf
      End If
    Next objIndex
  End If
  
  'Primary Key Constraints
  If Indexes.Count > 0 Then
    For Each objIndex In Indexes
      If objIndex.Primary = True Then
        If objIndex.Name <> "" Then
          szSQL = szSQL & "  CONSTRAINT " & QUOTE & objIndex.Name & QUOTE & " "
        Else
          szSQL = szSQL & "  "
        End If
        szSQL = szSQL & "PRIMARY KEY ("
        For Each vColumn In objIndex.IndexedColumns
          szSQL = szSQL & QUOTE & vColumn & QUOTE & ", "
        Next vColumn
        szSQL = Left(szSQL, Len(szSQL) - 2) & "), " & vbCrLf
      End If
    Next objIndex
  End If
  
  'Checks
  If Checks.Count > 0 Then
    For Each objCheck In Checks
      If objCheck.Name <> "" Then
        szSQL = szSQL & "  CONSTRAINT " & QUOTE & objCheck.Name & QUOTE & " "
      Else
        szSQL = szSQL & "  "
      End If
      szSQL = szSQL & "CHECK (" & objCheck.Definition & "), " & vbCrLf
    Next objCheck
  End If
  
  'Foreign Keys
  If ForeignKeys.Count > 0 Then
    For Each objForeignKey In ForeignKeys
      If objForeignKey.Name <> "" Then
        szSQL = szSQL & "  CONSTRAINT " & QUOTE & objForeignKey.Name & QUOTE & " "
      Else
        szSQL = szSQL & "  "
      End If
      szSQL = szSQL & "FOREIGN KEY (" & Replace(Left(objForeignKey.LocalColumns, Len(objForeignKey.LocalColumns) - 1), "|", ", ") & ") " & _
              "REFERENCES " & QUOTE & objForeignKey.ReferencedTable & QUOTE & _
              " (" & Replace(Left(objForeignKey.ReferencedColumns, Len(objForeignKey.ReferencedColumns) - 1), "|", ", ") & "), " & vbCrLf
    Next objForeignKey
  End If
  
  
  'Lose the ", " & vbcrlf from the last column/check/fkey
  szSQL = Mid(szSQL, 1, Len(szSQL) - 4) & vbCrLf & ")"
  
  'Inherited Tables
  If InheritedTables.Count > 0 Then
    szSQL = szSQL & " INHERITS("
    For Each vTable In InheritedTables
      szSQL = szSQL & QUOTE & vTable & QUOTE & ", "
    Next vTable
    szSQL = Left(szSQL, Len(szSQL) - 2) & ")"
  End If
  
  'OIDs?
  If objServer.dbVersion.VersionNum >= 7.2 Then
    If bHasOIDs Then
      szSQL = szSQL & " WITH OIDS"
    Else
      szSQL = szSQL & " WITHOUT OIDS"
    End If
  End If
  
  szSQL = szSQL & ";" & vbCrLf
  If szACL <> "" Then szSQL = szSQL & ParseACL(szName, szACL)
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON TABLE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rlLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

