VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private szTable As String
Private dTableOID As Double
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private lPosition As Long
Private lLength As Long
Private lNumericScale As Long
Private szDataType As String
Private bNotNull As Boolean
Private bPrimaryKey As Boolean
Private szDefault As String
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant
Private lStatistics As Long

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the parent table name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Table OID
Friend Property Let TableOID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the column OID (PostgreSQL 7.1 or lower)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Name(" & vData & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " RENAME COLUMN " & FormattedID & " TO " & fmtID(vData)
  Else
    szSQL = "ALTER TABLE " & fmtID(szTable) & " RENAME COLUMN " & FormattedID & " TO " & fmtID(vData)
  End If
  objServer.ExecSQL szSQL, cnDatabase
  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Column name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get ObjectType()", etFullDebug

  ObjectType = "Column"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Position
Friend Property Let Position(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Position(" & vData & ")", etFullDebug

  lPosition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the Column's Ordinal Position."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Position()", etFullDebug

  Position = lPosition
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Length
Friend Property Let Length(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Length(" & vData & ")", etFullDebug

  lLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Length() As Long
Attribute Length.VB_Description = "Sets/Returns the Column Length."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Length()", etFullDebug

  Length = lLength
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the NumericScale
Friend Property Let NumericScale(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let NumericScale(" & vData & ")", etFullDebug

  lNumericScale = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NumericScale() As Long
Attribute NumericScale.VB_Description = "Returns the numeric scale for numeric columns."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get NumericScale()", etFullDebug

  NumericScale = lNumericScale
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Type
Friend Property Let DataType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let DataType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDataType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get DataType() As String
Attribute DataType.VB_Description = "Returns the Column's Data Type"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get DataType()", etFullDebug

  DataType = szDataType
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return NotNull
Friend Property Let iNotNull(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iNotNull(" & vData & ")", etFullDebug

  bNotNull = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let NotNull(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let NotNull(" & vData & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 3, App.Title, "NOT NULL can only be set or dropped post column creation in PostgreSQL 7.3 or above."
    Exit Property
  End If
  
  If objVersion.VersionNum >= 7.3 Then
    If vData Then
      szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET NOT NULL"
    Else
      szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " DROP NOT NULL"
    End If
  Else
    If vData Then
      szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET NOT NULL"
    Else
      szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " DROP NOT NULL"
    End If
  End If
  objServer.ExecSQL szSQL, cnDatabase
  bNotNull = vData

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NotNull() As Boolean
Attribute NotNull.VB_Description = "Returns a value indicating whether or not NULL values are restricted."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get NotNull()", etFullDebug

  NotNull = bNotNull
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return PrimaryKey
Public Property Let PrimaryKey(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let PrimaryKey(" & vData & ")", etFullDebug

Dim szSQL As String
Dim szPkeyName As String
Dim objIndex As pgIndex
Dim objColumn As pgColumn

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 8, App.Title, "Primary Keys can only be added or dropped in PostgreSQL 7.3+."
    Exit Property
  End If
  
  If vData Then
   
    'First, set the column to NOT NULL if it is not already
    If Not bNotNull Then NotNull = True
    
    'Now, alter the table
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ADD PRIMARY KEY (" & FormattedID & ")"
    objServer.ExecSQL szSQL, cnDatabase
    
  Else
  
    'Drop all the primary indexes (in reality there is can only be one
    'unless something is *really* screwed.
  
    For Each objIndex In objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Indexes
      If objIndex.Primary Then objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Indexes.Remove objIndex.Identifier
    Next objIndex
  
    'Now reset the PrimaryKey property for all columns on the table.
    For Each objColumn In objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Columns
      objColumn.iPrimaryKey = False
    Next objColumn
    
  End If
  
  bPrimaryKey = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Let iPrimaryKey(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iPrimaryKey(" & vData & ")", etFullDebug

  bPrimaryKey = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get PrimaryKey() As Boolean
Attribute PrimaryKey.VB_Description = "Returns a boolean indicating whether the column forms part of a primary key."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get PrimaryKey()", etFullDebug

  PrimaryKey = bPrimaryKey
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Default
Friend Property Let iDefault(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iDefault(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let Default(ByVal vData As String)
Attribute Default.VB_Description = "Sets/Returns the columns Default Value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Default(" & vData & ")", etFullDebug

Dim szSQL As String
  
  If objVersion.VersionNum >= 7.3 Then
    If vData = "" Then
      szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " DROP DEFAULT"
    Else
      szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET DEFAULT " & vData
    End If
  Else
    If vData = "" Then
      szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " DROP DEFAULT"
    Else
      szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET DEFAULT " & vData
    End If
  End If
  objServer.ExecSQL szSQL, cnDatabase
  szDefault = vData

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Default() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Default()", etFullDebug

  Default = szDefault
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return Statistics
Friend Property Let iStatistics(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iStatistics(" & vData & ")", etFullDebug

  lStatistics = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let Statistics(ByVal vData As Long)
Attribute Statistics.VB_Description = "Sets/Returns the columns statistics value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Statistics(" & vData & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.2 Then
    Err.Raise vbObjectError + 15, App.Title, "Column statistics are only available in PostgreSQL 7.2+."
    Exit Property
  End If
  
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET STATISTICS " & vData
  Else
    szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & FormattedID & " SET STATISTICS " & vData
  End If
  objServer.ExecSQL szSQL, cnDatabase
  lStatistics = vData

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Statistics() As Long
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Statistics()", etFullDebug

  Statistics = lStatistics
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum >= 7.3 Then
    szSQL = "COMMENT ON COLUMN " & fmtID(szNamespace) & "." & fmtID(szTable) & "." & FormattedID & " IS '" & dbSZ(vData) & "'"
  Else
    szSQL = "COMMENT ON COLUMN " & fmtID(szTable) & "." & FormattedID & " IS '" & dbSZ(vData) & "'"
  End If
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Column comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Comment()", etFullDebug

  'In 7.1, comments were stored in pg_description against the column OID.
  'In 7.2+, use the table OID with a SubID of the Column Position
  If objVersion.VersionNum < 7.2 Then
    Comment = objServer.Databases(Database).LookupComment(dOID)
  Else
    Comment = objServer.Databases(Database).LookupComment(dTableOID, CDbl(lPosition))
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
