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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private szTable As String
Private lTableOID As Long
Private lOID As Long
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

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Table
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Table() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Table()", etFullDebug

  Table = szTable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table OID
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Name(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER TABLE " & QUOTE & dbSZ(szTable) & QUOTE & " RENAME COLUMN " & QUOTE & dbSZ(szName) & QUOTE & " TO " & QUOTE & dbSZ(vData) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Column name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get ObjectType()", etFullDebug

  ObjectType = "Column"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Position
Friend Property Let Position(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Position(" & vData & ")", etFullDebug

  lPosition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the Column's Ordinal Position."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Position()", etFullDebug

  Position = lPosition
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Length
Friend Property Let Length(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Length(" & vData & ")", etFullDebug

  lLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Length() As Long
Attribute Length.VB_Description = "Sets/Returns the Column Length."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Length()", etFullDebug

  Length = lLength
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the NumericScale
Friend Property Let NumericScale(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let NumericScale(" & vData & ")", etFullDebug

  lNumericScale = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NumericScale() As Long
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get NumericScale()", etFullDebug

  NumericScale = lNumericScale
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Type
Friend Property Let DataType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let DataType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDataType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get DataType() As String
Attribute DataType.VB_Description = "Returns the Column's Data Type"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get DataType()", etFullDebug

  DataType = szDataType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return NotNull
Friend Property Let NotNull(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let NotNull(" & vData & ")", etFullDebug

  bNotNull = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NotNull() As Boolean
Attribute NotNull.VB_Description = "Returns a value indicating whether or not NULL values are restricted."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get NotNull()", etFullDebug

  NotNull = bNotNull
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return PrimaryKey
Friend Property Let PrimaryKey(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let PrimaryKey(" & vData & ")", etFullDebug

  bPrimaryKey = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get PrimaryKey() As Boolean
Attribute PrimaryKey.VB_Description = "Returns a boolean indicating whether the column forms part of a primary key."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get PrimaryKey()", etFullDebug

  PrimaryKey = bPrimaryKey
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Default
Friend Property Let iDefault(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let iDefault(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Let Default(ByVal vData As String)
Attribute Default.VB_Description = "Sets/Returns the columns Default Value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgTable.Property Let Default(" & vData & ")", etFullDebug

Dim szSQL As String
  
  If Default = "" Then
    szSQL = "ALTER TABLE " & QUOTE & dbSZ(szTable) & QUOTE & " ALTER COLUMN " & QUOTE & dbSZ(szName) & QUOTE & " DROP DEFAULT"
  Else
    szSQL = "ALTER TABLE " & QUOTE & dbSZ(szTable) & QUOTE & " ALTER COLUMN " & QUOTE & dbSZ(szName) & QUOTE & " SET DEFAULT " & vData
  End If
  objServer.ExecSQL szSQL, cnDatabase
  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Default() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Default()", etFullDebug

  Default = szDefault
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON COLUMN " & QUOTE & szTable & QUOTE & "." & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Column comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgColumn.Property Get Comment()", etFullDebug

  'In 7.1, comments were stored in pg_description against the column OID.
  'In 7.2+, use the table OID with a SubID of the Column Position
  If objServer.dbVersion.VersionNum < 7.2 Then
    Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
  Else
    Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lTableOID, lPosition)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
