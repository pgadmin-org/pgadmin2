VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private dTableOID As Double
Private szTable As String
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objColumn As pgColumn

  For Each objColumn In mCol
    If objColumn.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objColumn
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, lPosition As Long, lLength As Long, szDataType As String, bNotNull As Boolean, bPrimaryKey As Boolean, szDefault As String, ByVal lStatistics As Long) As pgColumn
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & lPosition & ", " & lLength & ", " & QUOTE & szDataType & QUOTE & ", " & bNotNull & ", " & bPrimaryKey & ", " & QUOTE & szDefault & QUOTE & ", " & lStatistics & ")", etFullDebug

Dim objNewMember As pgColumn
Dim szHex As String
Dim objInhCol As pgColumn
Dim vInhTab As Variant
Dim szInhTab As String
Dim bFound As Boolean
Dim bNulls As Boolean
Dim bDefaults As Boolean

  'Create a new object
  Set objNewMember = New pgColumn
  If lPosition < 1 Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.iName = szName
  objNewMember.Table = szTable
  objNewMember.TableOID = dTableOID
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  
  'Only have an OID pre 7.2
  If objVersion.VersionNum < 7.2 Then
    objNewMember.OID = dOID
  End If
  
  objNewMember.Position = lPosition
  If szDataType = "bpchar" Then szDataType = "char"
  objNewMember.DataType = szDataType
  If szDataType = "numeric" Then
    szHex = Hex((Int(lLength) - 4) And &HFFFF)
    objNewMember.Length = CLng("&H" & Mid(szHex, 1, Len(szHex) - 4))
    objNewMember.NumericScale = CLng("&H" & Mid(szHex, Len(szHex) - 3, Len(szHex)))
  Else
    objNewMember.Length = lLength
    objNewMember.NumericScale = 0
  End If
  objNewMember.iNotNull = bNotNull
  objNewMember.iPrimaryKey = bPrimaryKey
  objNewMember.iDefault = szDefault
  objNewMember.iStatistics = lStatistics
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'We should only add columns here if they are not inherited. The following
  'rules determine that, and are taken from src/bin/pg_dump/common.c
  '
  ' * For each attr, check the parent info: if no parent has an attr
  ' * with the same name, then it's not inherited. If there *is* an
  ' * attr with the same name, then only dump it if:
  ' *
  ' * - it is NOT NULL and zero parents are NOT NULL OR - it has a
  ' * default value AND the default value does not match all parent
  ' * default values, or no parents specify a default.
  
  If (objServer.Databases(Database).Namespaces(Namespace).Tables(szTable).InheritedTables.Count > 0) And (objNewMember.Position > 0) Then
    For Each vInhTab In objServer.Databases(Database).Namespaces(Namespace).Tables(szTable).InheritedTables
      'Must convert vInhTab to a String or will get a By Ref error
      szInhTab = vInhTab
      For Each objInhCol In objServer.Databases(Database).Namespaces(Namespace).Tables(szInhTab).Columns
        If objInhCol.Position > 0 Then
          If objInhCol.Name = objNewMember.Name Then
            bFound = True
            If objNewMember.NotNull = True Then
              If objInhCol.NotNull = False Then bNulls = True
            End If
            If objNewMember.Default <> "" Then
              If objInhCol.Default <> objNewMember.Default Then bDefaults = True
            End If
          End If
        End If
      Next objInhCol
    Next vInhTab
  End If
  
  If (Not bFound) Or (bNulls Or bDefaults) Then
    mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
    Set iAdd = objNewMember
  End If
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal DataType As String, Optional ByVal Default As String, Optional ByVal Comment As String) As pgColumn
Attribute Add.VB_Description = "Add a new column."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & DataType & QUOTE & ", " & QUOTE & Default & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objColumn As pgColumn
Dim lStatistics As String
  
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    szSQL = "ALTER TABLE " & fmtID(szTable)
  End If
  szSQL = szSQL & " ADD COLUMN " & fmtID(Name) & " " & DataType
  objServer.ExecSQL szSQL, cnDatabase

  If Default <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " ALTER COLUMN " & fmtID(Name) & " SET DEFAULT " & Default
    Else
      szSQL = "ALTER TABLE " & fmtID(szTable) & " ALTER COLUMN " & fmtID(Name) & " SET DEFAULT " & Default
    End If
    objServer.ExecSQL szSQL, cnDatabase
  End If
  
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON COLUMN " & fmtID(szNamespace) & "." & fmtID(szTable) & "." & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON COLUMN " & fmtID(szTable) & "." & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If objVersion.VersionNum < 7.2 Then
    szSQL = SQL_GET_COLUMNS7_1 & " AND attrelid = " & dTableOID & "::oid AND attname = '" & dbSZ(Name) & "'"
  Else
    szSQL = SQL_GET_COLUMNS7_2 & " AND attrelid = " & dTableOID & "::oid AND attname = '" & dbSZ(Name) & "'"
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    If objVersion.VersionNum >= 7.2 Then
      lStatistics = rs!attstattarget
    Else
      lStatistics = 0
    End If
    Set objColumn = iAdd(rs!OID, rs!attname & "", Val(rs!attnum), Val(rs!Length), rs!typname & "", ToBool(rs!attnotnull), ToBool(rs!PrimaryKey), rs!Default & "", lStatistics)
  End If
    
  Set Add = objColumn
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Rename
Public Sub Rename(ByVal OldName As String, ByVal NewName As String)
Attribute Rename.VB_Description = "Rename the specified column."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objColumn As New pgColumn
  
  Set objColumn = mCol(ULEncode(OldName))
  objColumn.Name = NewName
  objColumn.Identifier = NewName
  mCol.Add objColumn, ULEncode(NewName), ULEncode(OldName)
  mCol.Remove ULEncode(OldName)
  
  'Refresh all subclasses of the parent table.
  objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Checks.Refresh
  objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).ForeignKeys.Refresh
  objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Indexes.Refresh
  objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Rules.Refresh
  objServer.Databases(Database).Namespaces(szNamespace).Tables(szTable).Triggers.Refresh
  
  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgColumn
Attribute Item.VB_Description = "Returns a Column object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Columns in the current Table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objColumn As pgColumn
Dim lCount As Long

  If IgnoreSys Then
    For Each objColumn In mCol
      If Not objColumn.SystemObject Then lCount = lCount + 1
    Next objColumn
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the collection of Columns."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Columns.Initialize()", etFullDebug

Dim rs As New Recordset
Dim lStatistics As Long

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum < 7.2 Then
      Set rs = objServer.ExecSQL(SQL_GET_COLUMNS7_1 & " AND attrelid = " & dTableOID & "::oid ORDER BY attnum", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_COLUMNS7_2 & " AND attrelid = " & dTableOID & "::oid ORDER BY attnum", cnDatabase)
    End If
    While Not rs.EOF
      If objVersion.VersionNum >= 7.2 Then
        lStatistics = rs!attstattarget
      Else
        lStatistics = 0
      End If
      iAdd rs!OID, rs!attname & "", Val(rs!attnum), Val(rs!Length), rs!typname & "", ToBool(rs!attnotnull), ToBool(rs!PrimaryKey), rs!Default & "", lStatistics
      rs.MoveNext
    Wend
  End If
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
objServer.iLogEvent "Entering " & App.Title & ":Columns.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
