VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get Tag()", etFullDebug

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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objTable As pgTable

  For Each objTable In mCol
    If objTable.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objTable
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Sub iAdd(dOID As Double, szName As String, szOwner As String, szACL As String, bHasOIDs As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ", " & bHasOIDs & ")", etFullDebug

Dim objNewMember As pgTable
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgTable
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.Oid = dOID
  objNewMember.iName = szName
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.iOwner = szOwner
  objNewMember.ACL = szACL
  objNewMember.HasOIDs = bHasOIDs
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Language
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'External Add
Public Function Add(ByVal Name As String, ByVal Columns As String, Optional ByVal PrimaryKeys As String, Optional ByVal Checks As String, Optional ByVal ForeignKeys As String, Optional ByVal Inherits As String, Optional ByVal Comment As String, Optional ByVal OIDs As Boolean = True) As pgTable
Attribute Add.VB_Description = "Add a new Table."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Columns & QUOTE & ", " & QUOTE & PrimaryKeys & QUOTE & ", " & QUOTE & Checks & QUOTE & ", " & QUOTE & Checks & QUOTE & ", " & QUOTE & ForeignKeys & QUOTE & ", " & QUOTE & Inherits & QUOTE & ", " & QUOTE & Comment & QUOTE & ", " & OIDs & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the SQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "CREATE TABLE " & fmtID(szNamespace) & "." & fmtID(Name) & " (" & Columns
  Else
    szSQL = "CREATE TABLE " & fmtID(Name) & " (" & Columns
  End If
  If PrimaryKeys <> "" Then szSQL = szSQL & ", PRIMARY KEY (" & PrimaryKeys & ")"
  If Checks <> "" Then szSQL = szSQL & ", " & Checks
  If ForeignKeys <> "" Then szSQL = szSQL & ", " & ForeignKeys
  szSQL = szSQL & ")"
  If Inherits <> "" Then szSQL = szSQL & " INHERITS (" & Inherits & ")"
  If objVersion.VersionNum >= 7.2 Then
    If OIDs Then
      szSQL = szSQL & " WITH OIDS"
    Else
      szSQL = szSQL & " WITHOUT OIDS"
    End If
  End If
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON TABLE " & fmtID(szNamespace) & "." & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON TABLE " & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If

  'Generate and execute the SQL to retrieve *all* properties of the new object
  If objVersion.VersionNum >= 7.3 Then
    szSQL = SQL_GET_TABLES7_2 & " AND relname = '" & Name & "' AND relnamespace = " & dNamespaceOID & "::oid"
  ElseIf objVersion.VersionNum = 7.2 Then
    szSQL = SQL_GET_TABLES7_2 & " AND relname = '" & Name & "'"
  Else
    szSQL = SQL_GET_TABLES7_1 & " AND relname = '" & Name & "'"
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    If objVersion.VersionNum >= 7.2 Then
      iAdd rs.Fields("oid"), rs!relname & "", rs!tableowner & "", rs!relacl & "", ToBool(rs!relhasoids)
    Else
      iAdd rs.Fields("oid"), rs!relname & "", rs!tableowner & "", rs!relacl & "", True
    End If
  End If
  
  Set Add = mCol(ULEncode(Name))
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
  
End Function

'Rename
Public Sub Rename(ByVal OldName As String, ByVal NewName As String)
Attribute Rename.VB_Description = "Rename the specified table."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objTable As New pgTable
    
  Set objTable = mCol(ULEncode(OldName))
  objTable.Name = NewName
  objTable.Identifier = NewName
  mCol.Add objTable, ULEncode(NewName), ULEncode(OldName)
  mCol.Remove ULEncode(OldName)
  
  Exit Sub
Err_Handler:
  objServer.iLogError Err.Number, Err.Description
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgTable
Attribute Item.VB_Description = "Returns a Table object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Tables in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objTable As pgTable
Dim lCount As Long

  If IgnoreSys Then
    For Each objTable In mCol
      If Not objTable.SystemObject Then lCount = lCount + 1
    Next objTable
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String, Optional RemoveSubItems As Boolean)
Attribute Remove.VB_Description = "Remove the specified Table."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim objIndex As pgIndex
Dim objRule As pgRule
Dim objTrigger As pgTrigger

  'Remove subitems first. We must remove Indexes, Rules & Triggers
  If RemoveSubItems Then
    For Each objIndex In mCol(ULEncode(Key)).Indexes
      If Not objIndex.SystemObject Then mCol(ULEncode(Key)).Indexes.Remove objIndex.Identifier
    Next objIndex
    For Each objRule In mCol(ULEncode(Key)).Rules
      If Not objRule.SystemObject Then mCol(ULEncode(Key)).Rules.Remove objRule.Identifier
    Next objRule
    For Each objTrigger In mCol(ULEncode(Key)).Triggers
      If Not objTrigger.SystemObject Then mCol(ULEncode(Key)).Triggers.Remove objTrigger.Identifier
    Next objTrigger
  End If
  
  'Generate/Execute SQL to drop the Table.
  szSQL = "DROP TABLE " & mCol(ULEncode(Key)).FormattedID
  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
  Exit Sub
Err_Handler:
  'Error 5 is thrown if the item is not in the collection...
  If Err.Number = 5 Then Exit Sub
  objServer.iLogError Err.Number, Err.Description
End Sub

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_TABLES7_2 & " AND relnamespace = " & dNamespaceOID & "::oid ORDER BY relname", cnDatabase)
    ElseIf objVersion.VersionNum = 7.2 Then
      Set rs = objServer.ExecSQL(SQL_GET_TABLES7_2 & " ORDER BY relname", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_TABLES7_1 & " ORDER BY relname", cnDatabase)
    End If
      While Not rs.EOF
        If objVersion.VersionNum >= 7.2 Then
          iAdd rs.Fields("oid"), rs!relname & "", rs!tableowner & "", rs!relacl & "", ToBool(rs!relhasoids)
        Else
          iAdd rs.Fields("oid"), rs!relname & "", rs!tableowner & "", rs!relacl & "", True
        End If
      rs.MoveNext
    Wend
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Tables."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Tables.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objTable As pgTable
  
  szSQL = "-- Tables" & vbCrLf & vbCrLf

  For Each objTable In mCol
    If IncludeSystem Or (Not objTable.SystemObject) Then szSQL = szSQL & objTable.SQL & vbCrLf
  Next objTable
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
