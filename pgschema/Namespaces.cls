VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Namespaces"
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
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objNamespace As pgNamespace

  For Each objNamespace In mCol
    If objNamespace.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objNamespace
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(lOID As Long, szName As String, szOwner As String, szACL As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim objNewMember As pgNamespace
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgNamespace
  If lOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.ACL = szACL
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Namespace
  mCol.Add objNewMember, ULEncode(szName)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional Owner As String, Optional ByVal Comment As String) As pgNamespace
Attribute Add.VB_Description = "Add a new Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  If objServer.dbVersion.VersionNum < 7.3 Then
    Err.Raise 5432006, App.Title, "You cannot create schemas in PostgreSQL versions prior to 7.3."
    Exit Function
  End If
  
  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE SCHEMA " & fmtID(Name)
  If Owner <> "" Then szSQL = szSQL & " AUTHORIZATION " & fmtID(Owner)
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    szSQL = "COMMENT ON SCHEMA " & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If

  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_NAMESPACES & " WHERE nspname = '" & Name & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    iAdd rs!OID, rs!nspname & "", rs!namespaceowner & "", rs!nspacl & ""
  End If
  
  'Return the object created
  Set Add = mCol(ULEncode(Name))
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgNamespace
Attribute Item.VB_Description = "Returns a Namespace object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Namespaces in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objNamespace As pgNamespace
Dim lCount As Long

  If IgnoreSys Then
    For Each objNamespace In mCol
      If Not objNamespace.SystemObject Then lCount = lCount + 1
    Next objNamespace
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objServer.dbVersion.VersionNum < 7.3 Then
    Err.Raise 5432005, App.Title, "You cannot drop dummy schemas in PostgreSQL versions prior to 7.3."
    Exit Sub
  End If
  
  'Generate/Execute SQL to drop the Namespace.
  szSQL = "DROP SCHEMA " & mCol(ULEncode(Key)).FormattedID
  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)
  
  Exit Sub
Err_Handler:
  'Error 5 is thrown if the item is not in the collection...
  If Err.Number = 5 Then Exit Sub
  objServer.iLogError Err
End Sub

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objServer.dbVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_NAMESPACES & " ORDER BY nspname", cnDatabase)
      While Not rs.EOF
        iAdd rs!OID, rs!nspname & "", rs!namespaceowner, rs!nspacl & ""
        rs.MoveNext
      Wend
    Else
      'Add a dummy schema for PostgreSQL < 7.3
      iAdd 0, "public", "postgres", "UC"
    End If
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Namespaces."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Refresh()", etFullDebug

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

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Namespaces.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objNamespace As pgNamespace
  
  szSQL = "-- Schemas" & vbCrLf & vbCrLf

  For Each objNamespace In mCol
    If IncludeSystem Or (Not objNamespace.SystemObject) Then szSQL = szSQL & objNamespace.SQL & vbCrLf
  Next objNamespace
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

