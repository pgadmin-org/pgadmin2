VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sequences"
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
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objSequence As pgSequence

  For Each objSequence In mCol
    If objSequence.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objSequence
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(dOID As Double, szName As String, szOwner As String, szACL As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim objNewMember As pgSequence
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgSequence
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = dOID
  objNewMember.iName = szName
  objNewMember.iOwner = szOwner
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.ACL = szACL
  
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "SELECT last_value, max_value, min_value, cache_value, increment_by, is_cycled FROM " & fmtID(szNamespace) & "." & fmtID(szName)
  Else
    szSQL = "SELECT last_value, max_value, min_value, cache_value, increment_by, is_cycled FROM " & fmtID(szName)
  End If
  On Error Resume Next
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  On Error GoTo Err_Handler
  If rs.State = adStateOpen Then
    If Not rs.EOF Then
      objNewMember.Maximum = rs!max_value
      objNewMember.Minimum = rs!min_value
      objNewMember.Increment = rs!increment_by
      objNewMember.Cache = rs!cache_value
      If rs!is_cycled = "t" Or rs!is_cycled = 1 Then
        objNewMember.Cycled = True
      Else
        objNewMember.Cycled = False
      End If
    End If
  End If
  If rs.State <> adStateClosed Then rs.Close
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Sequence
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional ByVal Increment As Variant, Optional ByVal Minimum As Variant, Optional ByVal Maximum As Variant, Optional ByVal Start As Variant, Optional ByVal Cache As Variant, Optional ByVal Cycled As Boolean, Optional ByVal Comment As String) As pgSequence
Attribute Add.VB_Description = "Add a new Sequence."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Add(" & QUOTE & Name & QUOTE & ", " & Increment & ", " & Minimum & ", " & Maximum & ", " & Start & ", " & Cache & ", " & Cycled & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgSequence
Dim szSQL As String
  
  'Create a new object
  Set objNewMember = New pgSequence

  'Generate the SQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "CREATE SEQUENCE " & fmtID(szNamespace) & "." & fmtID(Name)
  Else
    szSQL = "CREATE SEQUENCE " & fmtID(Name)
  End If
  
  'Increment
  If Increment <> "" Then szSQL = szSQL & " INCREMENT " & Increment
  If Minimum <> "" Then szSQL = szSQL & " MINVALUE " & Minimum
  If Maximum <> "" Then szSQL = szSQL & " MAXVALUE " & Maximum
  If Start <> "" Then szSQL = szSQL & " START " & Start
  If Cache <> "" Then szSQL = szSQL & " CACHE " & Cache
  If Cycled = True Then szSQL = szSQL & " CYCLE"
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON SEQUENCE " & fmtID(szNamespace) & "." & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON SEQUENCE " & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  If objVersion.VersionNum >= 7.3 Then
    szSQL = SQL_GET_SEQUENCES & " AND relname = '" & dbSZ(Name) & "' AND relnamespace = " & dNamespaceOID & "::oid"
  Else
    szSQL = SQL_GET_SEQUENCES & " AND relname = '" & dbSZ(Name) & "'"
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    iAdd rs!OID, rs!relname & "", rs!seqowner & "", rs!relacl & ""
  End If

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

'Rename
Public Sub Rename(ByVal OldName As String, ByVal NewName As String)
Attribute Rename.VB_Description = "Rename the specified sequence."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

  Dim objSequence As New pgSequence
    
  Set objSequence = mCol(ULEncode(OldName))
  objSequence.Name = NewName
  objSequence.Identifier = NewName
  mCol.Add objSequence, ULEncode(NewName), ULEncode(OldName)
  mCol.Remove ULEncode(OldName)
  
  Exit Sub
Err_Handler:
  objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgSequence
Attribute Item.VB_Description = "Returns a Sequence object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Sequences in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objSequence As pgSequence
Dim lCount As Long

  If IgnoreSys Then
    For Each objSequence In mCol
      If Not objSequence.SystemObject Then lCount = lCount + 1
    Next objSequence
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Sequence."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Sequence.
  szSQL = "DROP SEQUENCE " & mCol(ULEncode(Key)).FormattedID
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
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Sequences."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_SEQUENCES & " AND relnamespace = " & dNamespaceOID & "::oid ORDER BY relname", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_SEQUENCES & " ORDER BY relname", cnDatabase)
    End If
    While Not rs.EOF
      iAdd rs!OID, rs!relname & "", rs!seqowner & "", rs!relacl & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean, Optional Reset As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Property Get SQL(" & IncludeSystem & ", " & Reset & ")", etFullDebug
Dim szSQL As String
Dim objSequence As pgSequence
  
  szSQL = "-- Sequences" & vbCrLf & vbCrLf

  For Each objSequence In mCol
    If IncludeSystem Or (Not objSequence.SystemObject) Then szSQL = szSQL & objSequence.SQL(Reset) & vbCrLf
  Next objSequence
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
