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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

'Rollback to a prior version
Public Sub Rollback(Identifier As String, Version As Long)
Attribute Rollback.VB_Description = "Rollback the specified object to the specified version."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

Dim szDefinition As String
Dim objItem As rcItem

  'Get the definition to restore to first as it will vanish as soon as we
  'drop the current version
  For Each objItem In mCol(ULEncode(Identifier)).History
    If objItem.Version = Version Then
      szDefinition = objItem.Definition
      Exit For
    End If
  Next objItem
  If szDefinition = "" Then Exit Sub
  
  'Remove the current version
  Remove Identifier
  
  'Recreate the new object.
  objServer.Databases(Database).Execute szDefinition
  
  'Kill the Comment Cache and log the change
  objServer.Databases(Database).KillCommentCache
  Refresh
  mCol(ULEncode(Identifier)).Commit rcAdd, "Object rolled back to version " & Version & "."
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

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

  Database = GetDatabase(cnDatabase.ConnectionString)
  
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
Private Sub iAdd(dTimeStamp As Date, lOID As Long, szName As String, szOwner As String, szACL As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Sequences.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim objNewMember As pgSequence
Dim szSQL As String
Dim rs As New Recordset

  'Create a new object
  Set objNewMember = New pgSequence
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.ACL = szACL
  
  szSQL = "SELECT last_value, max_value, min_value, cache_value, increment_by, is_cycled FROM " & QUOTE & szName & QUOTE
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
  szSQL = "CREATE SEQUENCE " & QUOTE & Name & QUOTE
  
  'Increment
  If Increment <> "" Then szSQL = szSQL & " INCREMENT " & Increment
  If Minimum <> "" Then szSQL = szSQL & " MINVALUE " & Minimum
  If Maximum <> "" Then szSQL = szSQL & " MAXVALUE " & Maximum
  If Start <> "" Then szSQL = szSQL & " START " & Start
  If Cache <> "" Then szSQL = szSQL & " CACHE " & Cache
  If Cycled = True Then szSQL = szSQL & " CYCLE"
  objServer.ExecSQL szSQL, cnDatabase
  If Comment <> "" Then
    szSQL = "COMMENT ON SEQUENCE " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_SEQUENCES & " AND relname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    iAdd rs!ts, rs!OID, rs!relname & "", rs!seqowner & "", rs!relacl & ""
  End If
  
  If Not mCol(ULEncode(Name)).SystemObject Then mCol(ULEncode(Name)).Commit rcAdd, "Created object."
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
  szSQL = "DROP SEQUENCE " & QUOTE & dbSZ(mCol(ULEncode(Key)).Name) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rcDelete, "Deleted object."
  
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
    Set rs = objServer.ExecSQL(SQL_GET_SEQUENCES & " ORDER BY relname", cnDatabase)
    While Not rs.EOF
      iAdd rs!ts, rs!OID, rs!relname & "", rs!seqowner & "", rs!relacl & ""
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
