VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Indexes"
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

Dim lTableOID As Long
Dim szTable As String

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(lOID As Long, szName As String, bUnique As Boolean, bPrimary As Boolean, szDefinition As String) As pgIndex
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & bUnique & ", " & bPrimary & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgIndex

  'Create a new object
  Set objNewMember = New pgIndex
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Table = szTable
  objNewMember.Unique = bUnique
  objNewMember.Primary = bPrimary
  objNewMember.Definition = szDefinition
  objNewMember.Identifier = szName
  objNewMember.TableOID = lTableOID
  Set objNewMember.Connection = cnDatabase
  
  'Indexes are a pain in the butt because of the wierd way in which we must
  'define whether or not they are a system object. Baiscally we must create
  'pgIndex, and then ask it if it is a system object before adding it to
  'the collection
  
  If objServer.IncludeSys Then 'Show whatever...
    mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Else 'Only show non system...
    If Not objNewMember.SystemObject Then mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  End If
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal Unique As Boolean, ByVal ColumnList As String, Optional ByVal AccessMethod As String, Optional ByVal Comment As String, Optional ByVal Constraint As String) As pgIndex
Attribute Add.VB_Description = "Add a new Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Add(" & QUOTE & Name & QUOTE & ", " & Unique & ", " & QUOTE & ColumnList & QUOTE & ", " & QUOTE & Comment & QUOTE & ", " & QUOTE & Constraint & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objIndex As pgIndex
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE"
  If Unique Then szSQL = szSQL & " UNIQUE"
  szSQL = szSQL & " INDEX " & QUOTE & Name & QUOTE
  szSQL = szSQL & " ON " & QUOTE & szTable & QUOTE
  If AccessMethod <> "" Then szSQL = szSQL & " USING " & dbSZ(AccessMethod)
  szSQL = szSQL & " (" & dbSZ(ColumnList) & ")"
  'Constraint is only valid for PostgreSQL 7.2+
  If (Constraint <> "") And (objServer.dbVersion.VersionNum >= 7.2) Then
    szSQL = szSQL & " WHERE " & Constraint
  End If
  objServer.ExecSQL szSQL, cnDatabase
  
  If Comment <> "" Then
    szSQL = "COMMENT ON INDEX " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_INDEXES & " AND relname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objIndex = iAdd(rs!OID, rs!relname & "", ToBool(rs!indisunique), ToBool(rs!indisprimary), rs!Definition & "")
  End If
  
  If Not objIndex.SystemObject Then objIndex.Commit rlAdd, "Created object."
  Set Add = objIndex
  
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
Public Property Get Item(Key As String) As pgIndex
Attribute Item.VB_Description = "Returns an Index object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Indexes on the current Table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objIndex As pgIndex
Dim lCount As Long

  If IgnoreSys Then
    For Each objIndex In mCol
      If Not objIndex.SystemObject Then lCount = lCount + 1
    Next objIndex
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Index."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Rule.
  szSQL = "DROP INDEX " & QUOTE & mCol(ULEncode(Key)).Name & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rlDelete, "Deleted object."
  
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
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Indexes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_INDEXES & " AND indrelid = " & lTableOID & " ORDER BY relname", cnDatabase)
    While Not rs.EOF
      iAdd rs!OID, rs!relname & "", ToBool(rs!indisunique), ToBool(rs!indisprimary), rs!Definition & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Indexes.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objIndex As pgIndex
  
  szSQL = "-- Indexes" & vbCrLf & vbCrLf

  For Each objIndex In mCol
    If IncludeSystem Or (Not objIndex.SystemObject) Then szSQL = szSQL & objIndex.SQL & vbCrLf
  Next objIndex
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

