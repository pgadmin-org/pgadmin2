VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Databases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private mCol As Collection

'Internal Add
Private Sub iAdd(lOID As Long, szName As String, szPath As String, szOwner As String, lEncoding As Long, szEncoding As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szPath & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & lEncoding & ", " & szEncoding & ")", etFullDebug

Dim objNewMember As pgDatabase

  'Create a new object
  Set objNewMember = New pgDatabase
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName

  objNewMember.Path = szPath
  objNewMember.Owner = szOwner
  objNewMember.Encoding = lEncoding
  objNewMember.EncodingName = szEncoding
  objNewMember.Identifier = szName
  
  'Add the Database
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional ByVal Path As String, Optional Encoding As String, Optional ByVal Comment As String) As pgDatabase
Attribute Add.VB_Description = "Add a new database object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Path & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim rsComment As New Recordset
Dim rsUser As New Recordset
Dim objNewMember As pgDatabase
Dim szSQL As String
  
  'Create a new object
  Set objNewMember = New pgDatabase

  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE DATABASE " & QUOTE & Name & QUOTE & " WITH"
  If Path <> "" Then szSQL = szSQL & " LOCATION = '" & dbSZ(Path) & "'"
  szSQL = szSQL & " TEMPLATE = template0"
  If Encoding <> "" Then szSQL = szSQL & " ENCODING = '" & dbSZ(Encoding) & "'"
  objServer.ExecSQL szSQL, objServer.Connection
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_DATABASES & " WHERE datname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, objServer.Connection)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    objNewMember.OID = rs!OID
    objNewMember.Name = rs!datname & ""
    objNewMember.Path = rs!datpath & ""

    'Get the user info for the user object
    szSQL = SQL_GET_USERS & " WHERE usename = '" & dbSZ(rs!datowner & "") & "'"
    Set rsUser = objServer.ExecSQL(szSQL, objServer.Connection)
    If Not rsUser.EOF Then
      objNewMember.Owner = rsUser!usename & ""
    End If
   
    objNewMember.Encoding = rs!Encoding
    objNewMember.EncodingName = rs!EncodingName
    objNewMember.Identifier = Name
    
    'Set the Comment using the new object. This will ensure that the comment
    'is created in the new database
    If Comment <> "" Then objNewMember.Comment = Comment
    
    mCol.Add objNewMember, ULEncode(Name)
  End If
  
  'Return the object created
  Set Add = objNewMember
  
  'Cleanup
  Set objNewMember = Nothing
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  If rsComment.State <> adStateClosed Then rsComment.Close
  Set rsComment = Nothing
  objServer.iLogError Err
  
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgDatabase
Attribute Item.VB_Description = "Returns  a database object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of Database objects."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objDatabase As pgDatabase
Dim lCount As Long

  If IgnoreSys Then
    For Each objDatabase In mCol
      If Not objDatabase.SystemObject Then lCount = lCount + 1
    Next objDatabase
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Database object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim objDatabase As pgDatabase

  Set objDatabase = mCol(ULEncode(Key))
  
  'Generate/Execute SQL to drop the database.
  'The leading space fools PostgreSQL into letting us work in a transaction
  szSQL = " DROP DATABASE " & QUOTE & objDatabase.Name & QUOTE
  
  'If we are still connected to the Database this will definately fail
  'so destroy the Database Connection first.
  objDatabase.dbDisconnect
  Set objDatabase = Nothing
  
  'Drop the database
  objServer.ExecSQL szSQL, objServer.Connection
  
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
objServer.iLogEvent "Entering " & App.Title & ":Databases.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Databases."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Initialize()", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_DATABASES & " WHERE datname != 'template0' ORDER BY datname", objServer.Connection)
    While Not rs.EOF
      iAdd rs!OID, rs!datname & "", rs!datpath & "", rs!datowner & "", rs!Encoding, rs!EncodingName
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
objServer.iLogEvent "Entering " & App.Title & ":Databases.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Databases.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objDatabase As pgDatabase
  
  szSQL = "-- Databases" & vbCrLf & vbCrLf

  For Each objDatabase In mCol
    If IncludeSystem Or (Not objDatabase.SystemObject) Then szSQL = szSQL & objDatabase.SQL & vbCrLf
  Next objDatabase
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
