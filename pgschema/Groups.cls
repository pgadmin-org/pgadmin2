VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
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

'Internal Add
Private Sub iAdd(lOID As Long, szName As String, lID As Long, szMemberList As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & lID & ", " & QUOTE & szMemberList & QUOTE & ")", etFullDebug

Dim objNewMember As pgGroup

  'Create a new object
  Set objNewMember = New pgGroup
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.ID = lID
  objNewMember.MemberList = szMemberList
  objNewMember.MemberList = Replace(objNewMember.MemberList, "{", "")
  objNewMember.MemberList = Replace(objNewMember.MemberList, "}", "")
  objNewMember.Identifier = szName
  
  'Add the Group
  mCol.Add objNewMember, objNewMember.Identifier
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional ByVal ID As Long) As pgGroup
Attribute Add.VB_Description = "Add a new user group."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Add(" & QUOTE & Name & QUOTE & ", " & ID & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgGroup
Dim szSQL As String
  
  'Create a new object
  Set objNewMember = New pgGroup

  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE GROUP " & QUOTE & Name & QUOTE

  If Not IsMissing(ID) Then szSQL = szSQL & " WITH SYSID " & ID
  objServer.ExecSQL szSQL, objServer.Connection
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_GROUPS & " WHERE groname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, objServer.Connection)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    objNewMember.OID = rs!OID
    objNewMember.Name = rs!groname & ""
    objNewMember.ID = rs!grosysid
    objNewMember.MemberList = rs!grolist & ""
    objNewMember.MemberList = Replace(objNewMember.MemberList, "{", "")
    objNewMember.MemberList = Replace(objNewMember.MemberList, "}", "")
    objNewMember.Identifier = Name
    mCol.Add objNewMember, Name
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
  objServer.iLogError Err
  
End Function

'Return the specified Item
Public Property Get Item(IndexOrKey As Variant) As pgGroup
Attribute Item.VB_Description = "Returns a Group item."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Property Get Item(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

  Set Item = mCol(IndexOrKey)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Groups on the current server."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal IndexOrKey As Variant)
Attribute Remove.VB_Description = "Remove the specified Group."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Remove(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Group.
  szSQL = "DROP GROUP " & QUOTE & dbSZ(mCol(IndexOrKey).Name) & QUOTE
  objServer.ExecSQL szSQL, objServer.Connection
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove IndexOrKey
  
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
objServer.iLogEvent "Entering Groups.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_GROUPS & " ORDER BY groname", objServer.Connection)
    While Not rs.EOF
      iAdd rs!OID, rs!groname & "", rs!grosysid, rs!grolist & ""
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
objServer.iLogEvent "Entering Groups.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering Groups.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim objGroup As pgGroup
  
  szSQL = "-- Groups" & vbCrLf & vbCrLf

  For Each objGroup In mCol
    szSQL = szSQL & objGroup.SQL & vbCrLf
  Next objGroup
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property


