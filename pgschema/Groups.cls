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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objGroup As pgGroup

  For Each objGroup In mCol
    If objGroup.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objGroup
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Internal Add
Private Sub iAdd(szName As String, lID As Long, szMemberList As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.iAdd(" & QUOTE & szName & QUOTE & ", " & lID & ", " & QUOTE & szMemberList & QUOTE & ")", etFullDebug

Dim objNewMember As pgGroup

  'Create a new object
  Set objNewMember = New pgGroup
  objNewMember.Name = szName
  objNewMember.ID = lID
  objNewMember.MemberList = szMemberList
  objNewMember.MemberList = Replace(objNewMember.MemberList, "{", "")
  objNewMember.MemberList = Replace(objNewMember.MemberList, "}", "")
  objNewMember.Identifier = szName
  
  'Add the Group
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional ByVal ID As Long) As pgGroup
Attribute Add.VB_Description = "Add a new user group."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Add(" & QUOTE & Name & QUOTE & ", " & ID & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgGroup
Dim szSQL As String
  
  'Create a new object
  Set objNewMember = New pgGroup

  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE GROUP " & fmtID(Name)

  If Not IsMissing(ID) Then szSQL = szSQL & " WITH SYSID " & ID
  objServer.ExecSQL szSQL, objServer.Connection
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_GROUPS & " WHERE groname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, objServer.Connection)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    objNewMember.Name = rs!groname & ""
    objNewMember.ID = rs!grosysid
    objNewMember.MemberList = rs!grolist & ""
    objNewMember.MemberList = Replace(objNewMember.MemberList, "{", "")
    objNewMember.MemberList = Replace(objNewMember.MemberList, "}", "")
    objNewMember.Identifier = Name
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
  objServer.iLogError Err.Number, Err.Description
  
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgGroup
Attribute Item.VB_Description = "Returns a Group item."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Groups on the current server."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Group."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Group.
  szSQL = "DROP GROUP " & mCol(ULEncode(Key)).FormattedID
  objServer.ExecSQL szSQL, objServer.Connection
  
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
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreesh the collection."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_GROUPS & " ORDER BY groname", objServer.Connection)
    While Not rs.EOF
      iAdd rs!groname & "", rs!grosysid, rs!grolist & ""
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

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Groups.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim objGroup As pgGroup
  
  szSQL = "-- Groups" & vbCrLf & vbCrLf

  For Each objGroup In mCol
    szSQL = szSQL & objGroup.SQL & vbCrLf
  Next objGroup
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
