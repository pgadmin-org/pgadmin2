VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private szName As String
Private mCol As Collection

'Internal Add
Friend Sub iAdd(szName As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.iAdd(" & QUOTE & szName & QUOTE & ")", etFullDebug
  
  'Add the User
  mCol.Add szName, ULEncode(szName)
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'External Add
Public Sub Add(ByVal Name As String)
Attribute Add.VB_Description = "Add a new User."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Add(" & QUOTE & Name & QUOTE & ")", etFullDebug

Dim X As Long
Dim szSQL As String
  
  'Generate the SQL query to create the new object and execute
  szSQL = "ALTER GROUP " & QUOTE & szName & QUOTE & " ADD USER " & QUOTE & Name & QUOTE
  objServer.ExecSQL szSQL, objServer.Connection
  
  'Quick error check - if the user already exists in the members collection
  'then don't try to add him/her again.
  For X = 1 To mCol.Count
    If mCol(X) = Name Then
      objServer.iLogEvent "Not adding " & Name & " to group " & szName & "'s Members Collection because he/she already exists!", etMiniDebug
      Exit Sub
    End If
  Next X
  
  mCol.Add Name, ULEncode(Name)
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgUser
Attribute Item.VB_Description = "Returns a User object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Name for internal use.
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Property let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the User objects."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified user."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the User.
  szSQL = "ALTER GROUP " & QUOTE & szName & QUOTE & " DROP USER " & QUOTE & dbSZ(mCol(ULEncode(Key))) & QUOTE
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
objServer.iLogEvent "Entering " & App.Title & ":Members.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Initialise the collection if possible
Friend Sub Initialize(Optional szUserList As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Initialize(" & QUOTE & szUserList & QUOTE & ")", etFullDebug

Dim szMembers() As String
Dim szWhere As String
Dim X As Integer
Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Build a WHERE clause from any userlist
  If szUserList <> "" Then
    szMembers = Split(szUserList, ",")
    szWhere = " WHERE "
    For X = 0 To UBound(szMembers)
      szWhere = szWhere & "usesysid = " & szMembers(X) & " OR "
    Next X
    szWhere = Mid(szWhere, 1, Len(szWhere) - 3)
  
    'Get the data and loop through it
    If objServer.Connection.State = adStateOpen Then
      Set rs = objServer.ExecSQL(SQL_GET_USERS & szWhere, objServer.Connection)
      While Not rs.EOF
        If IsDate(rs!valuntil) Then
          iAdd rs!usename & ""
        Else
          iAdd rs!usename & ""
        End If
        rs.MoveNext
      Wend
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

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Members.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub


