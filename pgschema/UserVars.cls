VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private szUser As String
Private szName As String
Private szValue As String
Private mCol As Collection
Private vTag As Variant

'Set the User
Friend Property Let User(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Let User(" & vData & ")", etFullDebug

  szUser = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler: objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objVar As pgVar

  For Each objVar In mCol
    If objVar.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objVar
  
  Exit Function
Err_Handler: objServer.iLogError Err
End Function

'Internal Add
Friend Sub iAdd(szName As String, szValue)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.iAdd(" & QUOTE & szName & QUOTE & ")", etFullDebug
  
Dim objVar As New pgVar

  'Add the User
  objVar.Name = szName
  objVar.Value = szValue
  objVar.Identifier = szName
  
  mCol.Add objVar, ULEncode(objVar.Identifier)
  
  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

'External Add/Update
Public Sub AddOrUpdate(ByVal Name As String, ByVal Value As String)
Attribute AddOrUpdate.VB_Description = "Add/Update a User Variable."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.AddOrUpdate(" & QUOTE & Name & QUOTE & ", " & QUOTE & Value & QUOTE & ")", etFullDebug

Dim objVar As pgVar
Dim szSQL As String
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 10, App.Title, "User variables can only exist in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate the SQL query to create the new object and execute
  szSQL = "ALTER USER " & fmtID(szUser) & " SET " & fmtID(Name) & " = " & dbSZ(Value)
  objServer.ExecSQL szSQL, objServer.Connection

  'If the var already exists, then update it
  If Exists(Name) Then
    Set objVar = mCol(ULEncode(Name))
    objVar.Value = Value
  Else
    iAdd Name, Value
  End If
  
  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgVar
Attribute Item.VB_Description = "Returns a User Variable object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the User Variable  objects."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler: objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified User Variable."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 10, App.Title, "User variables can only exist in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate/Execute SQL to drop the User.
  szSQL = "ALTER USER " & fmtID(szUser) & " RESET " & fmtID(mCol(ULEncode(Key)).Name)
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
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Refresh()", etFullDebug

Dim szSQL As String
Dim szUserVarList As String
Dim rs As New Recordset

  If objVersion.VersionNum >= 7.3 Then
    Set rs = objServer.ExecSQL("SELECT useconfig FROM pg_user WHERE usename = '" & dbSZ(szUser) & "'", objServer.Connection)
    If Not rs.EOF Then szUserVarList = rs!useconfig & ""
    If Left(szUserVarList, 1) = "{" Then szUserVarList = Mid(szUserVarList, 2)
    If Right(szUserVarList, 1) = "}" Then szUserVarList = Mid(szUserVarList, 1, Len(szUserVarList) - 1)
    Initialize szUserVarList
  End If
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize(Optional szVarList As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Initialize(" & QUOTE & szVarList & QUOTE & ")", etFullDebug

Dim szUserVars() As String
Dim szChar As String
Dim X As Integer
Dim Y As Integer
Dim bInQuotes As Boolean
Dim bVal As String
Dim szVarName As String
Dim szVarVal As String

  'Create a new collection
  Set mCol = New Collection
  
  'Build a WHERE clause from any userlist
  bInQuotes = False
  bVal = False
  Y = Len(szVarList)
  For X = 1 To Y
    szChar = Mid(szVarList, X, 1)
    If szChar = QUOTE Then
      bInQuotes = Not bInQuotes
      If X = Y Then iAdd Trim(szVarName), Trim(szVarVal) 'Quote as last char
    ElseIf szChar = "=" Then
      bVal = True
    ElseIf (szChar = ",") And (Not bInQuotes) Then 'Comma delimiter
      iAdd Trim(szVarName), Trim(szVarVal)
      szVarName = ""
      szVarVal = ""
      bVal = False
    ElseIf X = Y Then 'Last character
      If bVal Then
        szVarVal = szVarVal & szChar
      Else
        szVarName = szVarName & szChar
      End If
      iAdd Trim(szVarName), Trim(szVarVal)
    Else
      If bVal Then
        szVarVal = szVarVal & szChar
      Else
        szVarName = szVarName & szChar
      End If
    End If
  Next X

  Exit Sub
  
Err_Handler: objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":UserVars.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub


