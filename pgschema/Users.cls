VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
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
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Users.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objUser As pgUser

  For Each objUser In mCol
    If objUser.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objUser
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Internal Add
Friend Sub iAdd(szName As String, lID As Long, bCreateDatabases As Boolean, bSuperuser As Boolean, bUpdateCatalogues As Boolean, dAccountExpires As Variant, Optional szUserVarList As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.iAdd(" & QUOTE & szName & QUOTE & ", " & lID & ", " & bCreateDatabases & ", " & bSuperuser & ", " & bUpdateCatalogues & ", " & QUOTE & dAccountExpires & QUOTE & ")", etFullDebug

Dim objNewMember As pgUser

  'Create a new object
  Set objNewMember = New pgUser
  objNewMember.Name = szName
  objNewMember.ID = lID
  objNewMember.iCreateDatabases = bCreateDatabases
  objNewMember.iSuperuser = bSuperuser
  objNewMember.UpdateCatalogues = bUpdateCatalogues
  objNewMember.iAccountExpires = dAccountExpires
  If Left(szUserVarList, 1) = "{" Then szUserVarList = Mid(szUserVarList, 2)
  If Right(szUserVarList, 1) = "}" Then szUserVarList = Mid(szUserVarList, 1, Len(szUserVarList) - 1)
  objNewMember.UserVarList = szUserVarList
  objNewMember.Identifier = szName
  
  'Add the User
  mCol.Add objNewMember, ULEncode(szName)
  Set objNewMember = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'External Add
Public Function Add(ByVal Name As String, Optional ByVal ID As Long, Optional ByVal Password As String, Optional ByVal CreateDatabases As Boolean, Optional ByVal Superuser As Boolean, Optional ByVal AccountExpires As Date) As pgUser
Attribute Add.VB_Description = "Add a new User."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Add(" & QUOTE & Name & QUOTE & ", " & ID & ", " & QUOTE & Password & QUOTE & ", " & CreateDatabases & ", " & Superuser & ", " & QUOTE & AccountExpires & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgUser
Dim szSQL As String
  
  'Create a new object
  Set objNewMember = New pgUser

  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE USER " & fmtID(Name)
  If (Not IsMissing(Password)) Or (Not IsMissing(ID)) Then
    szSQL = szSQL & " WITH "
  End If
  If Not IsMissing(ID) Then szSQL = szSQL & " SYSID " & ID
  
  'Password - PostgreSQL 7.2+ can encrypt passwords.
  If Not IsMissing(Password) Then
    If objVersion.VersionNum >= 7.2 Then
      If objServer.EncryptPasswords Then
        szSQL = szSQL & " ENCRYPTED PASSWORD '" & dbSZ(Password) & "'"
      Else
        szSQL = szSQL & " UNENCRYPTED PASSWORD '" & dbSZ(Password) & "'"
      End If
    Else
      szSQL = szSQL & " PASSWORD '" & dbSZ(Password) & "'"
    End If
  End If

  If CreateDatabases = True Then
    szSQL = szSQL & " CREATEDB"
  Else
    szSQL = szSQL & " NOCREATEDB"
  End If
  If Superuser = True Then
    szSQL = szSQL & " CREATEUSER"
  Else
    szSQL = szSQL & " NOCREATEUSER"
  End If
  If Not IsMissing(AccountExpires) Then szSQL = szSQL & " VALID UNTIL '" & Format(AccountExpires, "yyyy-MM-dd hh:mm:ss") & "'"
  objServer.ExecSQL szSQL, objServer.Connection
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_USERS & " WHERE usename = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, objServer.Connection)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    objNewMember.Name = rs!usename & ""
    objNewMember.ID = rs!usesysid
    objNewMember.iCreateDatabases = ToBool(rs!usecreatedb)
    objNewMember.iSuperuser = ToBool(rs!usesuper)
    objNewMember.UpdateCatalogues = ToBool(rs!usecatupd)
    If IsDate(rs!valuntil) Then
      objNewMember.iAccountExpires = rs!valuntil
    Else
      objNewMember.iAccountExpires = "2001-12-31 23:59:59"
    End If
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
Public Property Get Item(Key As String) As pgUser
Attribute Item.VB_Description = "Returns a User object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the User objects."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified user."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the User.
  szSQL = "DROP USER " & mCol(ULEncode(Key)).FormattedID
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
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Users."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
    
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_USERS & " ORDER BY usename", objServer.Connection)
    While Not rs.EOF
      If objVersion.VersionNum >= 7.3 Then
        If IsDate(rs!valuntil) Then
          iAdd rs!usename & "", rs!usesysid, ToBool(rs!usecreatedb), ToBool(rs!usesuper), ToBool(rs!usecatupd), rs!valuntil, rs!useconfig & ""
        Else
          iAdd rs!usename & "", rs!usesysid, ToBool(rs!usecreatedb), ToBool(rs!usesuper), ToBool(rs!usecatupd), "2099-12-31 23:59:59", rs!useconfig & ""
        End If
      Else
        If IsDate(rs!valuntil) Then
          iAdd rs!usename & "", rs!usesysid, ToBool(rs!usecreatedb), ToBool(rs!usesuper), ToBool(rs!usecatupd), rs!valuntil
        Else
          iAdd rs!usename & "", rs!usesysid, ToBool(rs!usecreatedb), ToBool(rs!usesuper), ToBool(rs!usecatupd), "2099-12-31 23:59:59"
        End If
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

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Users.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim objUser As pgUser
  
  szSQL = "-- Users" & vbCrLf & vbCrLf

  For Each objUser In mCol
    szSQL = szSQL & objUser.SQL & vbCrLf
  Next objUser
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
