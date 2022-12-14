VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Collections
Private colUserVars As UserVars

'Properties
Private szName As String
Private szIdentifier As String
Private lID As Long
Private bCreateDatabases As Boolean
Private bSuperuser As Boolean
Private bUpdateCatalogues As Boolean
Private dAccountExpires As Date
Private vTag As Variant
Private szUserVarList As String

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Username."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get ObjectType()", etFullDebug

  ObjectType = "User"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return ID
Friend Property Let ID(ByVal vData As Long)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let ID(" & vData & ")", etFullDebug

  lID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns the User ID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get ID()", etFullDebug

  ID = lID
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return CreateDatabases
Friend Property Let iCreateDatabases(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let iCreateDatabases(" & vData & ")", etFullDebug

  bCreateDatabases = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let CreateDatabases(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let CreateDatabases(" & vData & ")", etFullDebug

Dim szSQL As String

  If vData = True Then
    szSQL = "ALTER USER " & FormattedID & " CREATEDB"
  Else
    szSQL = "ALTER USER " & FormattedID & " NOCREATEDB"
  End If
  objServer.ExecSQL szSQL, objServer.Connection
  bCreateDatabases = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get CreateDatabases() As Boolean
Attribute CreateDatabases.VB_Description = "Returns a value indicating whether or not the User can create databases."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get CreateDatabases()", etFullDebug

  CreateDatabases = bCreateDatabases
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return Superuser
Friend Property Let iSuperuser(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let iSuperuser(" & vData & ")", etFullDebug
  
  bSuperuser = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let Superuser(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let Superuser(" & vData & ")", etFullDebug

Dim szSQL As String
 
  If vData = True Then
    szSQL = "ALTER USER " & FormattedID & " CREATEUSER"
  Else
    szSQL = "ALTER USER " & FormattedID & " NOCREATEUSER"
  End If
  objServer.ExecSQL szSQL, objServer.Connection
  bSuperuser = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Superuser() As Boolean
Attribute Superuser.VB_Description = "Returns a value indicating whether or not the User is a superuser."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get Superuser()", etFullDebug

  Superuser = bSuperuser
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return UpdateCatalogues
Friend Property Let UpdateCatalogues(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let UpdateCatalogues(" & vData & ")", etFullDebug

  bUpdateCatalogues = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get UpdateCatalogues() As Boolean
Attribute UpdateCatalogues.VB_Description = "Returns a value indicating whether or not the User can update system catalogues."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get UpdateCatalogues()", etFullDebug

  UpdateCatalogues = bUpdateCatalogues
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Expiry Date
Friend Property Let iAccountExpires(ByVal vData As Date)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let iAccountExpires(" & QUOTE & vData & QUOTE & ")", etFullDebug

  dAccountExpires = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Let AccountExpires(ByVal vData As Date)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let AccountExpires(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim szSQL As String

  szSQL = "ALTER USER " & FormattedID & " VALID UNTIL '" & Format(vData, "yyyy-MM-dd hh:mm:ss") & "'"
  objServer.ExecSQL szSQL, objServer.Connection
  dAccountExpires = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get AccountExpires() As Date
Attribute AccountExpires.VB_Description = "Returns the date/time the the User account expires."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get AccountExpires()", etFullDebug

  AccountExpires = dAccountExpires
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the user's password. Note that we always return ********
Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "Set's the User's password. This Property always returns ********."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Let Password(" & QUOTE & "********" & QUOTE & ")", etFullDebug

  If objVersion.VersionNum >= 7.2 Then
    If objServer.EncryptPasswords Then
      objServer.ExecSQL "ALTER USER " & FormattedID & " WITH ENCRYPTED PASSWORD '" & dbSZ(vData) & "'", objServer.Connection
    Else
      objServer.ExecSQL "ALTER USER " & FormattedID & " WITH UNENCRYPTED PASSWORD '" & dbSZ(vData) & "'", objServer.Connection
    End If
  Else
    objServer.ExecSQL "ALTER USER " & FormattedID & " WITH PASSWORD '" & dbSZ(vData) & "'", objServer.Connection
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Password() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get Password()", etFullDebug

  Password = "********"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the UserVars Collection
Public Property Get UserVars() As UserVars
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get UserVars()", etFullDebug

  If colUserVars Is Nothing Then
    Set colUserVars = New UserVars
    colUserVars.Initialize szUserVarList
    colUserVars.User = szName
  End If
  Set UserVars = colUserVars
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the UserVarList
Friend Property Let UserVarList(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let UserVarList(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szUserVarList = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Get UserVarList() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get UserVarList()", etFullDebug

  UserVarList = szUserVarList
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgUser.SQL()", etFullDebug

Dim szSQL As String
Dim szMember As Variant
Dim objVar As pgVar
  
  szSQL = "-- User: " & szName & vbCrLf & _
          "CREATE USER " & FormattedID & " WITH SYSID " & lID
          
  If bCreateDatabases Then
    szSQL = szSQL & " CREATEDB"
  Else
    szSQL = szSQL & " NOCREATEDB"
  End If
  If bSuperuser Then
    szSQL = szSQL & " CREATEUSER"
  Else
    szSQL = szSQL & " NOCREATEUSER"
  End If
  szSQL = szSQL & " VALID UNTIL '" & Format(dAccountExpires, "yyyy-MM-dd hh:mm:ss") & "'"
  szSQL = szSQL & ";" & vbCrLf
  
  'User Vars
  If objVersion.VersionNum >= 7.3 Then
    For Each objVar In UserVars
      szSQL = szSQL & "ALTER USER " & FormattedID & " SET " & objVar.FormattedID & " = " & fmtID(objVar.Value) & ";" & vbCrLf
    Next objVar
  End If

  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Class_Terminate()", etFullDebug

  Set colUserVars = Nothing

  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub
