VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private cnDatabase As Connection
Private szName As String
Private szValue As String
Private mCol As Collection
Private vTag As Variant

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objVar As pgVar

  For Each objVar In mCol
    If objVar.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objVar
  
  Exit Function
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Function

'Internal Add
Friend Sub iAdd(szName As String, szValue)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.iAdd(" & QUOTE & szName & QUOTE & ")", etFullDebug
  
Dim objVar As New pgVar

  'Add the User
  objVar.Name = szName
  objVar.Value = szValue
  objVar.Identifier = szName
  
  mCol.Add objVar, ULEncode(objVar.Identifier)
  
  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

'External Add/Update
Public Sub AddOrUpdate(ByVal Name As String, ByVal Value As String)
Attribute AddOrUpdate.VB_Description = "Add a new Database Variable."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.AddOrUpdate(" & QUOTE & Name & QUOTE & ", " & QUOTE & Value & QUOTE & ")", etFullDebug

Dim objVar As pgVar
Dim szSQL As String
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 11, App.Title, "Database variables can only exist in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate the SQL query to create the new object and execute
  szSQL = "ALTER DATABASE " & fmtID(Database) & " SET " & fmtID(Name) & " = " & fmtID(Value)
  objServer.ExecSQL szSQL, cnDatabase

  'If the var already exists, then update it
  If Exists(Name) Then
    Set objVar = mCol(ULEncode(Name))
    objVar.Value = Value
  Else
    iAdd Name, Value
  End If
  
  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

'Return the specified Item
Public Property Get Item(Key As String) As pgVar
Attribute Item.VB_Description = "Returns a variable object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the variable objects."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified variable."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 11, App.Title, "Database variables can only exist in PostgreSQL 7.3+."
    Exit Sub
  End If

  'Generate/Execute SQL to drop the Var.
  szSQL = "ALTER DATABASE " & fmtID(Database) & " RESET " & fmtID(mCol(ULEncode(Key)).Name)
  objServer.ExecSQL szSQL, cnDatabase

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
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Refresh()", etFullDebug

Dim szSQL As String
Dim szDatVarList As String
Dim rs As New Recordset

  If objVersion.VersionNum >= 7.3 Then
    Set rs = objServer.ExecSQL("SELECT datconfig FROM pg_database WHERE datname = '" & dbSZ(Database) & "'", objServer.Connection)
    If Not rs.EOF Then szDatVarList = rs!datconfig & ""
    If Left(szDatVarList, 1) = "{" Then szDatVarList = Mid(szDatVarList, 2)
    If Right(szDatVarList, 1) = "}" Then szDatVarList = Mid(szDatVarList, 1, Len(szDatVarList) - 1)
    Initialize szDatVarList
  End If
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize(Optional szVarList As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Initialize(" & QUOTE & szVarList & QUOTE & ")", etFullDebug

Dim szDatabaseVars() As String
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
  
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":DatabaseVars.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub


