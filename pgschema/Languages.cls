VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Languages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objLanguage As pgLanguage

  For Each objLanguage In mCol
    If objLanguage.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objLanguage
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(lOID As Long, szName As String, bTrusted As Boolean, lHandler As Long, szComment As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & bTrusted & ", " & lHandler & ", " & QUOTE & szComment & QUOTE & ")", etFullDebug

Dim objNewMember As pgLanguage
Dim rs As New Recordset
Dim szSQL As String

  'Create a new object
  Set objNewMember = New pgLanguage
  If lOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Sub
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Trusted = bTrusted
  
  'Get the Handler
  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = SQL_GET_FUNCTIONS7_3 & " AND oid = " & lHandler
  Else
    szSQL = SQL_GET_FUNCTIONS7_1 & " WHERE oid = " & lHandler
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    objNewMember.Handler = rs!proname & ""
  End If
  If rs.State <> adStateClosed Then rs.Close

  objNewMember.iComment = szComment
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Language
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set objNewMember = Nothing
  
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'External Add
Public Function Add(ByVal Name As String, ByVal Trusted As Boolean, ByVal Handler As String, Optional ByVal Comment As String) As pgLanguage
Attribute Add.VB_Description = "Add a new Language."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Add(" & QUOTE & Name & QUOTE & ", " & Trusted & ", " & QUOTE & Handler & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgLanguage
Dim szSQL As String
Dim objLanguage As pgLanguage
  
  'Create a new object
  Set objNewMember = New pgLanguage

  'Generate the SQL query to create the new object and execute
  If Trusted = True Then
    szSQL = "CREATE TRUSTED PROCEDURAL LANGUAGE '" & Name & "' HANDLER " & Handler & " LANCOMPILER '" & dbSZ(Comment) & "'"
  Else
    szSQL = "CREATE PROCEDURAL LANGUAGE '" & Name & "' HANDLER " & Handler & " LANCOMPILER '" & dbSZ(Comment) & "'"
  End If
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the SQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_LANGUAGES & " WHERE lanname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  
  'Create a new member of the collection
  If Not rs.EOF Then
    objNewMember.OID = rs!OID
    objNewMember.Name = rs!lanname & ""
    objNewMember.Trusted = ToBool(rs!lanpltrusted)
    objNewMember.Handler = rs!lanplcallfoid & ""
    objNewMember.iComment = rs!lancompiler & ""
    objNewMember.Identifier = Name
    Set objNewMember.Connection = cnDatabase
    mCol.Add objNewMember, ULEncode(Name)
  End If
  
  'Return the object created
  Set objLanguage = objNewMember
  
  'Invalidate the Language Cache...
  objServer.Databases(Database).KillLanguageCache
  
  Set Add = objLanguage
  
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
Public Property Get Item(Key As String) As pgLanguage
Attribute Item.VB_Description = "Returns a Language object."
Attribute Item.VB_UserMemId = 0
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Languages in the current database."
Attribute Count.VB_MemberFlags = "200"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objLanguage As pgLanguage
Dim lCount As Long

  If IgnoreSys Then
    For Each objLanguage In mCol
      If Not objLanguage.SystemObject Then lCount = lCount + 1
    Next objLanguage
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Language."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP PROCEDURAL LANGUAGE '" & mCol(ULEncode(Key)).Name & "'"
  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)

  'Invalidate the Language Cache...
  objServer.Databases(Database).KillLanguageCache
  
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
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Languages."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection

  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_LANGUAGES & " ORDER BY lanname", cnDatabase)
    While Not rs.EOF
      iAdd rs!OID, rs!lanname & "", ToBool(rs!lanpltrusted), rs!lanplcallfoid & "", rs!lancompiler & ""
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
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Languages.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objLanguage As pgLanguage
  
  szSQL = "-- Languages" & vbCrLf & vbCrLf

  For Each objLanguage In mCol
    If IncludeSystem Or (Not objLanguage.SystemObject) Then szSQL = szSQL & objLanguage.SQL & vbCrLf
  Next objLanguage
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
