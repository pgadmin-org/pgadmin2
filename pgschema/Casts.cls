VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Casts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Casts.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objCast As pgCast

  For Each objCast In mCol
    If objCast.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objCast
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(ByVal dOID As Double, ByVal szSource As String, ByVal szTarget As String, ByVal szFunc As String, ByVal szContext)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.iAdd(" & dOID & ", " & QUOTE & szSource & QUOTE & ", " & QUOTE & szTarget & QUOTE & ", " & QUOTE & szFunc & QUOTE & ", " & QUOTE & szContext & QUOTE & ")", etFullDebug

Dim objNewMember As pgCast
Dim rs As New Recordset
Dim szSQL As String

  'Create a new object
  Set objNewMember = New pgCast
  If dOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = dOID
  objNewMember.Source = szSource
  objNewMember.Target = szTarget
  objNewMember.Funct = szFunc
    
  Select Case UCase(szContext)
    Case "I"
      objNewMember.Context = "IMPLICIT"
    Case "E"
      objNewMember.Context = "EXPLICIT"
    Case "A"
      objNewMember.Context = "ASSIGNMENT"
  End Select
  
  objNewMember.Name = Trim(objNewMember.Source) & "::" & Trim(objNewMember.Target)
  objNewMember.Identifier = objNewMember.Name
  Set objNewMember.Connection = cnDatabase
  
  'Add the Cast
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal szSource As String, ByVal szTarget As String, ByVal szFunct As String, ByVal szContext As String) As pgCast
Attribute Add.VB_Description = "Add a new Cast."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Add(" & QUOTE & szSource & QUOTE & ", " & QUOTE & szTarget & QUOTE & ", " & QUOTE & szFunct & QUOTE & ", " & QUOTE & szContext & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim objNewMember As pgCast
Dim szSQL As String
Dim objCast As pgCast
  
  'Create a new object
  Set objNewMember = New pgCast

  'Generate the SQL query to create the new object and execute
  szSQL = "CREATE CAST"
  szSQL = szSQL & " ( " & szSource & " AS " & szTarget & " )" & vbCrLf
  
  If Len(szFunct) > 0 Then
    szSQL = szSQL & " WITH FUNCTION " & szFunct & "(" & szSource & ")"
  Else
    szSQL = szSQL & " WITHOUT FUNCTION" & vbCrLf
  End If
  
  Select Case szContext
    Case "ASSIGNMENT"
      szSQL = szSQL & " AS ASSIGNMENT " & vbCrLf
    Case "EXPLICIT"
    Case "IMPLICIT"
      szSQL = szSQL & " AS IMPLICIT " & vbCrLf
  End Select
          
  objServer.ExecSQL szSQL, cnDatabase
  
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_CASTS & " WHERE t1.typname = '" & szSource & "' AND t2.typname = '" & szTarget & "' AND p.proname = '" & szFunct & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objCast = iAdd(rs.Fields("oid"), rs!castsource & "", rs!casttarget & "", rs!castfunc & "", rs!castcontext & "")
  End If
  
  Set Add = objCast
  
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
Public Property Get Item(Key As String) As pgCast
Attribute Item.VB_Description = "Returns a Cast object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Casts in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objCast As pgCast
Dim lCount As Long

  If IgnoreSys Then
    For Each objCast In mCol
      If Not objCast.SystemObject Then lCount = lCount + 1
    Next objCast
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Cast."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Cast.
  szSQL = "DROP CAST "
  szSQL = szSQL & " ( " & Me(Key).Source & " AS " & Me(Key).Target & " )"
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
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Casts."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
'pippo  objServer.Databases(Database).KillCastCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection

  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_CASTS, cnDatabase)
    While Not rs.EOF
      iAdd rs.Fields("oid"), rs!castsource, rs!casttarget, rs!castfunc, rs!castcontext & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Casts.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Casts.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objCast As pgCast
  
  szSQL = "-- Casts" & vbCrLf & vbCrLf

  For Each objCast In mCol
    If IncludeSystem Or (Not objCast.SystemObject) Then szSQL = szSQL & objCast.SQL & vbCrLf
  Next objCast
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
