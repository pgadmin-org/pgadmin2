VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Triggers"
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
Private cnDatabase As Connection

Private dTableOID As Double
Private szTable As String
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objTrigger As pgTrigger

  For Each objTrigger In mCol
    If objTrigger.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objTrigger
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szTriggerFunction As String, szExecutes As String, szTriggerEvent As String, szForEach As String) As pgTrigger
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szTriggerFunction & QUOTE & ", " & QUOTE & szExecutes & QUOTE & ", " & QUOTE & szTriggerEvent & QUOTE & ", " & QUOTE & szForEach & QUOTE & ")", etFullDebug

Dim objNewMember As pgTrigger

  'Create a new object
  Set objNewMember = New pgTrigger
  If dOID <= objServer.LastSystemOID Or Left(szName, 20) = "RI_ConstraintTrigger" Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.Oid = dOID
  objNewMember.iName = szName
  objNewMember.Table = szTable
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.TriggerFunction = szTriggerFunction
  objNewMember.Executes = szExecutes
  objNewMember.TriggerEvent = szTriggerEvent
  objNewMember.ForEach = szForEach
  If objVersion.VersionNum >= 7.3 Then
    objNewMember.Identifier = szName
  Else
    objNewMember.Identifier = szName
  End If
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal TriggerFunction As String, ByVal Executes As String, ByVal TriggerEvent As String, ByVal ForEach As String, Optional ByVal Comment As String) As pgTrigger
Attribute Add.VB_Description = "Add a new Trigger."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.iAdd(" & QUOTE & Name & QUOTE & ", " & QUOTE & TriggerFunction & QUOTE & ", " & QUOTE & Executes & QUOTE & ", " & QUOTE & TriggerEvent & QUOTE & ", " & QUOTE & ForEach & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim szExecutes As String
Dim szForEach As String
Dim szEvent As String
Dim szArgs As String
Dim X As Integer
Dim objTrigger As pgTrigger
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE TRIGGER " & fmtID(Name)
  szSQL = szSQL & " " & Executes & " " & TriggerEvent
  If objVersion.VersionNum >= 7.3 Then
    szSQL = szSQL & " ON " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    szSQL = szSQL & " ON " & fmtID(szTable)
  End If
  szSQL = szSQL & " FOR EACH " & ForEach
  szSQL = szSQL & " EXECUTE PROCEDURE " & TriggerFunction
  objServer.ExecSQL szSQL, cnDatabase
  
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON TRIGGER " & fmtID(Name) & " ON " & fmtID(szNamespace) & "." & fmtID(szTable) & " IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON TRIGGER " & fmtID(Name) & " ON " & fmtID(szTable) & " IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_TRIGGERS & " AND tgname = '" & dbSZ(Name) & "' AND tgrelid = " & dTableOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
  
    'Args
    If rs!tgargs.ActualSize > 0 Then szArgs = "'"
    For X = 0 To (rs!tgargs.ActualSize - 1)
      If rs!tgargs(X) = 0 Then
        If X = (rs!tgargs.ActualSize - 1) Then
          szArgs = szArgs & "'"
        Else
          szArgs = szArgs & "', '"
        End If
      Else
        szArgs = szArgs & Chr(rs!tgargs(X))
      End If
    Next X
      
    'Executes
    If (rs!tgtype And 2) = 2 Then
      szExecutes = "BEFORE"
    Else
      szExecutes = "AFTER"
    End If
      
    'For Each
    If (rs!tgtype And 1) = 1 Then
      szForEach = "ROW"
    Else
      szForEach = "STATEMENT"
    End If
      
    If (rs!tgtype And 4) = 4 Then szEvent = szEvent & "INSERT OR "
    If (rs!tgtype And 8) = 8 Then szEvent = szEvent & "DELETE OR "
    If (rs!tgtype And 16) = 16 Then szEvent = szEvent & "UPDATE OR "
    If Len(szEvent) > 0 Then szEvent = Trim(Left(szEvent, Len(szEvent) - 3))
      
    Set objTrigger = iAdd(rs.Fields("oid"), rs!tgname & "", rs!proname & "(" & szArgs & ")", szExecutes, szEvent, szForEach)
  End If
  
  Set Add = objTrigger
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgTrigger
Attribute Item.VB_Description = "Returns a Trigger object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Triggers on the current Table."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objTrigger As pgTrigger
Dim lCount As Long

  If IgnoreSys Then
    For Each objTrigger In mCol
      If Not objTrigger.SystemObject Then lCount = lCount + 1
    Next objTrigger
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Trigger."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Trigger.
  szSQL = "DROP TRIGGER " & mCol(ULEncode(Key)).FormattedID
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

'Rename
Public Sub Rename(ByVal OldName As String, ByVal NewName As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Rename(" & QUOTE & OldName & QUOTE & ", " & QUOTE & NewName & QUOTE & ")", etFullDebug

Dim objTrigger As New pgTrigger
    
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 9, App.Title, "Triggers can only be renamed in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  Set objTrigger = mCol(ULEncode(OldName & " ON " & szNamespace & "." & szTable))
  objTrigger.Name = NewName
  objTrigger.Identifier = NewName & " ON " & szNamespace & "." & szTable
  mCol.Add objTrigger, ULEncode(objTrigger.Identifier), ULEncode(OldName & " ON " & szNamespace & "." & szTable)
  mCol.Remove ULEncode(OldName & " ON " & szNamespace & "." & szTable)
  
  Exit Sub
Err_Handler:
  objServer.iLogError Err.Number, Err.Description
End Sub

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Triggers."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
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
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szExecutes As String
Dim szForEach As String
Dim szEvent As String
Dim szArgs As String
Dim X As Integer

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_TRIGGERS & " AND tgrelid = " & dTableOID & "::oid ORDER BY tgname", cnDatabase)
    While Not rs.EOF
    
      'Args
      szArgs = ""
      If rs!tgargs.ActualSize > 0 Then szArgs = "'"
      For X = 0 To (rs!tgargs.ActualSize - 1)
        If rs!tgargs(X) = 0 Then
          If X = (rs!tgargs.ActualSize - 1) Then
            szArgs = szArgs & "'"
          Else
            szArgs = szArgs & "', '"
          End If
        Else
          szArgs = szArgs & Chr(rs!tgargs(X))
        End If
      Next X
      
      'Executes
      If (rs!tgtype And 2) = 2 Then
        szExecutes = "BEFORE"
      Else
        szExecutes = "AFTER"
      End If
        
      'For Each
      If (rs!tgtype And 1) = 1 Then
        szForEach = "ROW"
      Else
        szForEach = "STATEMENT"
      End If
      
      szEvent = ""
      If (rs!tgtype And 4) = 4 Then szEvent = szEvent & "INSERT OR "
      If (rs!tgtype And 8) = 8 Then szEvent = szEvent & "DELETE OR "
      If (rs!tgtype And 16) = 16 Then szEvent = szEvent & "UPDATE OR "
      If Len(szEvent) > 0 Then szEvent = Trim(Left(szEvent, Len(szEvent) - 3))
      
      iAdd rs.Fields("oid"), rs!tgname & "", rs!proname & "(" & szArgs & ")", szExecutes, szEvent, szForEach
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
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objTrigger As pgTrigger
  
  szSQL = "-- Triggers" & vbCrLf & vbCrLf

  For Each objTrigger In mCol
    If IncludeSystem Or (Not objTrigger.SystemObject) Then szSQL = szSQL & objTrigger.SQL & vbCrLf
  Next objTrigger
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
