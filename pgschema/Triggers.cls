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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

Dim lTableOID As Long
Dim szTable As String

'Rollback to a prior version
Public Sub Rollback(Identifier As String, Version As Long)
Attribute Rollback.VB_Description = "Rollback the specified object to the specified version."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

Dim szDefinition As String
Dim objItem As rcItem

  'Get the definition to restore to first as it will vanish as soon as we
  'drop the current version
  For Each objItem In mCol(ULEncode(Identifier)).History
    If objItem.Version = Version Then
      szDefinition = objItem.Definition
      Exit For
    End If
  Next objItem
  If szDefinition = "" Then Exit Sub
  
  'Remove the current version
  Remove Identifier
  
  'Recreate the new object.
  objServer.Databases(Database).Execute szDefinition
  
  'Kill the Comment Cache and log the change
  objServer.Databases(Database).KillCommentCache
  Refresh
  mCol(ULEncode(Identifier)).Commit rcAdd, "Object rolled back to version " & Version & "."
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objTrigger As pgTrigger

  For Each objTrigger In mCol
    If objTrigger.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objTrigger
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dTimeStamp As Date, lOID As Long, szName As String, szTriggerFunction As String, szExecutes As String, szTriggerEvent As String, szForEach As String) As pgTrigger
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szTriggerFunction & QUOTE & ", " & QUOTE & szExecutes & QUOTE & ", " & QUOTE & szTriggerEvent & QUOTE & ", " & QUOTE & szForEach & QUOTE & ")", etFullDebug

Dim objNewMember As pgTrigger

  'Create a new object
  Set objNewMember = New pgTrigger
  If (lOID <= objServer.LastSystemOID) And Left(szName, 20) <> "RI_ConstraintTrigger" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Table = szTable
  objNewMember.TriggerFunction = szTriggerFunction
  objNewMember.Executes = szExecutes
  objNewMember.TriggerEvent = szTriggerEvent
  objNewMember.ForEach = szForEach
  objNewMember.Identifier = szName & " ON " & szTable
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal TriggerFunction As String, ByVal Executes As String, ByVal TriggerEvent As String, ByVal ForEach As String, Optional ByVal Comment As String) As pgTrigger
Attribute Add.VB_Description = "Add a new Trigger."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.iAdd(" & QUOTE & Name & QUOTE & ", " & QUOTE & TriggerFunction & QUOTE & ", " & QUOTE & Executes & QUOTE & ", " & QUOTE & TriggerEvent & QUOTE & ", " & QUOTE & ForEach & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim szExecutes As String
Dim szForEach As String
Dim szEvent As String
Dim objTrigger As pgTrigger
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE TRIGGER " & QUOTE & Name & QUOTE
  szSQL = szSQL & " " & Executes & " " & TriggerEvent
  szSQL = szSQL & " ON " & QUOTE & szTable & QUOTE
  szSQL = szSQL & " FOR EACH " & ForEach
  szSQL = szSQL & " EXECUTE PROCEDURE " & TriggerFunction
  objServer.ExecSQL szSQL, cnDatabase
  
  If Comment <> "" Then
    szSQL = "COMMENT ON TRIGGER " & QUOTE & Name & QUOTE & " ON " & QUOTE & szTable & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_TRIGGERS & " AND tgname = '" & dbSZ(Name) & "' AND tgrelid = " & lTableOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    'Executes
    If (rs!tgtype And 2) = 2 Then
      szExecutes = "Before"
    Else
      szExecutes = "After"
    End If
      
    'For Each
    If (rs!tgtype And 1) = 1 Then
      szForEach = "Row"
    Else
      szForEach = "Statement"
    End If
      
    If (rs!tgtype And 4) = 4 Then szEvent = szEvent & "Insert OR "
    If (rs!tgtype And 8) = 8 Then szEvent = szEvent & "Delete OR "
    If (rs!tgtype And 16) = 16 Then szEvent = szEvent & "Update OR "
    If Len(szEvent) > 0 Then szEvent = Trim(Left(szEvent, Len(szEvent) - 3))
      
    Set objTrigger = iAdd(rs!ts, rs!OID, rs!tgname & "", rs!proname & "(" & rs!tgargs & ")", szExecutes, szEvent, szForEach)
  End If
  
  If Not objTrigger.SystemObject Then objTrigger.Commit rcAdd, "Created object."
  Set Add = objTrigger
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgTrigger
Attribute Item.VB_Description = "Returns a Trigger object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Triggers on the current Table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
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
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Trigger."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Trigger.
  szSQL = "DROP TRIGGER " & QUOTE & mCol(ULEncode(Key)).Name & QUOTE & " ON " & QUOTE & mCol(ULEncode(Key)).Table & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rcDelete, "Deleted object."
  
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
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Triggers."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szExecutes As String
Dim szForEach As String
Dim szEvent As String

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_TRIGGERS & " AND tgrelid = " & lTableOID & " ORDER BY tgname", cnDatabase)
    While Not rs.EOF
    
      'Executes
      If (rs!tgtype And 2) = 2 Then
        szExecutes = "Before"
      Else
        szExecutes = "After"
      End If
        
      'For Each
      If (rs!tgtype And 1) = 1 Then
        szForEach = "Row"
      Else
        szForEach = "Statement"
      End If
      
      If (rs!tgtype And 4) = 4 Then szEvent = szEvent & "Insert OR "
      If (rs!tgtype And 8) = 8 Then szEvent = szEvent & "Delete OR "
      If (rs!tgtype And 16) = 16 Then szEvent = szEvent & "Update OR "
      If Len(szEvent) > 0 Then szEvent = Trim(Left(szEvent, Len(szEvent) - 3))
      
      iAdd rs!ts, rs!OID, rs!tgname & "", rs!proname & "(" & rs!tgargs & ")", szExecutes, szEvent, szForEach
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
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Triggers.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objTrigger As pgTrigger
  
  szSQL = "-- Triggers" & vbCrLf & vbCrLf

  For Each objTrigger In mCol
    If IncludeSystem Or (Not objTrigger.SystemObject) Then szSQL = szSQL & objTrigger.SQL & vbCrLf
  Next objTrigger
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

