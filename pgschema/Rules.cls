VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rules"
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
Private dTableOID As Double
Private szTable As String
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objRule As pgRule

  For Each objRule In mCol
    If objRule.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objRule
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szDefinition As String) As pgRule
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgRule

  'Create a new object
  Set objNewMember = New pgRule
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = dOID
  objNewMember.Name = szName
  objNewMember.Table = szTable
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.Definition = szDefinition
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal TriggerEvent As String, ByVal Condition As String, ByVal Instead As Boolean, ByVal Action As String, Optional ByVal Comment As String) As pgRule
Attribute Add.VB_Description = "Add a new Rule."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & TriggerEvent & QUOTE & ", " & QUOTE & Condition & QUOTE & ", " & Instead & ", " & QUOTE & Action & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objRule As pgRule
  
  'Generate the szSQL query to create the new object and execute

  szSQL = "CREATE RULE " & fmtID(Name)
  szSQL = szSQL & " AS ON " & TriggerEvent
  If objVersion.VersionNum >= 7.3 Then
    szSQL = szSQL & " TO " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    szSQL = szSQL & " TO " & fmtID(szTable)
  End If
  If Condition <> "" Then szSQL = szSQL & " WHERE " & Condition
  szSQL = szSQL & " DO"
  If Instead = True Then szSQL = szSQL & " INSTEAD"
  szSQL = szSQL & " " & Action
  objServer.ExecSQL szSQL, cnDatabase
  
  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON RULE " & fmtID(Name) & " ON " & fmtID(szNamespace) & "." & fmtID(szTable) & " IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON RULE " & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If objVersion.VersionNum >= 7.3 Then
    szSQL = SQL_GET_RULES7_3 & " WHERE rulename = '" & dbSZ(Name) & "'"
  Else
    szSQL = SQL_GET_RULES7_1 & " WHERE rulename = '" & dbSZ(Name) & "'"
  End If
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set objRule = iAdd(rs.Fields("oid"), rs!rulename & "", rs!Definition & "")
  End If
  
  Set Add = objRule
  
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
Public Property Get Item(Key As String) As pgRule
Attribute Item.VB_Description = "Returns a Rule object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Rules on the current Table."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objRule As pgRule
Dim lCount As Long

  If IgnoreSys Then
    For Each objRule In mCol
      If Not objRule.SystemObject Then lCount = lCount + 1
    Next objRule
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Rule."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Rule.
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "DROP RULE " & mCol(ULEncode(Key)).FormattedID & " ON " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    szSQL = "DROP RULE " & mCol(ULEncode(Key)).FormattedID
  End If
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
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Rules."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Rules.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_RULES7_3 & " WHERE ev_class = " & dTableOID & "::oid ORDER BY rulename", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_RULES7_1 & " WHERE ev_class = " & dTableOID & "::oid ORDER BY rulename", cnDatabase)
    End If
    While Not rs.EOF
      iAdd rs.Fields("oid"), rs!rulename & "", rs!Definition & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Rules.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Rules.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objRule As pgRule
  
  szSQL = "-- Rules" & vbCrLf & vbCrLf

  For Each objRule In mCol
    If IncludeSystem Or (Not objRule.SystemObject) Then szSQL = szSQL & objRule.SQL & vbCrLf
  Next objRule
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
