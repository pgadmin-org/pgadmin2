VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Checks"
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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objCheck As pgCheck

  For Each objCheck In mCol
    If objCheck.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objCheck
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(szName As String, szDefinition As String) As pgCheck
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.iAdd(" & QUOTE & szName & QUOTE & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgCheck

  'Create a new object
  Set objNewMember = New pgCheck
  objNewMember.Name = szName
  objNewMember.SystemObject = False
  objNewMember.Table = szTable
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.Definition = szDefinition
  If szName <> "" Then
    objNewMember.Identifier = szName
  Else
    objNewMember.Identifier = "Unamed Check - " & GetUniqueID
  End If
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal Definition As String) As pgCheck
Attribute Add.VB_Description = "Add a new Check."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Definition & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable)
  Else
    szSQL = "ALTER TABLE " & fmtID(szTable)
  End If
  szSQL = szSQL & " ADD CONSTRAINT " & fmtID(Name)
  szSQL = szSQL & " CHECK (" & Definition & ")"
  objServer.ExecSQL szSQL, cnDatabase
    
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If objVersion.VersionNum < 7.3 Then
    Set rs = objServer.ExecSQL(SQL_GET_CHECKS7_2 & " AND rcrelid = " & dTableOID & "::oid AND rcname = '" & dbSZ(Name) & "'", cnDatabase)
    If Not rs.EOF Then
      Set Add = iAdd(rs!rcname & "", rs!rcsrc & "")
    End If
  Else
    Set rs = objServer.ExecSQL(SQL_GET_CHECKS7_3 & " AND conrelid = " & dTableOID & "::oid AND conname = '" & dbSZ(Name) & "'", cnDatabase)
    If Not rs.EOF Then
      Set Add = iAdd(rs!conname & "", rs!consrc & "")
    End If
  End If

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
Public Property Get Item(Key As String) As pgCheck
Attribute Item.VB_Description = "Returns a Check object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Checks on the current Table."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Count()", etFullDebug

    Count = mCol.Count

  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Check object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.2 Then Exit Sub
  
  'Generate/Execute SQL to drop the Check.
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " DROP CONSTRAINT " & mCol(ULEncode(Key)).FormattedID & " RESTRICT"
  Else
    szSQL = "ALTER TABLE " & fmtID(szTable) & " DROP CONSTRAINT " & mCol(ULEncode(Key)).FormattedID & " RESTRICT"
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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Checks."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum < 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_CHECKS7_2 & " AND rcrelid = " & dTableOID & "::oid ORDER BY rcname", cnDatabase)
      While Not rs.EOF
        iAdd rs!rcname & "", rs!rcsrc & ""
        rs.MoveNext
      Wend
    Else
      Set rs = objServer.ExecSQL(SQL_GET_CHECKS7_3 & " AND conrelid = " & dTableOID & "::oid ORDER BY conname", cnDatabase)
      While Not rs.EOF
        iAdd rs!conname & "", rs!consrc & ""
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
  objServer.iLogError Err.Number, Err.Description
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub
