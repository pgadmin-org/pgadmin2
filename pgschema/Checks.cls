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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private lTableOID As Long
Private szTable As String
Private lNamespaceOID As Long
Private szNamespace As String

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objCheck As pgCheck

  For Each objCheck In mCol
    If objCheck.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objCheck
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property let TableOID(" & vData & ")", etFullDebug

  lTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(szName As String, szDefinition As String) As pgCheck
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.iAdd(" & QUOTE & szName & QUOTE & ", " & QUOTE & szDefinition & QUOTE & ")", etFullDebug

Dim objNewMember As pgCheck

  'Create a new object
  Set objNewMember = New pgCheck
  objNewMember.Name = szName
  objNewMember.SystemObject = False
  objNewMember.Table = szTable
  objNewMember.NamespaceOID = lNamespaceOID
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
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal Definition As String) As pgCheck
Attribute Add.VB_Description = "Add a new Check."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Definition & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  'Generate the szSQL query to create the new object and execute
  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szTable & QUOTE
  Else
    szSQL = "ALTER TABLE " & QUOTE & szTable & QUOTE
  End If
  szSQL = szSQL & " ADD CONSTRAINT " & QUOTE & Name & QUOTE
  szSQL = szSQL & " CHECK (" & Definition & ")"
  objServer.ExecSQL szSQL, cnDatabase
    
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_CHECKS & " AND rcrelid = " & lTableOID & " AND rcname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Not rs.EOF Then
    Set Add = iAdd(rs!rcname & "", rs!rcsrc & "")
  End If

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
Public Property Get Item(Key As String) As pgCheck
Attribute Item.VB_Description = "Returns a Check object."
Attribute Item.VB_UserMemId = 0
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Checks on the current Table."
Attribute Count.VB_MemberFlags = "200"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get Count()", etFullDebug

    Count = mCol.Count

  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Check object."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim szName As String

  If objServer.dbVersion.VersionNum < 7.2 Then Exit Sub
  
  szName = mCol(ULEncode(Key)).Name
  
  'Generate/Execute SQL to drop the Check.
  If objServer.dbVersion.VersionNum >= 7.3 Then
    szSQL = "ALTER TABLE " & QUOTE & szNamespace & QUOTE & "." & QUOTE & szTable & QUOTE & " DROP CONSTRAINT " & QUOTE & szName & QUOTE & " RESTRICT"
  Else
    szSQL = "ALTER TABLE " & QUOTE & szTable & QUOTE & " DROP CONSTRAINT " & QUOTE & szName & QUOTE & " RESTRICT"
  End If
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
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Checks."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_CHECKS & " AND rcrelid = " & lTableOID & " ORDER BY rcname", cnDatabase)
    While Not rs.EOF
      iAdd rs!rcname & "", rs!rcsrc & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Checks.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
