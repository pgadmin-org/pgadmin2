VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ForeignKeys"
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
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Retruns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objForeignKey As pgForeignKey

  For Each objForeignKey In mCol
    If objForeignKey.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objForeignKey
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the TableOID for internal use.
Friend Property Let TableOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property let TableOID(" & vData & ")", etFullDebug

  dTableOID = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Table for internal use.
Friend Property Let Table(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property let Table(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szTable = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, lNArgs As Long, szArgs As String, bDeferrable As Boolean, bInitially As Boolean) As pgForeignKey
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & lNArgs & ", " & QUOTE & szArgs & QUOTE & ", " & bDeferrable & ", " & bInitially & ")", etFullDebug

Dim objNewMember As pgForeignKey
Dim szArgString As String
Dim szArgArray() As String
Dim szLocalColumns As String
Dim szReferencedColumns As String
Dim X As Long
Dim lKeyNum As Long

  'Create a new object
  Set objNewMember = New pgForeignKey
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.Oid = dOID
  objNewMember.Name = szName
  objNewMember.Table = szTable
  objNewMember.TableOID = dTableOID
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace

  If lNArgs >= 6 Then
    szArgArray = Split(szArgs, Chr(0))
    objNewMember.ReferencedTable = szArgArray(2)
    For X = 4 To UBound(szArgArray) Step 2
      If szArgArray(X) <> "" Then
        szLocalColumns = szLocalColumns & fmtID(szArgArray(X)) & "|"
        szReferencedColumns = szReferencedColumns & fmtID(szArgArray(X + 1)) & "|"
      End If
    Next X
  End If
  
  objNewMember.LocalColumns = szLocalColumns
  objNewMember.ReferencedColumns = szReferencedColumns
  objNewMember.Deferrable = bDeferrable
  If bInitially Then
    objNewMember.Initially = "Deferred"
  Else
    objNewMember.Initially = "Immediate"
  End If

  'Note, FKeys may have dup names, so we must make the identifiers unique.
  lKeyNum = 2
  If szName = "" Or szName = "<unnamed>" Then
    objNewMember.Identifier = "Unamed Foreign Key - " & dOID & " -> " & szArgArray(2)
    While Exists(objNewMember.Identifier)
      objNewMember.Identifier = "Unamed Foreign Key - " & dOID & " -> " & szArgArray(2) & " #" & lKeyNum
      lKeyNum = lKeyNum + 1
    Wend
  Else
    objNewMember.Identifier = szName & " -> " & szArgArray(2)
    While Exists(objNewMember.Identifier)
      objNewMember.Identifier = szName & " -> " & szArgArray(2) & " #" & lKeyNum
      lKeyNum = lKeyNum + 1
    Wend
  End If
  
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

Public Function Add(ByVal Name As String, ByVal Definition As String) As pgForeignKey
Attribute Add.VB_Description = "Add a new Foreign Keys."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Add(" & QUOTE & Name & QUOTE & "," & QUOTE & Definition & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim X As Long
Dim szArgs As String
  
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum < 7.3 Then Exit Function
  
  szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable)
  szSQL = szSQL & " ADD CONSTRAINT " & fmtID(Name)
  szSQL = szSQL & Definition
  objServer.ExecSQL szSQL, cnDatabase
    
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  Set rs = objServer.ExecSQL(SQL_GET_FOREIGN_KEYS & " AND tgrelid = " & dTableOID & "::oid AND tgconstrname = '" & dbSZ(Name) & "'", cnDatabase)
  If Not rs.EOF Then
    szArgs = ""
    For X = 0 To rs.Fields("tgargs").ActualSize - 1
      szArgs = szArgs & Chr(rs!tgargs(X))
    Next X
    Set Add = iAdd(rs.Fields("oid"), rs!tgconstrname & "", rs!tgnargs & "", szArgs, ToBool(rs!tgdeferrable), ToBool(rs!tginitdeferred))
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

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Foreign Keys object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.3 Then Exit Sub
  
  'Generate/Execute SQL to drop the ForeignKeys.
  szSQL = "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szTable) & " DROP CONSTRAINT " & mCol(ULEncode(Key)).FormattedID & " RESTRICT"
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

'Return the specified Item
Public Property Get Item(Key As String) As pgForeignKey
Attribute Item.VB_Description = "Returns a Foreign Key object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Foreign Keys on the current Table."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objForeignKey As pgForeignKey
Dim lCount As Long

  If IgnoreSys Then
    For Each objForeignKey In mCol
      If Not objForeignKey.SystemObject Then lCount = lCount + 1
    Next objForeignKey
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Foreign Keys."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Initialize()", etFullDebug

Dim rs As New Recordset
Dim X As Long
Dim szArgs As String

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_FOREIGN_KEYS & " AND tgrelid = " & dTableOID & "::oid ORDER BY tgconstrname", cnDatabase)
    While Not rs.EOF
      szArgs = ""
      For X = 0 To rs.Fields("tgargs").ActualSize - 1
        szArgs = szArgs & Chr(rs!tgargs(X))
      Next X
      iAdd rs.Fields("oid"), rs!tgconstrname & "", rs!tgnargs & "", szArgs, ToBool(rs!tgdeferrable), ToBool(rs!tginitdeferred)
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
objServer.iLogEvent "Entering " & App.Title & ":ForeignKeys.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub
