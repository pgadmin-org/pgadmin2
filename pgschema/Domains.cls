VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Domains"
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
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objDomain As pgDomain

  For Each objDomain In mCol
    If objDomain.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objDomain
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szOwner As String, szBaseType As String, lLength As Long, lScale As Long, szDefault As String, bNotNull As Boolean) As pgDomain
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szBaseType & QUOTE & ", " & lLength & ", " & lScale & ", " & QUOTE & szDefault & QUOTE & ", " & bNotNull & ")", etFullDebug

Dim objNewMember As pgDomain

  'Create a new object
  Set objNewMember = New pgDomain
  If dOID <= objServer.LastSystemOID Or Left(szName, 8) = "pgadmin_" Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.Oid = dOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.NamespaceOID = dNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.BaseType = szBaseType
  objNewMember.Length = lLength
  objNewMember.NumericScale = lScale
  objNewMember.Default = szDefault
  objNewMember.NotNull = bNotNull
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Domain
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal BaseType As String, Optional ByVal Length As Long, Optional ByVal NumericScale As Long, Optional ByVal Default As String, Optional ByVal NotNull As Boolean, Optional ByVal Comment As String)
Attribute Add.VB_Description = "Add a new Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & BaseType & QUOTE & ", " & Length & ", " & NumericScale & ", " & QUOTE & Default & QUOTE & ", " & NotNull & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objDomain As pgDomain
Dim szDataType As String
Dim szHex As String
Dim lLength As Long
Dim lScale As Long

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 2, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Function
  End If
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE DOMAIN " & fmtID(szNamespace) & "." & fmtID(Name)
  szSQL = szSQL & " AS " & BaseType
  If BaseType = "numeric" Then szSQL = szSQL & "(" & Length & ", " & NumericScale & ")"
  If BaseType = "char" Or BaseType = "varchar" Then szSQL = szSQL & "(" & Length & ")"
  If Default <> "" Then szSQL = szSQL & " DEFAULT " & Default
  If NotNull Then szSQL = szSQL & " NOT NULL"
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  Set rs = objServer.ExecSQL(SQL_GET_DOMAINS & " AND typname = '" & Name & "' AND typnamespace = " & dNamespaceOID & "::oid", cnDatabase)
    
  If Comment <> "" Then
    szSQL = "COMMENT ON DOMAIN " & fmtID(szNamespace) & "." & fmtID(Name) & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  If Not rs.EOF Then
    szDataType = objServer.Databases(Database).LookupType(rs!typbasetype)
    If szDataType = "bpchar" Then szDataType = "char"
    If szDataType = "numeric" Then
      szHex = Hex((Int(rs!typtypmod) - 4) And &HFFFF)
      lLength = CLng("&H" & Mid(szHex, 1, Len(szHex) - 4))
      lScale = CLng("&H" & Mid(szHex, Len(szHex) - 3, Len(szHex)))
    ElseIf szDataType = "char" Or szDataType = "varchar" Then
      lLength = rs!typtypmod - 4
    ElseIf rs!typlen = -1 And rs!typtypmod = -1 Then
      lLength = 0
    Else
      lLength = rs!typlen
    End If
    Set objDomain = iAdd(rs.Fields("oid"), rs!typname & "", rs!domainowner & "", szDataType, lLength, lScale, rs!typdefault & "", ToBool(rs!typnotnull))
  End If
  
  'Invalidate the Type Cache...
  objServer.Databases(Database).KillTypeCache

  Set Add = objDomain
  
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
Public Property Get Item(Key As String) As pgDomain
Attribute Item.VB_Description = "Returns a Domain object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 2, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Function
  End If
  
  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Domains in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objDomain As pgDomain
Dim lCount As Long

  If IgnoreSys Then
    For Each objDomain In mCol
      If Not objDomain.SystemObject Then lCount = lCount + 1
    Next objDomain
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Domain."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 2, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Sub
  End If

  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP DOMAIN " & mCol(ULEncode(Key)).FormattedID
  objServer.ExecSQL szSQL, cnDatabase

  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)

  'Invalidate the Type Cache...
  objServer.Databases(Database).KillTypeCache
  
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
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Domains."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Refresh()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Domains.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szDataType As String
Dim szHex As String
Dim lLength As Long
Dim lScale As Long

  'Create a new collection
  Set mCol = New Collection
  
  'Domains are only in 7.3+
  If objVersion.VersionNum >= 7.3 Then
  
    'Get the data and loop through it
    If objServer.Connection.State = adStateOpen Then
      Set rs = objServer.ExecSQL(SQL_GET_DOMAINS & " AND typnamespace = " & dNamespaceOID & "::oid ORDER BY typname", cnDatabase)
      While Not rs.EOF
      
        lLength = 0
        lScale = 0
        szDataType = objServer.Databases(Database).LookupType(rs!typbasetype)
        If szDataType = "bpchar" Then szDataType = "char"
        If szDataType = "numeric" Then
          szHex = Hex((Int(rs!typtypmod) - 4) And &HFFFF)
          lLength = CLng("&H" & Mid(szHex, 1, Len(szHex) - 4))
          lScale = CLng("&H" & Mid(szHex, Len(szHex) - 3, Len(szHex)))
        ElseIf szDataType = "char" Or szDataType = "varchar" Then
          lLength = rs!typtypmod - 4
        ElseIf rs!typlen = -1 And rs!typtypmod = -1 Then
          lLength = 0
        Else
          lLength = rs!typlen
        End If
  
        iAdd rs.Fields("oid"), rs!typname & "", rs!domainowner & "", szDataType, lLength, lScale, rs!typdefault & "", ToBool(rs!typnotnull)
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
objServer.iLogEvent "Entering " & App.Title & ":Types.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objDomain As pgDomain
  
  szSQL = "-- Domains" & vbCrLf & vbCrLf

  For Each objDomain In mCol
    If IncludeSystem Or (Not objDomain.SystemObject) Then szSQL = szSQL & objDomain.SQL & vbCrLf
  Next objDomain
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

