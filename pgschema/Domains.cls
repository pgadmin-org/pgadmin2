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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

'Rollback to a prior version
Public Sub Rollback(Identifier As String, Version As Long)
Attribute Rollback.VB_Description = "Rollback the specified object to the specified version."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Rollback(" & QUOTE & Identifier & QUOTE & ", " & Version & ")", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Domains.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objDomain As pgDomain

  For Each objDomain In mCol
    If objDomain.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objDomain
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(dTimeStamp As Date, lOID As Long, szName As String, szOwner As String, szBaseType As String, lLength As Long, lScale As Long, szDefault As String, bNotNull As Boolean) As pgDomain
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.iAdd(" & QUOTE & dTimeStamp & QUOTE & ", " & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szBaseType & QUOTE & ", " & lLength & ", " & lScale & ", " & QUOTE & szDefault & QUOTE & ", " & bNotNull & ")", etFullDebug

Dim objNewMember As pgDomain

  'Create a new object
  Set objNewMember = New pgDomain
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
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
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal BaseType As String, Optional ByVal Length As Long, Optional ByVal NumericScale As Long, Optional ByVal Default As String, Optional ByVal NotNull As Boolean, Optional ByVal Comment As String)
Attribute Add.VB_Description = "Add a new Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & BaseType & QUOTE & ", " & Length & ", " & NumericScale & ", " & QUOTE & Default & QUOTE & ", " & NotNull & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objDomain As pgDomain
Dim szDataType As String
Dim szHex As String
Dim lLength As Long
Dim lScale As Long

  If objServer.Version.VersionNum < 7.3 Then
    Err.Raise 5432002, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Function
  End If
  
  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE DOMAIN " & QUOTE & dbSZ(Name) & QUOTE
  szSQL = szSQL & " AS " & BaseType
  If BaseType = "numeric" Then szSQL = szSQL & "(" & Length & ", " & NumericScale & ")"
  If BaseType = "char" Or BaseType = "varchar" Then szSQL = szSQL & "(" & Length & ")"
  If Default <> "" Then szSQL = szSQL & " DEFAULT " & Default
  If NotNull Then szSQL = szSQL & " NOT NULL"
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  Set rs = objServer.ExecSQL(SQL_GET_DOMAINS & " AND typname = '" & dbSZ(Name) & "'", cnDatabase)
    
  If Comment <> "" Then
    szSQL = "COMMENT ON DOMAIN " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
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
    Else
      If szDataType = "char" Or szDataType = "varchar" Then
        lLength = rs!typtypmod - 4
      End If
    End If
    Set objDomain = iAdd(rs!ts, rs!OID, rs!typname & "", rs!domainowner & "", szDataType, lLength, lScale, rs!typdefault & "", ToBool(rs!typnotnull))
  End If
  
  'Invalidate the Type Cache...
  objServer.Databases(Database).KillTypeCache
  
  If Not objDomain.SystemObject Then objDomain.Commit rcAdd, "Created object."
  Set Add = objDomain
  
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
Public Property Get Item(Key As String) As pgDomain
Attribute Item.VB_Description = "Returns a Domain object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  If objServer.Version.VersionNum < 7.3 Then
    Err.Raise 5432002, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Function
  End If
  
  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Domains in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objType As pgType
Dim lCount As Long

  If IgnoreSys Then
    For Each objType In mCol
      If Not objType.SystemObject Then lCount = lCount + 1
    Next objType
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Domain."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  If objServer.Version.VersionNum < 7.3 Then
    Err.Raise 5432002, App.Title, "Domains only exist in PostgreSQL 7.3 or above."
    Exit Sub
  End If

  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP DOMAIN " & QUOTE & dbSZ(mCol(ULEncode(Key)).Name) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rcDelete, "Deleted object."
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)

  'Invalidate the Type Cache...
  objServer.Databases(Database).KillTypeCache
  
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
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Domains."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Refresh()", etFullDebug

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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szDataType As String
Dim szHex As String
Dim lLength As Long
Dim lScale As Long

  'Create a new collection
  Set mCol = New Collection
  
  'Domains are only in 7.3+
  If objServer.dbVersion.VersionNum >= 7.3 Then
  
    'Get the data and loop through it
    If objServer.Connection.State = adStateOpen Then
      Set rs = objServer.ExecSQL(SQL_GET_DOMAINS & " ORDER BY typname", cnDatabase)
      While Not rs.EOF
      
        szDataType = objServer.Databases(Database).LookupType(rs!typbasetype)
        If szDataType = "bpchar" Then szDataType = "char"
        If szDataType = "numeric" Then
          szHex = Hex((Int(rs!typtypmod) - 4) And &HFFFF)
          lLength = CLng("&H" & Mid(szHex, 1, Len(szHex) - 4))
          lScale = CLng("&H" & Mid(szHex, Len(szHex) - 3, Len(szHex)))
        Else
          If szDataType = "char" Or szDataType = "varchar" Then
            lLength = rs!typtypmod - 4
          End If
        End If
  
        iAdd rs!ts, rs!OID, rs!typname & "", rs!domainowner & "", szDataType, lLength, lScale, rs!typdefault & "", ToBool(rs!typnotnull)
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
  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objDomain As pgDomain
  
  szSQL = "-- Domains" & vbCrLf & vbCrLf

  For Each objDomain In mCol
    If IncludeSystem Or (Not objDomain.SystemObject) Then szSQL = szSQL & objDomain.SQL & vbCrLf
  Next objDomain
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all changes to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Domains.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objDomain As pgDomain

  For Each objDomain In mCol
    If Not objDomain.SystemObject Then objDomain.Commit UpdateType, UpdateComment
  Next objDomain
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

