VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private szDB As String

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(lOID As Long, szName As String, szOwner As String, szInputFunction As String, szOutputFunction As String, lInternalLength As Long, lExternalLength As Long, szDefault As String, lElement As Long, szDelimiter As String, szSendFunction As String, szReceiveFunction As String, bPassedByValue As Boolean, szAlignment As String, szStorage As String) As pgType
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & QUOTE & szInputFunction & QUOTE & ", " & QUOTE & szOutputFunction & QUOTE & ", " & lInternalLength & ", " & lExternalLength & ", " & QUOTE & szDefault & QUOTE & ", " & lElement & ", " & QUOTE & szDelimiter & QUOTE & ", " & QUOTE & szSendFunction & QUOTE & ", " & QUOTE & szReceiveFunction & QUOTE & ", " & bPassedByValue & ", " & QUOTE & szAlignment & QUOTE & ", " & QUOTE & szStorage & QUOTE & ")", etFullDebug

Dim objNewMember As pgType
Dim szTemp As String

  'Create a new object
  Set objNewMember = New pgType
  If (lOID <= objServer.LastSystemOID) Or (Left(szName, 1) = "_") Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.InputFunction = szInputFunction
  objNewMember.OutputFunction = szOutputFunction
  objNewMember.InternalLength = lInternalLength
  objNewMember.ExternalLength = lExternalLength
  objNewMember.Default = szDefault
  szTemp = objServer.Databases(szDB).LookupType(lElement)
  If szTemp = "opaque" Then szTemp = ""
  objNewMember.Element = szTemp
  objNewMember.Delimiter = szDelimiter
  objNewMember.SendFunction = szSendFunction
  objNewMember.ReceiveFunction = szReceiveFunction
  objNewMember.PassedByValue = bPassedByValue
  Select Case UCase(szAlignment)
    Case ""
      objNewMember.Alignment = ""
    Case "C"
      objNewMember.Alignment = "char"
    Case "S"
      objNewMember.Alignment = "short"
    Case "I"
      objNewMember.Alignment = "int"
    Case "D"
      objNewMember.Alignment = "double"
  End Select
  Select Case UCase(szStorage)
    Case ""
      objNewMember.Storage = "Plain"
    Case "P"
      objNewMember.Storage = "Plain"
    Case "E"
      objNewMember.Storage = "External"
    Case "M"
      objNewMember.Storage = "Extended"
    Case "X"
      objNewMember.Storage = "Main"
  End Select
  objNewMember.Identifier = szName
  Set objNewMember.Connection = cnDatabase
  
  'Add the Type
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'External Add
Public Function Add(ByVal Name As String, ByVal InputFunction As String, ByVal OutputFunction As String, ByVal InternalLength As Long, Optional ByVal ExternalLength As Long, Optional ByVal Default As String, Optional ByVal Element As String, Optional ByVal Delimiter As String, Optional ByVal SendFunction As String, Optional ByVal ReceiveFunction As String, Optional ByVal PassedByValue As Boolean, Optional ByVal Alignment As String, Optional ByVal Storage As String, Optional ByVal Comment As String)
Attribute Add.VB_Description = "Add a new Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.iAdd(" & QUOTE & Name & QUOTE & ", " & QUOTE & InputFunction & QUOTE & ", " & QUOTE & OutputFunction & QUOTE & ", " & InternalLength & ", " & ExternalLength & ", " & QUOTE & Default & QUOTE & ", " & QUOTE & Element & QUOTE & ", " & QUOTE & Delimiter & QUOTE & ", " & QUOTE & SendFunction & QUOTE & ", " & QUOTE & ReceiveFunction & QUOTE & ", " & PassedByValue & ", " & QUOTE & Alignment & QUOTE & ", " & QUOTE & Storage & QUOTE & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim objType As pgType

  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE TYPE " & QUOTE & dbSZ(Name) & QUOTE
  szSQL = szSQL & " (INPUT = " & QUOTE & InputFunction & QUOTE
  szSQL = szSQL & ", OUTPUT = " & QUOTE & OutputFunction & QUOTE
  szSQL = szSQL & ", INTERNALLENGTH = " & InternalLength
  If ExternalLength > 0 Then szSQL = szSQL & ", EXTERNALLENGTH = " & ExternalLength
  If Default <> "" Then szSQL = szSQL & ", DEFAULT = " & QUOTE & Default & QUOTE
  If Element <> "" Then szSQL = szSQL & ", ELEMENT = " & QUOTE & Element & QUOTE
  If Delimiter <> "" Then szSQL = szSQL & ", DELIMITER = " & QUOTE & Delimiter & QUOTE
  If SendFunction <> "" Then szSQL = szSQL & ", SEND = " & QUOTE & SendFunction & QUOTE
  If ReceiveFunction <> "" Then szSQL = szSQL & ", RECEIVE = " & QUOTE & ReceiveFunction & QUOTE
  '7.1+ support ALIGN & STORAGE
  If PassedByValue = True Then szSQL = szSQL & ", PASSEDBYVALUE"
  If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
    If Alignment <> "" Then szSQL = szSQL & ", ALIGNMENT = " & QUOTE & Alignment & QUOTE
    If Storage <> "" Then szSQL = szSQL & ", STORAGE = " & QUOTE & Storage & QUOTE
  End If
  szSQL = szSQL & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_TYPES & " AND typname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    
  If Comment <> "" Then
    szSQL = "COMMENT ON TYPE " & QUOTE & Name & QUOTE & " IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(szDB).KillCommentCache
  End If
  
  If Not rs.EOF Then
    If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
      Set objType = iAdd(rs!OID, rs!typname & "", rs!typeowner & "", rs!typinput & "", rs!typoutput & "", rs!typlen, rs!typprtlen, rs!typdefault & "", rs!typelem, rs!typdelim & "", rs!typsend & "", rs!typreceive & "", ToBool(rs!typbyval), rs!typalign & "", rs!typstorage & "")
    Else
      Set objType = iAdd(rs!OID, rs!typname & "", rs!typeowner & "", rs!typinput & "", rs!typoutput & "", rs!typlen, rs!typprtlen, rs!typdefault & "", rs!typelem, rs!typdelim & "", rs!typsend & "", rs!typreceive & "", ToBool(rs!typbyval), "", "")
    End If
  End If
  
  'Invalidate the Type Cache...
  objServer.Databases(szDB).KillTypeCache
  
  If Not objType.SystemObject Then objType.Commit rlAdd, "Created object."
  Set Add = objType
  
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
Public Property Get Item(Key As String) As pgType
Attribute Item.VB_Description = "Returns a Type object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Types in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Get Count(" & IgnoreSys & ")", etFullDebug

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
Attribute Remove.VB_Description = "Remove the specified Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP TYPE " & QUOTE & dbSZ(mCol(ULEncode(Key)).Name) & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  If Not mCol(ULEncode(Key)).SystemObject Then mCol(ULEncode(Key)).Commit rlDelete, "Deleted object."
  
  'If we get this far, then we can't have errored,
  'so remove the object from the collection.
  mCol.Remove ULEncode(Key)

  'Invalidate the Type Cache...
  objServer.Databases(szDB).KillTypeCache
  
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
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Types."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the Database Name
  szDB = GetDatabase(cnDatabase.ConnectionString)
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_TYPES & " ORDER BY typname", cnDatabase)
    While Not rs.EOF
      If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
        iAdd rs!OID, rs!typname & "", rs!typeowner & "", rs!typinput & "", rs!typoutput & "", rs!typlen, rs!typprtlen, rs!typdefault & "", rs!typelem, rs!typdelim & "", rs!typsend & "", rs!typreceive & "", ToBool(rs!typbyval), rs!typalign & "", rs!typstorage & ""
      Else
        iAdd rs!OID, rs!typname & "", rs!typeowner & "", rs!typinput & "", rs!typoutput & "", rs!typlen, rs!typprtlen, rs!typdefault & "", rs!typelem, rs!typdelim & "", rs!typsend & "", rs!typreceive & "", ToBool(rs!typbyval), "", ""
      End If
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
objServer.iLogEvent "Entering " & App.Title & ":Types.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Types.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objType As pgType
  
  szSQL = "-- Types" & vbCrLf & vbCrLf

  For Each objType In mCol
    If IncludeSystem Or (Not objType.SystemObject) Then szSQL = szSQL & objType.SQL & vbCrLf
  Next objType
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property



