VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
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

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(lOID As Long, szName As String, szOwner As String, lReturns As Long, szSource As String, lLanguage As Long, bCachable As Boolean, bStrict As Boolean, szArgumentList As String) As pgFunction
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.iAdd(" & lOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & lReturns & ", " & QUOTE & szSource & QUOTE & ", " & lLanguage & ", " & bCachable & ", " & bStrict & ", " & QUOTE & szArgumentList & QUOTE & ")", etFullDebug

Dim objNewMember As pgFunction
Dim rs As New Recordset
Dim szSQL As String
Dim szTemp() As String
Dim szArgs As String
Dim X As Long

  'Create a new object
  Set objNewMember = New pgFunction
  If lOID <= objServer.LastSystemOID Then
    If Not objServer.IncludeSys Then Exit Function
    objNewMember.SystemObject = True
  Else
    objNewMember.SystemObject = False
  End If
  objNewMember.OID = lOID
  objNewMember.Name = szName
  objNewMember.Owner = szOwner
  objNewMember.Returns = objServer.Databases(szDB).LookupType(lReturns)
  objNewMember.Source = szSource
  objNewMember.Language = objServer.Databases(szDB).LookupLanguage(lLanguage)
  objNewMember.Cachable = bCachable
  objNewMember.Strict = bStrict
  
  szTemp = Split(szArgumentList, " ")
  For X = 0 To UBound(szTemp)
    szArgs = szArgs & objServer.Databases(szDB).LookupType(Val(szTemp(X))) & ", "
  Next X
  If Len(szArgs) >= 2 Then szArgs = Left(szArgs, Len(szArgs) - 2)
  objNewMember.ArgumentList = szArgs
  Set objNewMember.Connection = cnDatabase
  
  'Add the Function
  objNewMember.Identifier = szName & "(" & szArgs & ")"
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
  
End Function

'External Add
Public Function Add(Name As String, Arguments As String, Returns As String, Source As String, Language As String, Cachable As Boolean, Strict As Boolean, Optional Comment As String) As pgFunction
Attribute Add.VB_Description = "Add a new function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Arguments & QUOTE & ", " & QUOTE & Returns & QUOTE & ", " & QUOTE & Source & QUOTE & ", " & QUOTE & Language & QUOTE & ", " & Cachable & ", " & Strict & ", " & QUOTE & Comment & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim fFunc As pgFunction
Dim bFlag As Boolean

  'Generate the szSQL query to create the new object and execute
  szSQL = "CREATE FUNCTION " & QUOTE & Name & QUOTE
  szSQL = szSQL & " (" & Arguments & ")"
  szSQL = szSQL & " RETURNS " & Returns
  szSQL = szSQL & " AS '" & Source & "'"
  szSQL = szSQL & " LANGUAGE '" & Language & "'"
  If Cachable Or Strict Then szSQL = szSQL & " WITH ("
  If Cachable Then szSQL = szSQL & "ISCACHABLE, "
  '7.1+ options
  If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
    If Strict Then szSQL = szSQL & "ISSTRICT, "
  End If
  If Right(szSQL, 2) = ", " Then szSQL = Left(szSQL, Len(szSQL) - 2) & ")"
  objServer.ExecSQL szSQL, cnDatabase

  'Generate and execute the szSQL to retrieve *all* properties of the new object
  szSQL = SQL_GET_FUNCTIONS & " WHERE proname = '" & dbSZ(Name) & "'"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)

  If Comment <> "" Then
    szSQL = "COMMENT ON FUNCTION " & QUOTE & Name & QUOTE & "(" & dbSZ(Arguments) & ") IS '" & dbSZ(Comment) & "'"
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  End If
  
  'There is no easy way to identify this function, so just
  'add any that don't already exist (by OID)
  While Not rs.EOF
    bFlag = False
    For Each fFunc In mCol
      If fFunc.OID = rs!OID Then
        bFlag = True
        Exit For
      End If
    Next fFunc
    If bFlag = False Then
      If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
        Set Add = iAdd(rs!OID, rs!proname & "", rs!proowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, rs!proiscachable, rs!proisstrict, rs!proargtypes & "")
      Else
        Set Add = iAdd(rs!OID, rs!proname & "", rs!proowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, rs!proiscachable, False, rs!proargtypes & "")
      End If
    End If
    rs.MoveNext
  Wend
  
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
Public Property Get Item(Key As String) As pgFunction
Attribute Item.VB_Description = "Returns a Type object."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Functions in the current database."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Count(" & IgnoreSys & ")", etFullDebug

Dim objFunction As pgFunction
Dim lCount As Long

  If IgnoreSys Then
    For Each objFunction In mCol
      If Not objFunction.SystemObject Then lCount = lCount + 1
    Next objFunction
    Count = lCount
  Else
    Count = mCol.Count
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Function."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim szArgs As String
Dim vArg As Variant

  szArgs = "("
  For Each vArg In mCol(ULEncode(Key)).Arguments
    szArgs = szArgs & QUOTE & vArg & QUOTE & ", "
  Next vArg
  If Len(szArgs) > 3 Then szArgs = Left(szArgs, Len(szArgs) - 2)
  szArgs = szArgs & ")"
  
  'Generate/Execute SQL to drop the Language.
  szSQL = "DROP FUNCTION " & QUOTE & dbSZ(mCol(ULEncode(Key)).Name) & QUOTE & szArgs
  objServer.ExecSQL szSQL, cnDatabase
  
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
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refrsh the collection of Functions."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Initialize()", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  'Get the Database Name
  szDB = GetDatabase(cnDatabase.ConnectionString)
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GET_FUNCTIONS & " ORDER BY proname", cnDatabase)
    While Not rs.EOF
      If objVersion.Major >= 7 And objVersion.Minor >= 1 Then
        iAdd rs!OID, rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), rs!proisstrict, rs!proargtypes & ""
      Else
        iAdd rs!OID, rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), False, rs!proargtypes & ""
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
objServer.iLogEvent "Entering " & App.Title & ":Functions.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objFunction As pgFunction
  
  szSQL = "-- Functions" & vbCrLf & vbCrLf

  For Each objFunction In mCol
    If IncludeSystem Or (Not objFunction.SystemObject) Then szSQL = szSQL & objFunction.SQL & vbCrLf
  Next objFunction
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

