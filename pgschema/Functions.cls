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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

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
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objFunction As pgFunction

  For Each objFunction In mCol
    If objFunction.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objFunction
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Internal Add
Private Function iAdd(dOID As Double, szName As String, szOwner As String, dReturns As Double, szSource As String, dLanguage As Double, bCachable As Boolean, bStrict As Boolean, szArgumentList As String, Optional szACL As String, Optional szVolatility As String, Optional bRetSet As Boolean, Optional bSecDef As Boolean) As pgFunction
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.iAdd(" & dOID & ", " & QUOTE & szName & QUOTE & ", " & QUOTE & szOwner & QUOTE & ", " & dReturns & ", " & QUOTE & szSource & QUOTE & ", " & dLanguage & ", " & bCachable & ", " & bStrict & ", " & QUOTE & szArgumentList & QUOTE & ", " & QUOTE & szACL & QUOTE & ", " & QUOTE & szVolatility & QUOTE & ", " & bRetSet & ")", etFullDebug

Dim objNewMember As pgFunction
Dim rs As New Recordset
Dim szSQL As String
Dim szTemp() As String
Dim szArgs As String
Dim X As Long

  'Create a new object
  Set objNewMember = New pgFunction
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
  objNewMember.Returns = objServer.Databases(Database).LookupType(dReturns)
  
  'Cleanup the source code
  szSource = Replace(szSource, "\", "\\")
  szSource = Replace(szSource, "'", "''")
  objNewMember.iSource = szSource
  
  objNewMember.Language = objServer.Databases(Database).LookupLanguage(dLanguage)
  objNewMember.Cachable = bCachable
  objNewMember.Strict = bStrict
  objNewMember.ACL = szACL
  If objVersion.VersionNum >= 7.3 Then
    objNewMember.RetSet = bRetSet
    objNewMember.SecDef = bSecDef
    Select Case UCase(szVolatility)
      Case "I"
        objNewMember.Volatility = "Immutable"
      Case "S"
        objNewMember.Volatility = "Stable"
      Case "V"
        objNewMember.Volatility = "Volatile"
    End Select
  End If
  
  szTemp = Split(szArgumentList, " ")
  For X = 0 To UBound(szTemp)
    szArgs = szArgs & objServer.Databases(Database).LookupType(Val(szTemp(X))) & ", "
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
  objServer.iLogError Err.Number, Err.Description
  
End Function

'External Add
Public Function Add(Name As String, Arguments As String, Returns As String, Source As String, Language As String, Cachable As Boolean, Strict As Boolean, Optional Comment As String, Optional Volatility As String, Optional SecurityDefiner As Boolean, Optional RetSet As Boolean) As pgFunction
Attribute Add.VB_Description = "Add a new function."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Add(" & QUOTE & Name & QUOTE & ", " & QUOTE & Arguments & QUOTE & ", " & QUOTE & Returns & QUOTE & ", " & QUOTE & Source & QUOTE & ", " & QUOTE & Language & QUOTE & ", " & Cachable & ", " & Strict & ", " & QUOTE & Comment & QUOTE & ", " & QUOTE & Volatility & QUOTE & ", " & SecurityDefiner & "," & RetSet & ")", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
Dim szArgs As String
Dim fFunc As pgFunction
Dim bFlag As Boolean
Dim objFunction As pgFunction
 
  'Generate the szSQL query to create the new object and execute
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "CREATE FUNCTION " & fmtID(szNamespace) & "." & fmtID(Name)
  Else
    szSQL = "CREATE FUNCTION " & fmtID(Name)
  End If
  szSQL = szSQL & " (" & Arguments & ")"
  
  szSQL = szSQL & " RETURNS "
  If objVersion.VersionNum >= 7.3 Then
    If RetSet Then szSQL = szSQL & " SETOF "
  End If
  szSQL = szSQL & Returns
  
  szSQL = szSQL & " AS '" & Source & "'"
  szSQL = szSQL & " LANGUAGE '" & Language & "'"
  
  If objVersion.VersionNum >= 7.3 Then
    If Volatility <> "" Then szSQL = szSQL & " " & UCase(Volatility)
    If Strict Then szSQL = szSQL & " STRICT"
    If SecurityDefiner Then szSQL = szSQL & " SECURITY DEFINER"
  Else
    If Cachable Or Strict Or (Volatility <> "") Then szSQL = szSQL & " WITH ("
    If Cachable Then szSQL = szSQL & "ISCACHABLE, "
    '7.1+ options
    If objVersion.VersionNum >= 7.1 Then
      If Strict Then szSQL = szSQL & "ISSTRICT, "
    End If
  End If
  
  If Right(szSQL, 2) = ", " Then szSQL = Left(szSQL, Len(szSQL) - 2) & ")"
  objServer.ExecSQL szSQL, cnDatabase

  'Generate and execute the szSQL to retrieve *all* properties of the new object
  If objVersion.VersionNum >= 7.3 Then
    szSQL = SQL_GET_FUNCTIONS7_3 & " AND proname = '" & Name & "' AND pronamespace = " & dNamespaceOID & "::oid ORDER BY proname"
  Else
    szSQL = SQL_GET_FUNCTIONS7_1 & " WHERE proname = '" & Name & "' ORDER BY proname"
  End If
    
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)

  If Comment <> "" Then
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "COMMENT ON FUNCTION " & fmtID(szNamespace) & "." & fmtID(Name) & "(" & Arguments & ") IS '" & dbSZ(Comment) & "'"
    Else
      szSQL = "COMMENT ON FUNCTION " & fmtID(Name) & "(" & Arguments & ") IS '" & dbSZ(Comment) & "'"
    End If
    objServer.ExecSQL szSQL, cnDatabase
    objServer.Databases(Database).KillCommentCache
  End If
  
  'There is no easy way to identify this function, so just
  'add any that don't already exist (by OID)
  While Not rs.EOF
    bFlag = False
    For Each fFunc In mCol
      If fFunc.Oid = rs.Fields("oid") Then
        bFlag = True
        Exit For
      End If
    Next fFunc
    If bFlag = False Then
      'If pronargs = 1, but no arg types are specified then it should default to opaque
      If rs!pronargs = 1 And rs!proargtypes & "" = "" Then
        szArgs = "0"
      Else
        szArgs = rs!proargtypes & ""
      End If
      If objVersion.VersionNum >= 7.3 Then
        Set objFunction = iAdd(rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, False, rs!proisstrict, szArgs, rs!proacl & "", rs!provolatile & "", ToBool(rs!proretset), ToBool(rs!prosecdef))
      ElseIf objVersion.VersionNum >= 7.1 Then
        Set objFunction = iAdd(rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), rs!proisstrict, szArgs)
      Else
        Set objFunction = iAdd(rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), False, szArgs)
      End If
    End If
    rs.MoveNext
  Wend
  
  Set Add = objFunction
  
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
Public Property Get Item(Key As String) As pgFunction
Attribute Item.VB_Description = "Returns a Type object."
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Item count
Public Property Get Count(Optional IgnoreSys As Boolean) As Long
Attribute Count.VB_Description = "Returns a count of the Functions in the current database."
Attribute Count.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
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
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Remove the specified Item
Public Sub Remove(ByVal Key As String)
Attribute Remove.VB_Description = "Remove the specified Function."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Remove(" & QUOTE & Key & QUOTE & ")", etFullDebug

Dim szSQL As String
  
  'Generate/Execute SQL to drop the Function.
  szSQL = "DROP FUNCTION " & mCol(ULEncode(Key)).FormattedID
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
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection of Functions."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Refresh()", etFullDebug


  'Kill the relevant caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillTypeCache

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szArgs As String

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL(SQL_GET_FUNCTIONS7_3 & " AND pronamespace = " & dNamespaceOID & "::oid ORDER BY proname", cnDatabase)
    Else
      Set rs = objServer.ExecSQL(SQL_GET_FUNCTIONS7_1 & " ORDER BY proname", cnDatabase)
    End If
    While Not rs.EOF
      'If pronargs = 1, but no arg types are specified then it should default to opaque
      If rs!pronargs = 1 And rs!proargtypes & "" = "" Then
        szArgs = "0"
      Else
        szArgs = rs!proargtypes & ""
      End If
      If objVersion.VersionNum >= 7.3 Then
        iAdd rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, False, rs!proisstrict, szArgs, rs!proacl & "", rs!provolatile & "", ToBool(rs!proretset), ToBool(rs!prosecdef)
      ElseIf objVersion.VersionNum >= 7.1 Then
        iAdd rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), rs!proisstrict, szArgs
      Else
        iAdd rs.Fields("oid"), rs!proname & "", rs!funcowner & "", rs!prorettype, rs!prosrc & "", rs!prolang, ToBool(rs!proiscachable), False, szArgs
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
  objServer.iLogError Err.Number, Err.Description
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional IncludeSystem As Boolean) As String
Attribute SQL.VB_Description = "Returns the collection's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Functions.Property Get SQL(" & IncludeSystem & ")", etFullDebug
Dim szSQL As String
Dim objFunction As pgFunction
  
  szSQL = "-- Functions" & vbCrLf & vbCrLf

  For Each objFunction In mCol
    If IncludeSystem Or (Not objFunction.SystemObject) Then szSQL = szSQL & objFunction.SQL & vbCrLf
  Next objFunction
  
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
