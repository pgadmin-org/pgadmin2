VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Collections
Private colArguments As Arguments

'Properties
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szReturns As String
Private szSource As String
Private szLanguage As String
Private bCachable As Boolean
Private bStrict As Boolean
Private szArgumentList As String
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant
Private szACL As String
Private szVolatility As String
Private bRetSet As Boolean
Private bSecDef As Boolean
Private colDependent As Collection
Private colReferenced As Collection

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get FormattedID()", etFullDebug

Dim szArgs As String
Dim vArg As Variant
Dim X As Long

  'Get the arguments if we don't have them
  If colArguments Is Nothing Then
    Set colArguments = New Arguments
    
    'Pass the connection to the class
    Set colArguments.Connection = cnDatabase
    colArguments.Initialize szArgumentList
  End If
  
  szArgs = "("
  For Each vArg In colArguments
    szArgs = szArgs & vArg & ", "
  Next vArg
  If Len(szArgs) > 3 Then szArgs = Left(szArgs, Len(szArgs) - 2)
  szArgs = szArgs & ")"

  If objVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szNamespace) & "." & fmtID(szName) & szArgs
  Else
    FormattedID = fmtID(szName) & szArgs
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Tag()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let NameSpace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NameSpace() As String
Attribute NameSpace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Namespace()", etFullDebug

  NameSpace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Arguments Collection
Public Property Get Arguments() As Arguments
Attribute Arguments.VB_Description = "Returns a collection of Function Arguments."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Arguments()", etFullDebug

  If colArguments Is Nothing Then
    Set colArguments = New Arguments
    
    'Pass the connection to the class
    Set colArguments.Connection = cnDatabase
    colArguments.Initialize szArgumentList
  End If
  Set Arguments = colArguments
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let Oid(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Oid() As Double
Attribute Oid.VB_Description = "Returns the Type OID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get OID()", etFullDebug

  Oid = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Type name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get ObjectType()", etFullDebug

  ObjectType = "Function"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns an object representing the Type's owner."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Return Type
Friend Property Let Returns(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Returns(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReturns = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Returns() As String
Attribute Returns.VB_Description = "Returns the Function Return data type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Returns()", etFullDebug

  Returns = szReturns
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Function Source
Public Property Let Source(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Source(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim rs As New Recordset

  If objVersion.VersionNum < 7.2 Then
    Err.Raise vbObjectError + 5, App.Title, "Functions can only be updated in PostgreSQL 7.2 or above."
    Exit Property
  End If
  
  'Set the source
  szSource = vData
  objServer.Databases(Database).Execute SQL(True)
    
  'Select the prosrc back from the DB...
  Set rs = objServer.ExecSQL("SELECT prosrc FROM pg_proc WHERE oid = " & dOID & "::oid", cnDatabase)
  If Not rs.EOF Then szSource = rs!prosrc & ""
  
  'Clean code
  szSource = Replace(szSource, "\", "\\")
  szSource = Replace(szSource, "'", "''")
  
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Property
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Let iSource(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let iSource(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szSource = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Source() As String
Attribute Source.VB_Description = "Sets/Returns the Function Source. This may be the path to an object library, or Procedural Language code. This property is Read-Only with PostgreSQL < v7.2"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Source()", etFullDebug

  Source = szSource
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Language
Friend Property Let Language(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Language(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLanguage = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Language() As String
Attribute Language.VB_Description = "RFeturns the Function Language."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Language()", etFullDebug

  Language = szLanguage
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return Cachable?
Friend Property Let Cachable(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Cachable(" & vData & ")", etFullDebug

  bCachable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Cachable() As Boolean
Attribute Cachable.VB_Description = "Returns a value indicating whether the Function result may be cached."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Cachable()", etFullDebug

  Cachable = bCachable
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return Strict?
Friend Property Let Strict(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Strict(" & vData & ")", etFullDebug

  bStrict = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Strict() As Boolean
Attribute Strict.VB_Description = "Returns a value indicating whether the Function is Strict."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Strict()", etFullDebug

  Strict = bStrict
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return RetSet?
Friend Property Let RetSet(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let RetSet(" & vData & ")", etFullDebug

  bRetSet = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get RetSet() As Boolean
Attribute RetSet.VB_Description = "Returns a value indicating whether the function returns a set."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get RetSet()", etFullDebug

  RetSet = bRetSet
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return SecDef?
Friend Property Let SecDef(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let SecDef(" & vData & ")", etFullDebug

  bSecDef = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SecDef() As Boolean
Attribute SecDef.VB_Description = "Returns a value indicating whether the function is a security definer."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get SecDef()", etFullDebug

  SecDef = bSecDef
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Volatility
Friend Property Let Volatility(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Volatility(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szVolatility = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Volatility() As String
Attribute Volatility.VB_Description = "Returns a value indicating the volatility of the result."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Volatility()", etFullDebug

  Volatility = szVolatility
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON FUNCTION " & FormattedID & " IS '" & dbSZ(vData) & "'"

  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Type comment."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(dOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the ArgumentList
Friend Property Let ArgumentList(ByVal vData As String)
Attribute ArgumentList.VB_Description = "Returns the Function Arguments."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let ArgumentList(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szArgumentList = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the objects Access Control List."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant permissions on the object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 13, App.Title, "Function security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclExecute) = aclExecute Then szSQL = szSQL & "EXECUTE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON FUNCTION " & FormattedID & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT proacl FROM pg_proc WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!proacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke permissions from the object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If objVersion.VersionNum < 7.2 Then
    Err.Raise vbObjectError + 13, App.Title, "Function security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclExecute) = aclExecute Then szSQL = szSQL & "EXECUTE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON FUNCTION " & FormattedID & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT proacl FROM pg_proc WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!proacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL(Optional OrReplace As Boolean = False) As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get SQL(" & OrReplace & ")", etFullDebug
Dim szSQL As String
Dim szArgs As String
Dim vArg As Variant
Dim X As Long

  szSQL = "-- Function: " & FormattedID & vbCrLf
  If (objVersion.VersionNum >= 7.2) And OrReplace Then
    szSQL = szSQL & "CREATE OR REPLACE FUNCTION " & FormattedID
  Else
    szSQL = szSQL & "CREATE FUNCTION " & FormattedID
  End If
  
  szSQL = szSQL & vbCrLf
  szSQL = szSQL & "  RETURNS "
  If objVersion.VersionNum >= 7.3 Then
    If bRetSet Then szSQL = szSQL & " SETOF "
  End If
  szSQL = szSQL & szReturns
  
  szSQL = szSQL & " AS" & vbCrLf
  szSQL = szSQL & "'"
  szSQL = szSQL & szSource & vbCrLf
  szSQL = szSQL & "'" & vbCrLf
  szSQL = szSQL & "  LANGUAGE '" & dbSZ(szLanguage) & "'"
  
  If objVersion.VersionNum >= 7.3 Then
    If szVolatility <> "" Then szSQL = szSQL & " " & UCase(szVolatility)
    If bStrict Then szSQL = szSQL & " STRICT"
    If bSecDef Then szSQL = szSQL & " SECURITY DEFINER"
  Else
    If bCachable Or bStrict Or (szVolatility <> "") Then szSQL = szSQL & " WITH ("
    If bCachable Then szSQL = szSQL & "ISCACHABLE, "
    If bStrict Then szSQL = szSQL & "ISSTRICT, "
    If bCachable Or bStrict Or (szVolatility <> "") Then szSQL = Mid(szSQL, 1, Len(szSQL) - 2) & ")"
  End If
  
  szSQL = szSQL & ";" & vbCrLf
  
  If szACL <> "" Then szSQL = szSQL & ParseACL(FormattedID, szACL, aclFunction)
  
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON FUNCTION " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the dependent object is in
Public Function Dependent() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Function Dependent()", etFullDebug
  
  If colDependent Is Nothing Then Set colDependent = DepRef(dOID, cnDatabase, Me.Database, EDR_Depend)
  Set Dependent = colDependent
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Return the referenced object is in
Public Function Referenced() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Function Dependent()", etFullDebug
  
  If colReferenced Is Nothing Then Set colReferenced = DepRef(dOID, cnDatabase, Me.Database, EDR_Reference)
  Set Referenced = colReferenced
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Class_Terminate()", etFullDebug
  
  Set colDependent = Nothing
  Set colReferenced = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

