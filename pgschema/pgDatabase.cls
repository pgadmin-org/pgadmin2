VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colLanguages As Languages
Private colNamespaces As Namespaces
Private colDatabaseVars As DatabaseVars

'Properties
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szPath As String
Private szOwner As String
Private szServerEncoding As String
Private lStatus As dbStatus
Private bAllowConnections As Boolean
Private cnDatabase As New Connection
Private szDatabaseVarList As String
Private szACL As String

Private szTypeCache As Collection
Private szLanguageCache As Collection
Private szOperatorCache As Collection
Private szCommentCache As Collection
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Friend Sub KillTypeCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillTypeCache()", etFullDebug

  objServer.iLogEvent "Invalidating Type cache (database: " & szName & ")...", etMiniDebug
  Set szTypeCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

Friend Function LookupType(dOID As Double) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupType(" & dOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Populate the cache
  If szTypeCache Is Nothing Then
    objServer.iLogEvent "Populating Type cache (database: " & szName & ")...", etMiniDebug
    Set szTypeCache = New Collection
    If objVersion.VersionNum >= 7.3 Then
      Set rs = objServer.ExecSQL("SELECT t.oid, n.nspname, CASE WHEN (t.typlen = -1 AND t.typelem != 0) THEN (SELECT at.typname FROM pg_type at WHERE at.oid = t.typelem) || '[]' ELSE t.typname END AS typname FROM pg_type t, pg_namespace n WHERE n.oid = t.typnamespace", cnDatabase)
      While Not rs.EOF
        If rs!nspname & "" = "pg_catalog" Then
          szTypeCache.Add fmtTypeID(rs!typname & ""), "T" & rs!OID
        Else
          szTypeCache.Add fmtTypeID(rs!nspname & "") & "." & fmtID(rs!typname & ""), "T" & rs!OID
        End If
        rs.MoveNext
      Wend
    Else
      Set rs = objServer.ExecSQL("SELECT oid, typname FROM pg_type", cnDatabase)
      While Not rs.EOF
        szTypeCache.Add fmtTypeID(rs!typname & ""), "T" & rs!OID
        rs.MoveNext
      Wend
    End If
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Completed populating Type cache (database: " & szName & ").", etMiniDebug
  End If
  
  If dOID = 0 Then
    LookupType = "opaque"
  Else
    On Error Resume Next
    LookupType = szTypeCache("T" & dOID)
    On Error GoTo Err_Handler
  End If
 
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

Friend Sub KillLanguageCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillLanguageCache()", etFullDebug

  objServer.iLogEvent "Invalidating Language cache (database: " & szName & ")...", etMiniDebug
  Set szLanguageCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

Friend Function LookupLanguage(dOID As Double) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupLanguage(" & dOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Populate the cache
  If szLanguageCache Is Nothing Then
    objServer.iLogEvent "Populating Language cache (database: " & szName & ")...", etMiniDebug
    Set szLanguageCache = New Collection
    Set rs = objServer.ExecSQL(SQL_GET_LANGUAGES, cnDatabase)
    While Not rs.EOF
      szLanguageCache.Add rs!lanname & "", "L" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Complete populating Language cache (database: " & szName & ").", etMiniDebug
  End If
  
  On Error Resume Next
  LookupLanguage = szLanguageCache("L" & dOID)
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

Friend Sub KillOperatorCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillOperatorCache()", etFullDebug

  objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
  Set szOperatorCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

Friend Function LookupOperator(dOID As Double) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupOperator(" & dOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Build the Operator Cache
  If szOperatorCache Is Nothing Then
    objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
    Set szOperatorCache = New Collection
    If objVersion.VersionNum >= 7.3 Then
      'Select an O in front of the name to avoid strange behaviour with !
      Set rs = objServer.ExecSQL("SELECT o.oid, n.nspname, o.oprname FROM pg_operator o, pg_namespace n WHERE n.oid = o.oprnamespace", cnDatabase)
      While Not rs.EOF
        'Prefix the operator name otherwise we get bizarre errors
        If rs!nspname & "" = "pg_catalog" Then
          szOperatorCache.Add "O" & rs!oprname, "O" & rs!OID
        Else
          szOperatorCache.Add "O" & fmtID(rs!nspname) & "." & rs!oprname, "O" & rs!OID
        End If
        rs.MoveNext
      Wend
    Else
      'Select an O in front of the name to avoid strange behaviour with !
      Set rs = objServer.ExecSQL("SELECT oid, oprname FROM pg_operator", cnDatabase)
      While Not rs.EOF
        'Prefix the operator name otherwise we get bizarre errors
        szOperatorCache.Add "O" & rs!oprname, "O" & rs!OID
        rs.MoveNext
      Wend
    
    End If
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Completed populating Operator cache (database: " & szName & ").", etMiniDebug
  End If
  
  On Error Resume Next
  LookupOperator = Mid(szOperatorCache("O" & dOID), 2)
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

Friend Sub KillCommentCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillCommentCache()", etFullDebug

  objServer.iLogEvent "Invalidating Comment cache (database: " & szName & ")...", etMiniDebug
  Set szCommentCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err.Number, Err.Description
End Sub

Friend Function LookupComment(dOID As Double, Optional lSubID As Double) As String
'If anything goes wrong here, just continue.
On Error Resume Next
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupComment(" & dOID & ", " & lSubID & ")", etFullDebug
Dim rs As New Recordset
Dim arrTmp As Variant
Dim X As Long

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Build the Comment Cache
  If szCommentCache Is Nothing Then
    objServer.iLogEvent "Populating Comment cache (database: " & szName & ")...", etMiniDebug
    Set szCommentCache = New Collection
    
    '7.2+ Uses OID + SubID for some comments.
    If objVersion.VersionNum < 7.2 Then
      Set rs = objServer.ExecSQL("SELECT objoid, description FROM pg_description", cnDatabase)
      arrTmp = rs.GetRows
      For X = 0 To UBound(arrTmp, 2)
        szCommentCache.Add arrTmp(1, X) & "", "C" & arrTmp(0, X) & "S0"
      Next X
    Else
      Set rs = objServer.ExecSQL("SELECT objoid, objsubid, description FROM pg_description", cnDatabase)
      arrTmp = rs.GetRows
      For X = 0 To UBound(arrTmp, 2)
        szCommentCache.Add arrTmp(2, X) & "", "C" & arrTmp(0, X) & "S" & arrTmp(1, X)
      Next X
    End If
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Completed populating Comment cache (database: " & szName & ").", etMiniDebug
  End If
  
  LookupComment = szCommentCache("C" & dOID & "S" & lSubID)

End Function

'Return the Languages Collection
Public Property Get Languages() As Languages
Attribute Languages.VB_Description = "Returns a collection of pgLanguage objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Languages()", etFullDebug

  If colLanguages Is Nothing Then
    If dbConnect = False Then Exit Property
    Set colLanguages = New Languages
    
    'Pass the connection to the class
    Set colLanguages.Connection = cnDatabase
    colLanguages.Initialize
  End If
  Set Languages = colLanguages
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Namespaces Collection
Public Property Get Namespaces() As Namespaces
Attribute Namespaces.VB_Description = "Returns a collection of pgNamespace objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Namespaces()", etFullDebug

  If colNamespaces Is Nothing Then
    If dbConnect = False Then Exit Property
    Set colNamespaces = New Namespaces
        
    'Pass the connection to the class
    Set colNamespaces.Connection = cnDatabase
    colNamespaces.Initialize
  End If
  Set Namespaces = colNamespaces
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the DatabaseVars Collection
Public Property Get DatabaseVars() As DatabaseVars
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgUser.Property Get DatabaseVars()", etFullDebug

  If colDatabaseVars Is Nothing Then
    If dbConnect = False Then Exit Property
    Set colDatabaseVars = New DatabaseVars
    
    'Pass the connection to the class
    Set colDatabaseVars.Connection = cnDatabase
    colDatabaseVars.Initialize szDatabaseVarList
  End If
  Set DatabaseVars = colDatabaseVars
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the ADO Connection String
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_Description = "Returns the database connection string."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ConnectionString()", etFullDebug

  If cnDatabase.State <> adStateOpen Then
    ConnectionString = ""
  Else
    ConnectionString = cnDatabase.ConnectionString
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the OID of the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Database Name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ObjectType()", etFullDebug

  ObjectType = "Database"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Path
Friend Property Let Path(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Path(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szPath = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns the Database path."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Path()", etFullDebug

  Path = szPath
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Database owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Database Server Encoding
Friend Property Let ServerEncoding(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let ServerEncoding(" & vData & ")", etFullDebug

  szServerEncoding = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ServerEncoding() As String
Attribute ServerEncoding.VB_Description = "Returns the name of the encoding scheme."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ServerEncoding()", etFullDebug

  ServerEncoding = szServerEncoding
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
Attribute Comment.VB_Description = "Sets/Returns the Database comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Comment(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Connect if required
  If dbConnect = False Then Exit Property
  
  szSQL = "COMMENT ON DATABASE " & FormattedID & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Comment()", etFullDebug
  
  If objServer.DeferConnection And lStatus <> statOpen Then
    Comment = ""
  Else
    Comment = LookupComment(dOID)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Status
Public Property Get Status() As dbStatus
Attribute Status.VB_Description = "Returns the database connection status."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Status()", etFullDebug

  Status = lStatus
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return AllowConnections
Friend Property Let AllowConnections(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let AllowConnections(" & vData & ")", etFullDebug

  bAllowConnections = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get AllowConnections() As Boolean
Attribute AllowConnections.VB_Description = "Returns a value indicating whether this database is currently allowing connections to be made to it."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get AllowConnections()", etFullDebug

  AllowConnections = bAllowConnections
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Class_Terminate()", etFullDebug

  Set colLanguages = Nothing
  Set colNamespaces = Nothing
  Set colDatabaseVars = Nothing
  Set szTypeCache = Nothing
  Set szLanguageCache = Nothing
  Set szOperatorCache = Nothing
  Set szCommentCache = Nothing
  
  If cnDatabase.State <> adStateClosed Then
    objServer.iLogEvent "Closing connection to: " & GetDatabase(cnDatabase.ConnectionString), etMiniDebug
    cnDatabase.Close
  End If
  Set cnDatabase = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Execute a statement
Public Function Execute(ByVal szSQL As String, Optional ByVal RecordsAffected As Long, Optional ByVal Options As Long = -1, Optional QryType As QueryType = 1) As Recordset
Attribute Execute.VB_Description = "Execute an SQL query on the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Execute(" & QUOTE & szSQL & QUOTE & ", " & RecordsAffected & ", " & Options & ", " & QryType & ")", etFullDebug

  dbConnect
  Set Execute = objServer.ExecSQL(szSQL, cnDatabase, RecordsAffected, Options, QryType)
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Vacuum the Database
Public Sub Vacuum(Optional ByVal Analyse As Boolean, Optional ByVal Table As String, Optional ByVal Columns As String)
Attribute Vacuum.VB_Description = "Vacuum the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Vacuum(" & Analyse & ", " & QUOTE & Table & QUOTE & ", " & QUOTE & Columns & QUOTE & ")", etFullDebug

Dim szSQL As String

  dbConnect
  szSQL = "VACUUM"
  If Analyse = True Then szSQL = szSQL & " ANALYZE"
  If Table <> "" Then szSQL = szSQL & " " & Table
  If Table <> "" And Columns <> "" Then szSQL = szSQL & " (" & Columns & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Set/Return the DatabaseVarList
Friend Property Let DatabaseVarList(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let DatabaseVarList(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDatabaseVarList = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Get DatabaseVarList() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get DatabaseVarList()", etFullDebug

  DatabaseVarList = szDatabaseVarList
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Connect to the Database
Public Function dbConnect() As Boolean
Attribute dbConnect.VB_Description = "Open a connection to the database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.dbConnect()", etFullDebug

Dim szConnect As String

  'Don't connect if AllowConnections is False
  If (Not bAllowConnections) Or (lStatus = statInaccessible) Then
    lStatus = statInaccessible
    Exit Function
  End If
  
  'Open the connection if required
  If cnDatabase.State <> adStateOpen Then
    szConnect = Replace(objServer.Connection.ConnectionString, ";DATABASE=" & objServer.MasterDB, ";DATABASE=" & szName)
    If objServer.ShowPassword Then
      objServer.iLogEvent "Opening cnDatabase: " & szConnect, etMiniDebug
    Else
      objServer.iLogEvent "Opening cnDatabase: " & Replace(szConnect, ";PWD=" & objServer.iPassword, ";PWD=********"), etMiniDebug
    End If
    cnDatabase.Open szConnect
  End If
  lStatus = statOpen
  dbConnect = True
      
  Exit Function
Err_Handler:
  'If an error occurs, mark the database as inaccessible.
  lStatus = statInaccessible
  dbConnect = False
End Function

'Close the Database Connection
Public Sub dbDisconnect()
Attribute dbDisconnect.VB_Description = "Disconnect from the database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Sub dbDisconnect()", etFullDebug

  If objServer.ShowPassword Then
    objServer.iLogEvent "Closing cnDatabase: " & cnDatabase.ConnectionString, etMiniDebug
  Else
    objServer.iLogEvent "Closing cnDatabase: " & Replace(cnDatabase.ConnectionString, ";PWD=" & objServer.iPassword, ";PWD=********"), etMiniDebug
  End If
  lStatus = statClosed
  Set colLanguages = Nothing
  Set szTypeCache = Nothing
  Set szLanguageCache = Nothing
  Set szOperatorCache = Nothing
  Set szCommentCache = Nothing
  If cnDatabase.State <> adStateClosed Then cnDatabase.Close
  Set cnDatabase = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the objects Access Control List."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant permissions on the object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 12, App.Title, "Database security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclCreate) = aclCreate Then szSQL = szSQL & "CREATE, "
    If (Privilege And aclTemp) = aclTemp Then szSQL = szSQL & "TEMP, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON DATABASE " & FormattedID & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT datacl FROM pg_database WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!datacl & ""
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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 12, App.Title, "Database security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclCreate) = aclCreate Then szSQL = szSQL & "CREATE, "
    If (Privilege And aclTemp) = aclTemp Then szSQL = szSQL & "TEMP, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON DATABASE " & FormattedID & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT datacl FROM pg_database WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!datacl & ""
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

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the objects reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.SQL()", etFullDebug

Dim szSQL As String
Dim objVar As pgVar
  
  szSQL = "-- Database: " & FormattedID & vbCrLf & _
          "CREATE DATABASE " & FormattedID & " WITH"

  If szPath <> "" Then szSQL = szSQL & " LOCATION = '" & dbSZ(szPath) & "'"
  szSQL = szSQL & " ENCODING = '" & szServerEncoding & "'"
  szSQL = szSQL & ";" & vbCrLf
   
  'Database Vars/ACL
  If objVersion.VersionNum >= 7.3 Then
    For Each objVar In DatabaseVars
      szSQL = szSQL & "ALTER DATABASE " & FormattedID & " SET " & objVar.FormattedID & " = " & objVar.Value & ";" & vbCrLf
    Next objVar
    
    If szACL <> "" Then szSQL = szSQL & ParseACL(FormattedID, szACL, aclDatabase)
  End If
  
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON DATABASE " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Public Property Get Schema(Optional ResetSequences As Boolean) As String
Attribute Schema.VB_Description = "Returns the SQL representing the entire database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.Schema(" & ResetSequences & ")", etFullDebug
Dim szSQL As String
Dim vObject As Variant
Dim vChildObject As Variant
Dim vNamespaceObject As Variant
Dim arrObjects() As Variant
Dim X As Long
Dim Y As Long
Dim vTemp As Variant

  szSQL = "-- pgSchema:  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
          "-- Timestamp: " & Now & vbCrLf & _
          "-- Server:    " & objServer.Server & vbCrLf & _
          "-- Port:      " & objServer.Port & vbCrLf & _
          "-- Database:  " & szName & vbCrLf & vbCrLf
  
  'We will output the Schema object by object in OID order. This should work
  'most of the time unless (for example) a table is altered to use a later
  'created function as a default. Hopefully future releases of PostgreSQL
  'will include a pg_dependency table that we can use instead.
  
  'First, copy all objects into a single array: Aggregates, Functions,
  'Indexes, Languages, Operators, Rules, Sequences, Tables, Triggers,
  'Types & Views
  
  ReDim arrObjects(0)
  
  'Languages
  For Each vObject In Languages
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject
  
  'Namespaces
  For Each vNamespaceObject In Namespaces
    If (Not vNamespaceObject.SystemObject) Or (vNamespaceObject.Name = "public") Then
      If objVersion.VersionNum >= 7.3 Then
        Set arrObjects(UBound(arrObjects)) = vNamespaceObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
  
      'Aggregates
      For Each vObject In vNamespaceObject.Aggregates
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
      
      'Domains
      For Each vObject In vNamespaceObject.Domains
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
    
      'Functions
      For Each vObject In vNamespaceObject.Functions
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
    
      'Operators
      For Each vObject In vNamespaceObject.Operators
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
    
      'Sequences
      For Each vObject In vNamespaceObject.Sequences
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
    
      'Tables
      For Each vObject In vNamespaceObject.Tables
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        
          'Indexes
          For Each vChildObject In vObject.Indexes
            If Not vChildObject.SystemObject Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          Next vChildObject
          
          'Rules
          For Each vChildObject In vObject.Rules
            If Not vChildObject.SystemObject Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          Next vChildObject
          
          'Triggers
          For Each vChildObject In vObject.Triggers
            If Not vChildObject.SystemObject Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          Next vChildObject
        End If
      Next vObject
    
      'Types
      For Each vObject In vNamespaceObject.Types
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
    
      'Views
      For Each vObject In vNamespaceObject.Views
        If Not vObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vObject
  
    End If
  Next vNamespaceObject
  
  'Lose the last empty element
  If UBound(arrObjects) > 0 Then ReDim Preserve arrObjects(UBound(arrObjects) - 1)
  
  'Now bubble sort the array by OID.
  For X = UBound(arrObjects) To LBound(arrObjects) Step -1
    For Y = LBound(arrObjects) + 1 To X
      If arrObjects(Y - 1).OID > arrObjects(Y).OID Then
        Set vTemp = arrObjects(Y - 1)
        Set arrObjects(Y - 1) = arrObjects(Y)
        Set arrObjects(Y) = vTemp
      End If
    Next Y
  Next X

  'Now loop through the array and output the SQL
  For X = 0 To UBound(arrObjects)
    If arrObjects(X).ObjectType = "Sequence" Then
      szSQL = szSQL & arrObjects(X).SQL(ResetSequences) & vbCrLf
    Else
      szSQL = szSQL & arrObjects(X).SQL & vbCrLf
    End If
  Next X
    
  Schema = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
