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

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szPath As String
Private szOwner As String
Private lServerEncoding As Long
Private szServerEncoding As String
Private lStatus As dbStatus
Private bAllowConnections As Boolean
Private cnDatabase As New Connection

Private szTypeCache As Collection
Private szLanguageCache As Collection
Private szOperatorCache As Collection
Private szCommentCache As Collection

Friend Sub KillTypeCache()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillTypeCache()", etFullDebug

  objServer.iLogEvent "Invalidating Type cache (database: " & szName & ")...", etMiniDebug
  Set szTypeCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupType(lOID As Long) As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupType(" & lOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Populate the cache
  If szTypeCache Is Nothing Then
    objServer.iLogEvent "Populating Type cache (database: " & szName & ")...", etMiniDebug
    Set szTypeCache = New Collection
    Set rs = objServer.ExecSQL("SELECT oid, typname FROM pg_type", cnDatabase)
    While Not rs.EOF
      szTypeCache.Add rs!typname & "", "T" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Completed populating Type cache (database: " & szName & ").", etMiniDebug
  End If
  
  If lOID = 0 Then
    LookupType = "opaque"
  Else
    On Error Resume Next
    LookupType = szTypeCache("T" & lOID)
    'On Error Goto Err_Handler
  End If
 
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillLanguageCache()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillLanguageCache()", etFullDebug

  objServer.iLogEvent "Invalidating Language cache (database: " & szName & ")...", etMiniDebug
  Set szLanguageCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupLanguage(lOID As Long) As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupLanguage(" & lOID & ")", etFullDebug
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
  LookupLanguage = szLanguageCache("L" & lOID)
  'On Error Goto Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillOperatorCache()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillOperatorCache()", etFullDebug

  objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
  Set szOperatorCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupOperator(lOID As Long) As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupOperator(" & lOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  If dbConnect = False Then Exit Function
  
  'Build the Operator Cache
  If szOperatorCache Is Nothing Then
    objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
    Set szOperatorCache = New Collection
    'Select an O in front of the name to avoid strange behaviour with !
    Set rs = objServer.ExecSQL(SQL_GET_OPERATORS, cnDatabase)
    While Not rs.EOF
      'Prefix the operator name otherwise we get bizarre errors
      szOperatorCache.Add "O" & rs!oprname, "O" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    objServer.iLogEvent "Completed populating Operator cache (database: " & szName & ").", etMiniDebug
  End If
  
  On Error Resume Next
  LookupOperator = Mid(szOperatorCache("O" & lOID), 2)
  'On Error Goto Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillCommentCache()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillCommentCache()", etFullDebug

  objServer.iLogEvent "Invalidating Comment cache (database: " & szName & ")...", etMiniDebug
  Set szCommentCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupComment(lOID As Long, Optional lSubID As Long) As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupComment(" & lOID & ")", etFullDebug
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
    If objServer.dbVersion.VersionNum < 7.2 Then
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
  
  On Error Resume Next
  LookupComment = szCommentCache("C" & lOID & "S" & lSubID)
  'On Error Goto Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the Languages Collection
Public Property Get Languages() As Languages
Attribute Languages.VB_Description = "Returns a collection of Languages in the Database."
'On Error Goto Err_Handler
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
Err_Handler:  objServer.iLogError Err
End Property

'Return the Namespaces Collection
Public Property Get Namespaces() As Namespaces
'On Error Goto Err_Handler
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
Err_Handler:  objServer.iLogError Err
End Property

'Return the ADO Connection String
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_Description = "Returns the database connection string."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ConnectionString()", etFullDebug

  If cnDatabase.State <> adStateOpen Then
    ConnectionString = ""
  Else
    ConnectionString = cnDatabase.ConnectionString
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the OID of the Database."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Database Name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ObjectType()", etFullDebug

  ObjectType = "Database"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Path
Friend Property Let Path(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Path(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szPath = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns the Database path."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Path()", etFullDebug

  Path = szPath
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Database owner."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Database Server Encoding
Friend Property Let ServerEncoding(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let ServerEncoding(" & vData & ")", etFullDebug

  szServerEncoding = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ServerEncoding() As String
Attribute ServerEncoding.VB_Description = "Returns the name of the encoding scheme."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ServerEncoding()", etFullDebug

  ServerEncoding = szServerEncoding
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
Attribute Comment.VB_Description = "Sets/Returns the Database comment."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Comment(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Connect if required
  If dbConnect = False Then Exit Property
  
  szSQL = "COMMENT ON DATABASE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Comment()", etFullDebug
  
  If objServer.DeferConnection And lStatus <> statOpen Then
    Comment = ""
  Else
    Comment = LookupComment(lOID)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Status
Public Property Get Status() As dbStatus
Attribute Status.VB_Description = "Returns the database connection status."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Status()", etFullDebug

  Status = lStatus
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return AllowConnections
Friend Property Let AllowConnections(ByVal vData As Boolean)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let AllowConnections(" & vData & ")", etFullDebug

  bAllowConnections = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get AllowConnections() As Boolean
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get AllowConnections()", etFullDebug

  AllowConnections = bAllowConnections
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Cleanup
Private Sub Class_Terminate()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Class_Terminate()", etFullDebug

  Set colLanguages = Nothing
  Set colNamespaces = Nothing
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
Err_Handler:  objServer.iLogError Err
End Sub

'Execute a statement
Public Function Execute(ByVal szSQL As String, Optional ByVal RecordsAffected As Long, Optional ByVal Options As Long = -1, Optional QryType As QueryType = 1) As Recordset
Attribute Execute.VB_Description = "Execute an SQL query on the Database."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Execute(" & QUOTE & szSQL & QUOTE & ", " & RecordsAffected & ", " & Options & ", " & QryType & ")", etFullDebug

  dbConnect
  Set Execute = objServer.ExecSQL(szSQL, cnDatabase, RecordsAffected, Options, QryType)
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Vacuum the Database
Public Sub Vacuum(Optional ByVal Analyse As Boolean, Optional ByVal Table As String, Optional ByVal Columns As String)
Attribute Vacuum.VB_Description = "Vacuum the Database."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Vacuum(" & Analyse & ", " & QUOTE & Table & QUOTE & ", " & QUOTE & Columns & QUOTE & ")", etFullDebug

Dim szSQL As String

  dbConnect
  szSQL = "VACUUM"
  If Analyse = True Then szSQL = szSQL & " ANALYZE"
  If Table <> "" Then szSQL = szSQL & " " & Table
  If Table <> "" And Columns <> "" Then szSQL = szSQL & " (" & Columns & ")"
  objServer.ExecSQL szSQL, cnDatabase
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Connect to the Database
Public Function dbConnect() As Boolean
Attribute dbConnect.VB_Description = "Open a connection to the database."
'On Error Goto Err_Handler
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
'On Error Goto Err_Handler
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
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the objects reverse engineered SQL."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Database: " & szName & vbCrLf & _
          "CREATE DATABASE " & QUOTE & szName & QUOTE & " WITH"

  If szPath <> "" Then szSQL = szSQL & " LOCATION = '" & dbSZ(szPath) & "'"
  szSQL = szSQL & " ENCODING = '" & szServerEncoding & "'"
  szSQL = szSQL & ";" & vbCrLf
  
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON DATABASE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property


