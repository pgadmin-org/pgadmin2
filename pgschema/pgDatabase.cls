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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colAggregates As Aggregates
Private colLanguages As Languages
Private colSequences As Sequences
Private colViews As Views
Private colTypes As Types
Private colFunctions As Functions
Private colOperators As Operators
Private colTables As Tables

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szPath As String
Private szOwner As String
Private lEncoding As Long
Private cnDatabase As New Connection

Private szTypeCache As Collection
Private szLanguageCache As Collection
Private szOperatorCache As Collection
Private szCommentCache As Collection

Friend Sub KillTypeCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillTypeCache()", etFullDebug

  Set szTypeCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupType(lOID As Long) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupType(" & lOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  dbConnect
  
  'Get the Return Type
  If szTypeCache Is Nothing Then
    Set szTypeCache = New Collection
    Set rs = objServer.ExecSQL("SELECT oid, typname FROM pg_type", cnDatabase)
    While Not rs.EOF
      szTypeCache.Add rs!typname & "", "T" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
  End If
  
  If lOID = 0 Then
    LookupType = "opaque"
  Else
    On Error Resume Next
    LookupType = szTypeCache("T" & lOID)
    On Error GoTo Err_Handler
  End If
 
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillLanguageCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillLanguageCache()", etFullDebug

  Set szLanguageCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupLanguage(lOID As Long) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupLanguage(" & lOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  dbConnect
  
  'Get the Function Language
  If szLanguageCache Is Nothing Then
    Set szLanguageCache = New Collection
    Set rs = objServer.ExecSQL(SQL_GET_LANGUAGES, cnDatabase)
    While Not rs.EOF
      szLanguageCache.Add rs!lanname & "", "L" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
  End If
  
  On Error Resume Next
  LookupLanguage = szLanguageCache("L" & lOID)
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillOperatorCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillOperatorCache()", etFullDebug

  Set szOperatorCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupOperator(lOID As Long) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupOperator(" & lOID & ")", etFullDebug
Dim rs As New Recordset
Dim szLOT As String
Dim szROT As String

  'Connect if required
  dbConnect
  
  'Build the Operator Cache
  If szOperatorCache Is Nothing Then
    Set szOperatorCache = New Collection
    'Select an O in front of the name to avoid strange behaviour with !
    Set rs = objServer.ExecSQL(SQL_GET_OPERATORS, cnDatabase)
    While Not rs.EOF
      szLOT = LookupType(rs!oprleft & "")
      szROT = LookupType(rs!oprright & "")
      If szLOT = "" Then szLOT = "NONE"
      If szROT = "" Then szROT = "NONE"
      szOperatorCache.Add rs!oprname & "(" & szLOT & ", " & szROT & ")", "O" & rs!OID
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
  End If
  
  On Error Resume Next
  LookupOperator = szOperatorCache("O" & lOID)
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

Friend Sub KillCommentCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillCommentCache()", etFullDebug

  Set szCommentCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupComment(lOID As Long) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LookupComment(" & lOID & ")", etFullDebug
Dim rs As New Recordset

  'Connect if required
  dbConnect
  
  'Build the Comment Cache
  If szCommentCache Is Nothing Then
    Set szCommentCache = New Collection
    Set rs = objServer.ExecSQL("SELECT * FROM pg_description", cnDatabase)
    While Not rs.EOF
      szCommentCache.Add rs!Description & "", "C" & rs!objoid
      rs.MoveNext
    Wend
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
  End If
  
  On Error Resume Next
  LookupComment = szCommentCache("C" & lOID)
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the Aggregates Collection
Public Property Get Aggregates() As Aggregates
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Aggregates()", etFullDebug

  If colAggregates Is Nothing Then
    Set colAggregates = New Aggregates
    dbConnect
    
    'Pass the connection to the class
    Set colAggregates.Connection = cnDatabase
    colAggregates.Initialize
  End If
  Set Aggregates = colAggregates
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Tables Collection
Public Property Get Tables() As Tables
Attribute Tables.VB_Description = "Returns a collection of Tables."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Tables()", etFullDebug

  If colTables Is Nothing Then
    Set colTables = New Tables
    dbConnect
    
    'Pass the connection to the class
    Set colTables.Connection = cnDatabase
    colTables.Initialize
  End If
  Set Tables = colTables
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Operators Collection
Public Property Get Operators() As Operators
Attribute Operators.VB_Description = "Returns a collection of Operators."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Operators()", etFullDebug

  If colOperators Is Nothing Then
    Set colOperators = New Operators
    dbConnect
    
    'Pass the connection to the class
    Set colOperators.Connection = cnDatabase
    colOperators.Initialize
  End If
  Set Operators = colOperators
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Functions Collection
Public Property Get Functions() As Functions
Attribute Functions.VB_Description = "Returns a collection of Functions."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Functions()", etFullDebug

  If colFunctions Is Nothing Then
    Set colFunctions = New Functions
    dbConnect
    
    'Pass the connection to the class
    Set colFunctions.Connection = cnDatabase
    colFunctions.Initialize
  End If
  Set Functions = colFunctions
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Types Collection
Public Property Get Types() As Types
Attribute Types.VB_Description = "Returns a collection of Types in the Database"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Types()", etFullDebug

  If colTypes Is Nothing Then
    Set colTypes = New Types
    dbConnect
    
    'Pass the connection to the class
    Set colTypes.Connection = cnDatabase
    colTypes.Initialize
  End If
  Set Types = colTypes
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Views Collection
Public Property Get Views() As Views
Attribute Views.VB_Description = "Returns a collection of Views in the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Views()", etFullDebug

  If colViews Is Nothing Then
    Set colViews = New Views
    dbConnect
    
    'Pass the connection to the class
    Set colViews.Connection = cnDatabase
    colViews.Initialize
  End If
  Set Views = colViews
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Languages Collection
Public Property Get Languages() As Languages
Attribute Languages.VB_Description = "Returns a collection of Languages in the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Languages()", etFullDebug

  If colLanguages Is Nothing Then
    Set colLanguages = New Languages
    dbConnect
    
    'Pass the connection to the class
    Set colLanguages.Connection = cnDatabase
    colLanguages.Initialize
  End If
  Set Languages = colLanguages
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Sequences Collection
Public Property Get Sequences() As Sequences
Attribute Sequences.VB_Description = "Returns a collection of Sequences in the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Sequences()", etFullDebug

  If colSequences Is Nothing Then
    Set colSequences = New Sequences
    dbConnect
    
    'Pass the connection to the class
    Set colSequences.Connection = cnDatabase
    colSequences.Initialize
  End If
  Set Sequences = colSequences
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the ADO Connection String
Public Property Get ConnectionString() As String
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the OID of the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Database Name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ObjectType()", etFullDebug

  ObjectType = "Database"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Path
Friend Property Let Path(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Path(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szPath = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns the Database path."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Path()", etFullDebug

  Path = szPath
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Database owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Database Encoding
Friend Property Let Encoding(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Encoding(" & vData & ")", etFullDebug

  lEncoding = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Encoding() As Long
Attribute Encoding.VB_Description = "Returns the Database Encoding."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Encoding()", etFullDebug

  Encoding = lEncoding
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
Attribute Comment.VB_Description = "Sets/Returns the Database comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Comment(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim szSQL As String

  'Connect if required
  dbConnect
  
  szSQL = "COMMENT ON DATABASE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Comment()", etFullDebug
  
  Comment = LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Class_Terminate()", etFullDebug

  Set colAggregates = Nothing
  Set colLanguages = Nothing
  Set colSequences = Nothing
  Set colViews = Nothing
  Set colTypes = Nothing
  Set colFunctions = Nothing
  Set colOperators = Nothing
  Set colTables = Nothing
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
Public Function Execute(ByVal szSQL As String, Optional ByVal RecordsAffected As Long, Optional ByVal Options As Long = -1) As Recordset
Attribute Execute.VB_Description = "Execute an SQL query on the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Execute(" & QUOTE & szSQL & QUOTE & ", " & RecordsAffected & ", " & Options & ")", etFullDebug

  dbConnect
  Set Execute = objServer.ExecSQL(szSQL, cnDatabase, RecordsAffected, Options)
  
  Exit Function
Err_Handler:  objServer.iLogError Err
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
Err_Handler:  objServer.iLogError Err
End Sub

'Connect to the Database
Friend Sub dbConnect()
Attribute dbConnect.VB_Description = "Open a connection to the database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.dbConnect()", etFullDebug

Dim szConnect As String

  'Open the connection if required
  If cnDatabase.State <> adStateOpen Then
    szConnect = Replace(objServer.Connection.ConnectionString, ";DATABASE=template1", ";DATABASE=" & szName)
    If objServer.ShowPassword Then
      objServer.iLogEvent "Opening cnDatabase: " & szConnect, etMiniDebug
    Else
      objServer.iLogEvent "Opening cnDatabase: " & Replace(szConnect, ";PWD=" & objServer.iPassword, ";PWD=********"), etMiniDebug
    End If
    cnDatabase.Open szConnect
  End If
      
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Close the Database Connection
Friend Sub dbDisconnect()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Sub dbDisconnect()", etFullDebug

  If objServer.ShowPassword Then
    objServer.iLogEvent "Closing cnDatabase: " & cnDatabase.ConnectionString, etMiniDebug
  Else
    objServer.iLogEvent "Closing cnDatabase: " & Replace(cnDatabase.ConnectionString, ";PWD=" & objServer.iPassword, ";PWD=********"), etMiniDebug
  End If
  Set colAggregates = Nothing
  Set colLanguages = Nothing
  Set colSequences = Nothing
  Set colViews = Nothing
  Set colTypes = Nothing
  Set colFunctions = Nothing
  Set colOperators = Nothing
  Set colTables = Nothing
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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Database: " & szName & vbCrLf & _
          "CREATE DATABASE " & QUOTE & szName & QUOTE
  
  If (szPath <> "") Or (lEncoding) <> 0 Then szSQL = szSQL & " WITH"
  If szPath <> "" Then szSQL = szSQL & " LOCATION = '" & dbSZ(szPath) & "'"
  If lEncoding <> 0 Then szSQL = szSQL & " ENCODING = '" & lEncoding & "'"
  szSQL = szSQL & ";" & vbCrLf
  
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON DATABASE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property
