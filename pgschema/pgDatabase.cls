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
Private colAggregates As Aggregates
Private colLanguages As Languages
Private colSequences As Sequences
Private colViews As Views
Private colTypes As Types
Private colFunctions As Functions
Private colOperators As Operators
Private colTables As Tables
Private colGraveyard As Graveyard

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
Private bRevisionControl As Boolean
Private bAllowConnections As Boolean
Private cnDatabase As New Connection

Private szTypeCache As Collection
Private szLanguageCache As Collection
Private szOperatorCache As Collection
Private szCommentCache As Collection

Friend Sub KillTypeCache()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.KillTypeCache()", etFullDebug

  objServer.iLogEvent "Invalidating Type cache (database: " & szName & ")...", etMiniDebug
  Set szTypeCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupType(lOID As Long) As String
On Error GoTo Err_Handler
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

  objServer.iLogEvent "Invalidating Language cache (database: " & szName & ")...", etMiniDebug
  Set szLanguageCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupLanguage(lOID As Long) As String
On Error GoTo Err_Handler
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

  objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
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
  If dbConnect = False Then Exit Function
  
  'Build the Operator Cache
  If szOperatorCache Is Nothing Then
    objServer.iLogEvent "Populating Operator cache (database: " & szName & ")...", etMiniDebug
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
    objServer.iLogEvent "Completed populating Operator cache (database: " & szName & ").", etMiniDebug
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

  objServer.iLogEvent "Invalidating Comment cache (database: " & szName & ")...", etMiniDebug
  Set szCommentCache = Nothing

  Exit Sub
Err_Handler: objServer.iLogError Err
End Sub

Friend Function LookupComment(lOID As Long, Optional lSubID As Long) As String
On Error GoTo Err_Handler
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
  On Error GoTo Err_Handler

  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Function

'Return the Aggregates Collection
Public Property Get Aggregates() As Aggregates
Attribute Aggregates.VB_Description = "Returns the Aggregates collection."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Aggregates()", etFullDebug

  If colAggregates Is Nothing Then
    If dbConnect = False Then Exit Property
    Set colAggregates = New Aggregates
    
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
    If dbConnect = False Then Exit Property
    Set colTables = New Tables
    
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
    If dbConnect = False Then Exit Property
    Set colOperators = New Operators
        
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
    If dbConnect = False Then Exit Property
    Set colFunctions = New Functions
        
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
    If dbConnect = False Then Exit Property
    Set colTypes = New Types
        
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
    If dbConnect = False Then Exit Property
    Set colViews = New Views
        
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

'Return the Sequences Collection
Public Property Get Sequences() As Sequences
Attribute Sequences.VB_Description = "Returns a collection of Sequences in the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Sequences()", etFullDebug

  If colSequences Is Nothing Then
    If dbConnect = False Then Exit Property
    Set colSequences = New Sequences
    
    'Pass the connection to the class
    Set colSequences.Connection = cnDatabase
    colSequences.Initialize
  End If
  Set Sequences = colSequences
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the History collection.
Public Property Get Graveyard() As Graveyard
Attribute Graveyard.VB_Description = "Returns the graveyard collection of dropped objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get History()", etFullDebug

  If RevisionControl Then
    If colGraveyard Is Nothing Then
      Set colGraveyard = New Graveyard
     
      'Pass the options to the class
      Set colGraveyard.Connection = cnDatabase
      colGraveyard.Initialize
    End If
    Set Graveyard = colGraveyard
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
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
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
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

'Set/Return the Database Server Encoding
Friend Property Let ServerEncoding(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let ServerEncoding(" & vData & ")", etFullDebug

  szServerEncoding = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ServerEncoding() As String
Attribute ServerEncoding.VB_Description = "Returns the name of the encoding scheme."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get ServerEncoding()", etFullDebug

  ServerEncoding = szServerEncoding
    
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
  If dbConnect = False Then Exit Property
  
  szSQL = "COMMENT ON DATABASE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Status()", etFullDebug

  Status = lStatus
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Database Revision Logging
Public Property Let RevisionControl(ByVal vData As Boolean)
Attribute RevisionControl.VB_Description = "Sets/Returns whether or not Revision Logging is active."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let RevisionControl(" & vData & ")", etFullDebug

Dim vObject As Variant
Dim vChildObject As Variant

  If dbConnect = False Then Exit Property
  If vData Then
    If Not RevisionControl Then
      bRevisionControl = True
      objServer.ExecSQL SQL_CREATE_REVLOG, cnDatabase
      objServer.ExecSQL "ALTER TABLE pgadmin_rclog OWNER TO " & QUOTE & Owner & QUOTE, cnDatabase
      
      'Log the current database
      'Aggregates
      For Each vObject In Aggregates
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Functions
      For Each vObject In Functions
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Languages
      For Each vObject In Languages
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Operators
      For Each vObject In Operators
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Sequences
      For Each vObject In Sequences
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Tables
      For Each vObject In Tables
        If Not vObject.SystemObject Then
          vObject.Commit rcAdd, "Revision logging enabled."
        
          'Indexes
          For Each vChildObject In vObject.Indexes
            If Not vChildObject.SystemObject Then vChildObject.Commit rcAdd, "Revision logging enabled."
          Next vChildObject
          
          'Rules
          For Each vChildObject In vObject.Rules
            If Not vChildObject.SystemObject Then vChildObject.Commit rcAdd, "Revision logging enabled."
          Next vChildObject
         
          'Triggers
          For Each vChildObject In vObject.Triggers
            If Not vChildObject.SystemObject Then vChildObject.Commit rcAdd, "Revision logging enabled."
          Next vChildObject
        End If
      Next vObject
    
      'Types
      For Each vObject In Types
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    
      'Views
      For Each vObject In Views
        If Not vObject.SystemObject Then vObject.Commit rcAdd, "Revision logging enabled."
      Next vObject
    End If
  Else
    If RevisionControl Then
      bRevisionControl = False
      objServer.ExecSQL SQL_DROP_REVLOG, cnDatabase
    End If
  End If
  
  'Kill the tables collection and comment cache
  Set colTables = Nothing
  KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get RevisionControl() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get RevisionControl()", etFullDebug

Static bCached As Boolean
Dim szSQL As String
Dim rs As New Recordset

  If dbConnect = False Then Exit Property
  If Not bCached Then
    szSQL = "SELECT * FROM pg_class WHERE relname = 'pgadmin_rclog'"
    Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    If rs.EOF Then
      bRevisionControl = False
    Else
      bRevisionControl = True
    End If
  End If
  
  RevisionControl = bRevisionControl
  bCached = True
  Set rs = Nothing
    
  Exit Property
Err_Handler:
  Set rs = Nothing
  objServer.iLogError Err
End Property

'Set/Return AllowConnections
Friend Property Let AllowConnections(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let AllowConnections(" & vData & ")", etFullDebug

  bAllowConnections = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get AllowConnections() As Boolean
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get AllowConnections()", etFullDebug

  AllowConnections = bAllowConnections
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Friend Sub LogRevision(szType As String, szIdentifier As String, lObjOID As Long, szDefinition As String, szComment As String, rlAction As rcLogType, Optional szParentTable As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.LogRevision(" & QUOTE & szType & QUOTE & ", " & QUOTE & szIdentifier & QUOTE & ", " & lOID & ", " & QUOTE & szDefinition & QUOTE & ", " & QUOTE & szComment & QUOTE & ", " & rlAction & ", " & QUOTE & szParentTable & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim szAction As String

  If dbConnect = False Then Exit Sub
  If Not RevisionControl Then Exit Sub
  szType = dbSZ(szType)
  szIdentifier = dbSZ(szIdentifier)
  szDefinition = dbSZ(szDefinition)
  szComment = dbSZ(szComment)
  Select Case rlAction
    Case rcAdd
      szAction = "A"
    Case rcUpdate
      szAction = "U"
    Case rcDelete
      szAction = "D"
  End Select
  szSQL = "INSERT INTO pgadmin_rclog(rc_action, rc_type, rc_identifier, rc_oid, rc_table, rc_version, rc_definition, rc_comment) VALUES ('" & szAction & "', '" & szType & "', '" & szIdentifier & "', " & lObjOID & ", '" & dbSZ(szParentTable) & "', (SELECT CASE WHEN max(rc_version) IS NULL THEN 0 ELSE max(rc_version) +1::int4 END from pgadmin_rclog WHERE rc_type = '" & szType & "' AND rc_identifier = '" & szIdentifier & "'), '" & szDefinition & "', '" & szComment & "');"
  objServer.ExecSQL szSQL, cnDatabase
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

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
Public Function Execute(ByVal szSQL As String, Optional ByVal RecordsAffected As Long, Optional ByVal Options As Long = -1, Optional QryType As QueryType = qrySystem) As Recordset
Attribute Execute.VB_Description = "Execute an SQL query on the Database."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Execute(" & QUOTE & szSQL & QUOTE & ", " & RecordsAffected & ", " & Options & ", " & QryType & ")", etFullDebug

  dbConnect
  Set Execute = objServer.ExecSQL(szSQL, cnDatabase, RecordsAffected, Options, QryType)
  
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
Attribute SQL.VB_Description = "Returns the objects reverse engineered SQL."
On Error GoTo Err_Handler
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

Public Property Get Schema(Optional ResetSequences As Boolean) As String
Attribute Schema.VB_Description = "Retunrs the database schema a an SQL script."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.Schema(" & ResetSequences & ")", etFullDebug
Dim szSQL As String
Dim vObject As Variant
Dim vChildObject As Variant
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
  
  'Aggregates
  For Each vObject In Aggregates
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Functions
  For Each vObject In Functions
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Languages
  For Each vObject In Languages
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Operators
  For Each vObject In Operators
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Sequences
  For Each vObject In Sequences
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Tables
  For Each vObject In Tables
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
  For Each vObject In Types
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Views
  For Each vObject In Views
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject
  
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
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit all objects to Revision Control."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

Dim objAggregate As pgAggregate
Dim objFunction As pgFunction
Dim objLanguage As pgLanguage
Dim objOperator As pgOperator
Dim objSequence As pgSequence
Dim objTable As pgTable
Dim objType As pgType
Dim objView As pgView

  For Each objAggregate In Aggregates
    If Not objAggregate.SystemObject Then objAggregate.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objAggregate
  For Each objFunction In Functions
    If Not objFunction.SystemObject Then objFunction.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objFunction
  For Each objLanguage In Languages
    If Not objLanguage.SystemObject Then objLanguage.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objLanguage
  For Each objOperator In Operators
    If Not objOperator.SystemObject Then objOperator.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objOperator
  For Each objSequence In Sequences
    If Not objSequence.SystemObject Then objSequence.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objSequence
  For Each objTable In Tables
    If Not objTable.SystemObject Then objTable.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objTable
  For Each objType In Types
    If Not objType.SystemObject Then objType.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objType
  For Each objView In Views
    If Not objView.SystemObject Then objView.Commit UpdateType, "(Database committed)" & vbCrLf & vbCrLf & UpdateComment
  Next objView
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

