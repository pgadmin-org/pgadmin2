Attribute VB_Name = "basVarDb"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basVarDb.bas - Contains type var of database

Option Explicit

'type variable database
Public Enum TypeVarDb
  TVDB_FLOAT
  TVDB_INTEGR
  TVDB_BOOLEAN
  TVDB_STRING
  TVDB_CAST
End Enum

'detail variable
Public Type VarDb
  Name As String
  Type As TypeVarDb
  CastValue As New Collection     'value variable
End Type

'buffer for copy setting var
Public ColVarDbBuffer As Collection

Private VariableDb() As VarDb

'Initalization
Public Sub InitVarDb()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basVarDb.InitVarDb()", etFullDebug

Dim ColTmp As Collection

  ReDim VariableDb(0) As VarDb
  Set ColVarDbBuffer = New Collection

  'Planner and Optimizer Tuning
  AddVarDb "CPU_INDEX_TUPLE_COST", TVDB_FLOAT
  AddVarDb "CPU_OPERATOR_COST", TVDB_FLOAT
  AddVarDb "CPU_TUPLE_COST", TVDB_FLOAT
  AddVarDb "DEFAULT_STATISTICS_TARGET", TVDB_INTEGR
  AddVarDb "EFFECTIVE_CACHE_SIZE", TVDB_FLOAT
  AddVarDb "ENABLE_HASHJOIN", TVDB_BOOLEAN
  AddVarDb "ENABLE_INDEXSCAN", TVDB_BOOLEAN
  AddVarDb "ENABLE_MERGEJOIN", TVDB_BOOLEAN
  AddVarDb "ENABLE_NESTLOOP", TVDB_BOOLEAN
  AddVarDb "ENABLE_SEQSCAN", TVDB_BOOLEAN
  AddVarDb "ENABLE_SORT", TVDB_BOOLEAN
  AddVarDb "ENABLE_TIDSCAN", TVDB_BOOLEAN
  AddVarDb "GEQO", TVDB_BOOLEAN
  AddVarDb "GEQO_EFFORT", TVDB_INTEGR
  AddVarDb "GEQO_GENERATIONS", TVDB_INTEGR
  AddVarDb "GEQO_POOL_SIZE", TVDB_INTEGR
  AddVarDb "GEQO_RANDOM_SEED", TVDB_INTEGR
  AddVarDb "GEQO_SELECTION_BIAS", TVDB_FLOAT
  AddVarDb "GEQO_THRESHOLD", TVDB_INTEGR
  AddVarDb "RANDOM_PAGE_COST", TVDB_FLOAT

  'Logging and Debugging
  Set ColTmp = New Collection
  ColTmp.Add "DEBUG5"
  ColTmp.Add "DEBUG4"
  ColTmp.Add "DEBUG3"
  ColTmp.Add "DEBUG2"
  ColTmp.Add "DEBUG1"
  ColTmp.Add "INFO"
  ColTmp.Add "NOTICE"
  ColTmp.Add "WARNING"
  ColTmp.Add "ERROR"
  ColTmp.Add "LOG"
  ColTmp.Add "FATAL"
  ColTmp.Add "PANIC"
  AddVarDb "SERVER_MIN_MESSAGES", TVDB_CAST, ColTmp

  Set ColTmp = New Collection
  ColTmp.Add "DEBUG5"
  ColTmp.Add "DEBUG4"
  ColTmp.Add "DEBUG3"
  ColTmp.Add "DEBUG2"
  ColTmp.Add "DEBUG1"
  ColTmp.Add "INFO"
  ColTmp.Add "NOTICE"
  ColTmp.Add "WARNING"
  ColTmp.Add "ERROR"
  AddVarDb "CLIENT_MIN_MESSAGES", TVDB_CAST, ColTmp

  AddVarDb "DEBUG_ASSERTIONS", TVDB_BOOLEAN
  AddVarDb "DEBUG_PRINT_PARSE", TVDB_BOOLEAN
  AddVarDb "DEBUG_PRINT_REWRITTEN", TVDB_BOOLEAN
  AddVarDb "DEBUG_PRINT_PLAN", TVDB_BOOLEAN
  AddVarDb "DEBUG_PRETTY_PRINT", TVDB_BOOLEAN
  
  AddVarDb "EXPLAIN_PRETTY_PRINT", TVDB_BOOLEAN
  AddVarDb "HOSTNAME_LOOKUP", TVDB_BOOLEAN
  AddVarDb "LOG_CONNECTIONS", TVDB_BOOLEAN
  AddVarDb "LOG_DURATION", TVDB_BOOLEAN

  Set ColTmp = New Collection
  ColTmp.Add "DEBUG5"
  ColTmp.Add "DEBUG4"
  ColTmp.Add "DEBUG3"
  ColTmp.Add "DEBUG2"
  ColTmp.Add "DEBUG1"
  ColTmp.Add "INFO"
  ColTmp.Add "NOTICE"
  ColTmp.Add "WARNING"
  ColTmp.Add "ERROR"
  ColTmp.Add "FATAL"
  ColTmp.Add "PANIC"
  AddVarDb "LOG_MIN_ERROR_STATEMENT", TVDB_CAST, ColTmp

  AddVarDb "LOG_PID", TVDB_BOOLEAN
  AddVarDb "LOG_STATEMENT", TVDB_BOOLEAN
  AddVarDb "LOG_TIMESTAMP", TVDB_BOOLEAN
  AddVarDb "SHOW_STATEMENT_STATS", TVDB_BOOLEAN
  AddVarDb "SHOW_PARSER_STATS", TVDB_BOOLEAN
  AddVarDb "SHOW_PLANNER_STATS", TVDB_BOOLEAN
  AddVarDb "SHOW_EXECUTOR_STATS", TVDB_BOOLEAN
  AddVarDb "SHOW_SOURCE_PORT", TVDB_BOOLEAN
  AddVarDb "STATS_COMMAND_STRING", TVDB_BOOLEAN

  AddVarDb "STATS_BLOCK_LEVEL", TVDB_BOOLEAN
  AddVarDb "STATS_ROW_LEVEL", TVDB_BOOLEAN
  AddVarDb "STATS_RESET_ON_SERVER_START", TVDB_BOOLEAN
  AddVarDb "STATS_START_COLLECTOR", TVDB_BOOLEAN

  Set ColTmp = New Collection
  ColTmp.Add "0"
  ColTmp.Add "1"
  ColTmp.Add "2"
  AddVarDb "SYSLOG", TVDB_CAST, ColTmp

  Set ColTmp = New Collection
  ColTmp.Add "LOCAL0"
  ColTmp.Add "LOCAL1"
  ColTmp.Add "LOCAL2"
  ColTmp.Add "LOCAL3"
  ColTmp.Add "LOCAL4"
  ColTmp.Add "LOCAL5"
  ColTmp.Add "LOCAL6"
  ColTmp.Add "LOCAL7"
  AddVarDb "SYSLOG_FACILITY", TVDB_CAST, ColTmp

  AddVarDb "SYSLOG_IDENT", TVDB_STRING
  AddVarDb "TRACE_NOTIFY", TVDB_BOOLEAN

  'General Operation
  AddVarDb "AUTOCOMMIT", TVDB_BOOLEAN
  AddVarDb "AUSTRALIAN_TIMEZONES", TVDB_BOOLEAN
  AddVarDb "AUTHENTICATION_TIMEOUT", TVDB_INTEGR
  AddVarDb "CLIENT_ENCODING", TVDB_STRING
  AddVarDb "DATESTYLE", TVDB_STRING
  AddVarDb "DB_USER_NAMESPACE", TVDB_BOOLEAN
  AddVarDb "DEADLOCK_TIMEOUT", TVDB_INTEGR

  Set ColTmp = New Collection
  ColTmp.Add "read committed"
  ColTmp.Add "serializable"
  AddVarDb "DEFAULT_TRANSACTION_ISOLATION", TVDB_CAST, ColTmp

  AddVarDb "DYNAMIC_LIBRARY_PATH", TVDB_STRING
  AddVarDb "KRB_SERVER_KEYFILE", TVDB_STRING
  AddVarDb "FSYNC", TVDB_BOOLEAN
  AddVarDb "LC_MESSAGES", TVDB_STRING
  AddVarDb "LC_MONETARY", TVDB_STRING
  AddVarDb "LC_NUMERIC", TVDB_STRING
  AddVarDb "LC_TIME", TVDB_STRING
  AddVarDb "MAX_CONNECTIONS", TVDB_INTEGR
  AddVarDb "MAX_EXPR_DEPTH", TVDB_INTEGR
  AddVarDb "MAX_FILES_PER_PROCESS", TVDB_INTEGR
  AddVarDb "MAX_FSM_RELATIONS", TVDB_INTEGR
  AddVarDb "MAX_FSM_PAGES", TVDB_INTEGR
  AddVarDb "MAX_LOCKS_PER_TRANSACTION", TVDB_INTEGR
  AddVarDb "PASSWORD_ENCRYPTION", TVDB_BOOLEAN
  AddVarDb "PORT", TVDB_INTEGR
  AddVarDb "SEARCH_PATH", TVDB_STRING
  AddVarDb "STATEMENT_TIMEOUT", TVDB_INTEGR
  AddVarDb "SHARED_BUFFERS", TVDB_INTEGR
  AddVarDb "SILENT_MODE", TVDB_BOOLEAN
  AddVarDb "SORT_MEM", TVDB_INTEGR
  AddVarDb "SQL_INHERITANCE", TVDB_BOOLEAN
  AddVarDb "SSL", TVDB_BOOLEAN
  AddVarDb "SUPERUSER_RESERVED_CONNECTIONS", TVDB_INTEGR
  AddVarDb "TCPIP_SOCKET", TVDB_BOOLEAN
  AddVarDb "TIMEZONE", TVDB_STRING
  AddVarDb "TRANSFORM_NULL_EQUALS", TVDB_BOOLEAN
  AddVarDb "UNIX_SOCKET_DIRECTORY", TVDB_STRING
  AddVarDb "UNIX_SOCKET_GROUP", TVDB_STRING
  AddVarDb "UNIX_SOCKET_PERMISSIONS", TVDB_INTEGR
  AddVarDb "VACUUM_MEM", TVDB_INTEGR
  AddVarDb "VIRTUAL_HOST", TVDB_STRING

  'WAL
  AddVarDb "CHECKPOINT_SEGMENTS", TVDB_INTEGR
  AddVarDb "CHECKPOINT_TIMEOUT", TVDB_INTEGR
  AddVarDb "COMMIT_DELAY", TVDB_INTEGR
  AddVarDb "COMMIT_SIBLINGS", TVDB_INTEGR
  AddVarDb "WAL_BUFFERS", TVDB_INTEGR
  AddVarDb "WAL_DEBUG", TVDB_INTEGR

  Set ColTmp = New Collection
  ColTmp.Add "FSYNC"
  ColTmp.Add "FDATASYNC"
  ColTmp.Add "OPEN_SYNC"
  ColTmp.Add "OPEN_DATASYNC"
  AddVarDb "WAL_SYNC_METHOD", TVDB_CAST, ColTmp

  ReDim Preserve VariableDb(UBound(VariableDb) - 1) As VarDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basVarDb.InitVarDb"
End Sub
'Add var db to Collection
Private Sub AddVarDb(szName As String, TypeVar As TypeVarDb, Optional CastValue As Collection = Nothing)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basVarDb.AddVarDb(" & szName & "," & TypeVar & ")", etFullDebug

Dim iNumEl As Integer
  
  iNumEl = UBound(VariableDb)
  VariableDb(iNumEl).Name = szName
  VariableDb(iNumEl).Type = TypeVar
  If Not CastValue Is Nothing Then Set VariableDb(iNumEl).CastValue = CastValue
  
  ReDim Preserve VariableDb(iNumEl + 1) As VarDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basVarDb.AddVarDb"
End Sub

'return the definition var db from Collection
Public Function GetVarDb(szName As String) As VarDb
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basVarDb.GetVarDb(" & szName & ")", etFullDebug
Dim ii As Integer
Dim DummyVarDb As VarDb

  'find variable
  For ii = 0 To UBound(VariableDb)
    If LCase(VariableDb(ii).Name) = LCase(szName) Then
      GetVarDb = VariableDb(ii)
      Exit Function
    End If
  Next
  
  'if not exists Create a dummy variable string
  DummyVarDb.Name = szName
  DummyVarDb.Type = TVDB_STRING
  GetVarDb = DummyVarDb

  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basVarDb.GetVarDb"
End Function

'return the image name of value
Public Function GetImageFromVal(szValue As String, TypeVar As TypeVarDb) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basVarDb.GetImageFromValCast(" & szValue & "," & TypeVar & ")", etFullDebug

Dim szTemp As String
Dim szImg As String
Dim vData
Dim vDataTemp

  szImg = "property"      'image default
  
  If TypeVar = TVDB_CAST Then
    vData = Array("info", "error", "warning", "debug", "log")
    For Each vDataTemp In vData
      If LCase(Left(szValue, Len(vDataTemp))) = vDataTemp Then
        szImg = vDataTemp
        Exit For
      End If
    Next
  ElseIf TypeVar = TVDB_BOOLEAN Then
    Select Case UCase(szValue)
      Case "ON", "TRUE", "YES", "1"
        szImg = "on"
      Case "OFF", "FALSE", "NO", "0"
        szImg = "off"
    End Select
  End If
  
  GetImageFromVal = szImg
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basVarDb.GetImageFromValCast"
End Function
