VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'pgAdmin II application
Public pgApp As Object

'Key Word Reserved for escape name object
Public KeyWordReserved As New Collection
Attribute KeyWordReserved.VB_VarDescription = "Key Word Reserved for escape name"

'Error Handler Stuff
Dim lErrorIndex As Long
Dim pgErrors As New Collection

'Events
Public Event EventLog(EventLevel As LogLevel, EventMessage As String)
Attribute EventLog.VB_Description = "Executes whenever a Log event occurs."
Public Event SQLExecute(SQL As String, Connection As Connection)
Attribute SQLExecute.VB_Description = "Execute an SQL query."

'Types
Public Type Version
  Description As String
  Major As Integer
  Minor As Integer
  Revision As Integer
  VersionNum As Single
End Type

'Enums
Public Enum LogLevel
  llNone = 0
  llErrors = 1
  llsql = 2
  llMiniDebug = 3
  llFullDebug = 4
End Enum

Public Enum ErrorType
  etErrors = 1
  etSQL = 2
  etMiniDebug = 3
  etFullDebug = 4
End Enum

Public Enum aclPrivileges
  aclAll = 1
  aclSelect = 2
  aclInsert = 4
  aclUpdate = 8
  aclDelete = 16
  aclRule = 32
  aclReferences = 64
  aclTrigger = 128
  aclExecute = 256
  aclCreate = 512
  aclTemp = 1024
  aclUsage = 2048
End Enum

Public Enum aclType
  aclClass = 0
  aclDatabase = 1
  aclLanguage = 2
  aclFunction = 3
  aclSchema = 4
End Enum

Public Enum dbStatus
  statClosed = 0
  statOpen = 1
  statInaccessible = 2
End Enum

Public Enum QueryType
  qrySystem = 1
  qryData = 2
  qryUser = 4
End Enum

'Collections
Private colDatabases As Databases
Private colUsers As Users
Private colGroups As Groups

'Properties
Private cnPostgreSQL As New Connection
Private lLogging As LogLevel
Private szLogfile As String
Private szServer As String
Private lPort As Long
Private szUsername As String
Private szPassword As String
Private lLastSystemOID As Double
Private bIncludeSys As Boolean
Private bDontLogErrors As Boolean
Private bShowPassword As Boolean
Private szMasterDB As String
Private bEncryptPasswords As Boolean
Private objDriverVersion As Version
Private lUserLogOptions As Long
Private szUserLogfile As String
Private bUserLog As Boolean
Private bDeferConnection As Boolean
Private vTag As Variant

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Error Log
'If an application is logging it's errors through the mechanism here, then any errors
'generated internally will be logged internally AND logged by the host application.
'To fix this, internally we use iLogEvent, and the hosting application can set DontLogErrors
'as required.
Friend Sub iLogError(ByVal lError As Long, ByVal szError As String)
'Don't include error handling here *of any kind* 'cos you'll break it!

  If Not bDontLogErrors Then
    LogEvent "Error: " & lError & " - " & szError, etErrors
  End If
  Err.Raise lError, App.Title, szError
End Sub
Friend Sub iLogEvent(szMessage As String, etType As ErrorType)

  If Not ((bDontLogErrors) And (etType = etErrors)) Then
  
    'If this is an error, add it to the collection
    If etType = etErrors Then
      pgErrors.Add szMessage
    End If
    
    LogEvent szMessage, etType
  End If
End Sub
Public Sub LogEvent(ByVal szMessage As String, ByVal etType As ErrorType)
Attribute LogEvent.VB_Description = "Log an event to the libraries logging system."
Dim fNum As Integer
Dim vError As Variant
Static bHaveWarned As Boolean

  RaiseEvent EventLog(etType, szMessage)
  If lLogging >= etType Then
    fNum = FreeFile
    
    On Error GoTo Err_Handler
    Open szLogfile For Append As #fNum
    
    'If an error has occured, make it a bit more obvious and log extra info as well
    If etType = etErrors Then
      Print #fNum, ""
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - *******************************************************************"
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Error(s)"
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - *******************************************************************"
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - " & szMessage
      If pgErrors.Count > 0 Then
        Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - The following error(s) have been logged:"
        Print #fNum, ""
        For Each vError In pgErrors
          Print #fNum, "  " & vError
        Next vError
        Set pgErrors = New Collection
      End If
      Print #fNum, ""
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Windows Version: " & WinName & " v" & WinVer & " build " & WinBuild & " " & WinInfo
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - pgSchema Version: " & App.Major & "." & App.Minor & "." & App.Revision
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - MDAC Version: " & cnPostgreSQL.Version
      If cnPostgreSQL.State = adStateOpen Then
        Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - DBMS Version: " & cnPostgreSQL.Properties("DBMS Version")
        If bShowPassword Then
          Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Connection String (Master Connection)" & ": " & cnPostgreSQL.ConnectionString
        Else
          Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - Connection String (Master Connection)" & ": " & Replace(cnPostgreSQL.ConnectionString, ";PWD=" & szPassword, ";PWD=********")
        End If
      Else
        Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - cnPostgreSQL.State = adStateClosed"
      End If
      Print #fNum, ""
    Else
      Print #fNum, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - " & szMessage
    End If
    Close #fNum
    
  End If
  Exit Sub
  
Err_Handler:
  If Err.Number = 76 Then
    If bHaveWarned = False Then
      MsgBox "The logfile could not be opened. Please check the logging settings.", vbExclamation, "Error"
      bHaveWarned = True
    End If
  Else
    Resume Next
  End If
End Sub

Friend Function ExecSQL(SQL As String, Connection As Connection, Optional RecordsAffected As Long, Optional Options As Long = -1, Optional QryType As QueryType = qrySystem) As Recordset
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.ExecSQL(" & QUOTE & SQL & QUOTE & ", " & QUOTE & "Connection: " & Connection.ConnectionString & QUOTE & ", " & RecordsAffected & ", " & Options & ", " & QryType & ")", etFullDebug
  
Dim fNum As Integer
Dim objError As Error
Dim szQueryType As String

  RaiseEvent SQLExecute(SQL, Connection)
  objServer.iLogEvent "SQL (" & GetDatabase(Connection.ConnectionString) & "): " & SQL, etSQL
  Set ExecSQL = Connection.Execute(SQL)
  
  'Userlog
  If (bUserLog And ((lUserLogOptions And QryType) = QryType)) Then
  
    'Get the Query Type
    If (QryType And qrySystem) = qrySystem Then
      szQueryType = "System"
    ElseIf (QryType And qryData) = qryData Then
      szQueryType = "Data"
    ElseIf (QryType And qryUser) = qryUser Then
      szQueryType = "User"
    End If
    
    fNum = FreeFile
    Open szUserLogfile For Append As #fNum
    
    Print #fNum, "-- "
    Print #fNum, "-- Date/Time : " & Format(Now, "yyyy-MM-dd hh:mm:ss")
    Print #fNum, "-- Connection: " & szUsername & "@" & szServer & ":" & lPort
    Print #fNum, "-- Database  : " & GetDatabase(Connection.ConnectionString)
    Print #fNum, "-- Query Type: " & szQueryType & vbCrLf
    Print #fNum, SQL & vbCrLf
    Print #fNum, "-- End of log entry" & vbCrLf
    
    Close #fNum
  End If
  
  Exit Function
Err_Handler:
  'If this is an error, add it to the collection
  For Each objError In Connection.Errors
    pgErrors.Add objError.Number & " - " & objError.Description
  Next objError
  objServer.iLogError Err.Number, Err.Description
End Function

'Refresh the collections
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collections."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Refresh()", etFullDebug

  Set colDatabases = Nothing
  Set colGroups = Nothing
  Set colUsers = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Connect to the database
Public Function Connect(Optional ByVal Server As String, Optional ByVal Port As Long, Optional ByVal Username As String, Optional ByVal Password As String) As Long
Attribute Connect.VB_Description = "Open a connection to the specified database."
Attribute Connect.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Connect(" & QUOTE & Server & QUOTE & ", " & lPort & ", " & QUOTE & Username & QUOTE & ", " & QUOTE & "********" & QUOTE & ")", etFullDebug

Dim szConnect As String
Dim rs As New Recordset
Dim X As Integer
Dim szVersion() As String

  'Close the connection if already open
  If cnPostgreSQL.State <> adStateClosed Then
    objServer.iLogEvent "Closing cnPostgreSQL", etSQL
    cnPostgreSQL.Close
  End If
  
  'Override the current values with those passed to us
  If Server <> "" Then szServer = Server
  If Port <> 0 Then lPort = Port
  If Username <> "" Then szUsername = Username
  If Password <> "" Then szPassword = Password
  
  'Default required values
  If lPort = 0 Then lPort = 5432
  If szServer = "" Then szServer = "127.0.0.1"
  
  'Clear collections
  Set colDatabases = Nothing
  Set colUsers = Nothing
  Set colGroups = Nothing
  
  'Connection String
  szConnect = "SERVER=" & szServer & ";PORT=" & lPort & ";UID=" & szUsername & ";PWD=" & szPassword & ";DATABASE=" & szMasterDB & ";" & ODBC_CONNECT_OPTIONS
  
  'Set the ODBC logging
  Select Case lLogging
    Case llMiniDebug
      szConnect = szConnect & "COMMLOG=1;DEBUG=0;"
    Case llFullDebug
      szConnect = szConnect & "COMMLOG=1;DEBUG=1;"
    Case Else
      szConnect = szConnect & "COMMLOG=0;DEBUG=0;"
  End Select
  
  'Open the connection
  If bShowPassword Then
    objServer.iLogEvent "Opening cnPostgreSQL: " & szConnect, etMiniDebug
  Else
    objServer.iLogEvent "Opening cnPostgreSQL: " & Replace(szConnect, ";PWD=" & szPassword, ";PWD=********"), etMiniDebug
  End If
  cnPostgreSQL.Open szConnect
  
  'Get the Server Version
  Set rs = objServer.ExecSQL("SELECT version()", cnPostgreSQL)
  If Not rs.EOF Then objVersion.Description = rs!Version & ""
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  'Split the version down.
  X = InStr(1, objVersion.Description, " ")
  szVersion = Split(Mid(objVersion.Description, X + 1, InStr(X + 1, objVersion.Description, " ") - X), ".")
  If UBound(szVersion) >= 0 Then objVersion.Major = Val(szVersion(0))
  If UBound(szVersion) >= 1 Then objVersion.Minor = Val(szVersion(1))
  If UBound(szVersion) >= 2 Then objVersion.Revision = Val(szVersion(2))
  objVersion.VersionNum = Val(objVersion.Major & "." & objVersion.Minor)
  
  'Check the version
  If objVersion.VersionNum < 7.1 Then
    objServer.iLogEvent "Closing cnPostgreSQL.", etMiniDebug
    If cnPostgreSQL.State <> adStateClosed Then cnPostgreSQL.Close
    Set cnPostgreSQL = Nothing
    objVersion.Description = ""
    objVersion.Major = 0
    objVersion.Minor = 0
    objVersion.Revision = 0
    objVersion.VersionNum = 0
    Err.Raise vbObjectError + 1, App.Title, "This program can only be used with PostgreSQL 7.1 or higher."
    Exit Function
  End If
  
  'Get the last system OID!
  Set rs = objServer.ExecSQL("SELECT datlastsysoid FROM pg_database WHERE datname = 'template1'", cnPostgreSQL)
  If Not rs.EOF Then lLastSystemOID = rs!datlastsysoid & ""
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  'Get the Driver Version
  szVersion = Split(cnPostgreSQL.Properties("Driver Version"), ".")
  If UBound(szVersion) >= 0 Then objDriverVersion.Major = Val(szVersion(0))
  If UBound(szVersion) >= 1 Then objDriverVersion.Minor = Val(szVersion(1))
  If UBound(szVersion) >= 2 Then objDriverVersion.Revision = Val(szVersion(2))
  objDriverVersion.VersionNum = Val(objDriverVersion.Major & "." & objDriverVersion.Minor)
  objDriverVersion.Description = Mid(cnPostgreSQL.Properties("DBMS Version"), 12)
  
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Function

'Return the Databases Collection
Public Property Get Databases() As Databases
Attribute Databases.VB_Description = "Returns a collection of Database objects on the current Server."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Databases()", etFullDebug

  If colDatabases Is Nothing Then
    Set colDatabases = New Databases
    colDatabases.Initialize
  End If
  Set Databases = colDatabases
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Users Collection
Public Property Get Users() As Users
Attribute Users.VB_Description = "Returns a collection of User objects on the current Server."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Users()", etFullDebug

  If colUsers Is Nothing Then
    Set colUsers = New Users
    colUsers.Initialize
  End If
  Set Users = colUsers
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Groups Collection
Public Property Get Groups() As Groups
Attribute Groups.VB_Description = "Returns a collection of User Group objects on the current Server."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Groups()", etFullDebug

  If colGroups Is Nothing Then
    Set colGroups = New Groups
    colGroups.Initialize
  End If
  Set Groups = colGroups
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Setup default settings
Private Sub Class_Initialize()
'Can't do error handling properly yet
Dim szFilename As String
Dim lCount As Integer

  'Where are we running?
  szFilename = String(255, 0)
  lCount = GetModuleFileName(App.hInstance, szFilename, 255)
  szFilename = Left(szFilename, lCount)
  If UCase(Right(szFilename, 7)) = "VB6.EXE" Then
    inIDE = True
  Else
    inIDE = False
  End If
  
  Set objServer = Me
  szServer = "127.0.0.1"
  lPort = 5432
  lLogging = llNone
  szLogfile = "pgSchema_" & App.ThreadID & ".log"
  szUserLogfile = "pgSchema_User_" & App.ThreadID & ".log"
  szMasterDB = "template1"
End Sub

'Cleanup
Private Sub Class_Terminate()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
  objServer.iLogEvent "Entering " & App.Title & ":pgServer.Class_Terminate()", etFullDebug
  
  Set colDatabases = Nothing
  Set colUsers = Nothing
  Set colGroups = Nothing
  If cnPostgreSQL.State <> adStateClosed Then
    objServer.iLogEvent "Closing cnPostgreSQL", etMiniDebug
    cnPostgreSQL.Close
  End If
  Set cnPostgreSQL = Nothing
    
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get ObjectType()", etFullDebug

  ObjectType = "Server"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Identifier
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Identifier()", etFullDebug

  Identifier = szUsername & "@" & szServer & ":" & lPort
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the ADO Connection String
Friend Property Get iConnectionString() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get iConnectionString()", etFullDebug

  If cnPostgreSQL.State <> adStateOpen Then
    iConnectionString = ""
  Else
    iConnectionString = cnPostgreSQL.ConnectionString
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_Description = "Returns the ADO connection string currently in use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get ConnectionString()", etFullDebug

  If cnPostgreSQL.State <> adStateOpen Then
    ConnectionString = ""
  Else
    If bShowPassword Then
      ConnectionString = cnPostgreSQL.ConnectionString
    Else
      ConnectionString = Replace(cnPostgreSQL.ConnectionString, ";PWD=" & szPassword, ";PWD=********")
    End If
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the ODBC Driver Name
Public Property Get DriverName() As String
Attribute DriverName.VB_Description = "Returns the ODBC driver name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get DriverName()", etFullDebug

  If cnPostgreSQL.State <> adStateOpen Then
    DriverName = ""
  Else
    DriverName = cnPostgreSQL.Properties("DBMS Name")
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Last System OID
Public Property Get LastSystemOID() As Double
Attribute LastSystemOID.VB_Description = "Returns the last OID used by the Database System."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get LastSystemOID()", etFullDebug

  LastSystemOID = lLastSystemOID
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Driver Version
Public Property Get DriverVersion() As Version
Attribute DriverVersion.VB_Description = "Returns the ODBC driver version."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get DriverVersion()", etFullDebug

  DriverVersion = objDriverVersion
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the PostgreSQL Version
Public Property Get dbVersion() As Version
Attribute dbVersion.VB_Description = "Returns a Version structure indicating the Version of the PostgreSQL Server."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get dbVersion()", etFullDebug

  dbVersion = objVersion
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return Our Version
Public Property Get Version() As Version
Attribute Version.VB_Description = "Returns a Version structure indicating the Version of the object library."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Version()", etFullDebug

  Version.Description = App.Comments & " v" & App.Major & "." & App.Minor & "." & App.Revision
  Version.Major = App.Major
  Version.Minor = App.Minor
  Version.Revision = App.Revision
  Version.VersionNum = Val(App.Major & "." & App.Minor)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return IncludeSys
Public Property Let IncludeSys(ByVal vData As Boolean)
Attribute IncludeSys.VB_Description = "Sets/Returns a flag indicating whether or not to include System Objects in the Object Hierarchy."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let IncludeSys(" & vData & ")", etFullDebug

  bIncludeSys = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get IncludeSys() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get IncludeSys()", etFullDebug

  IncludeSys = bIncludeSys
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return DontLogErrors
Public Property Let DontLogErrors(ByVal vData As Boolean)
Attribute DontLogErrors.VB_Description = "If an application is logging it's errors using LogEvent, then any errors generated internally will be logged internally AND logged by the host application. This property (when True) prevents errors being logged internally."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let DontLogErrors(" & vData & ")", etFullDebug

  bDontLogErrors = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get DontLogErrors() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get DontLogErrors()", etFullDebug

  DontLogErrors = bDontLogErrors
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return EncryptPasswords
Public Property Let EncryptPasswords(ByVal vData As Boolean)
Attribute EncryptPasswords.VB_Description = "Should passwords be encrypted in PostgreSQL 7.2+?"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let EncryptPasswords(" & vData & ")", etFullDebug

  bEncryptPasswords = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get EncryptPasswords() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get EncryptPasswords()", etFullDebug

  EncryptPasswords = bEncryptPasswords
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return ShowPassword
Public Property Let ShowPassword(ByVal vData As Boolean)
Attribute ShowPassword.VB_Description = "Sets/Returns whether passwords should be included in logfiles."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let ShowPassword(" & vData & ")", etFullDebug

  bShowPassword = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ShowPassword() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get ShowPassword()", etFullDebug

  ShowPassword = bShowPassword
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Port
Public Property Let Port(ByVal vData As Long)
Attribute Port.VB_Description = "Returns the port number that the PostgreSQL server is listening on."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Port(" & vData & ")", etFullDebug

  lPort = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Port() As Long
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Port()", etFullDebug

  Port = lPort
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Public Property Let Server(ByVal vData As String)
Attribute Server.VB_Description = "Sets/Returns the PostgreSQL Server Hostname or IP address."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Server(" & QUOTE & vData & QUOTE & ")", etFullDebug

  Server = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Server() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Server()", etFullDebug

  Server = szServer
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Password
Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "Sets/Returns the Password used for the connection."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Password(" & QUOTE & "********" & QUOTE & ")", etFullDebug

  Password = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Password() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Password()", etFullDebug

  If bShowPassword Then
    Password = szPassword
  Else
    Password = "********"
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Get iPassword() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get iPassword()", etFullDebug

  iPassword = szPassword
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Username
Public Property Let Username(ByVal vData As String)
Attribute Username.VB_Description = "Sets/Returns the Username used for the connection."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Username(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szUsername = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Username() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Username()", etFullDebug

  Username = szUsername
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Logfile
Public Property Let Logfile(ByVal vData As String)
Attribute Logfile.VB_Description = "Sets/Returns the name of the logfile to use. Include %ID in the filename which will be replaced with a process ID (hInstance) if required."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let Logfile(" & QUOTE & vData & QUOTE & ")", etFullDebug

  If vData = "" Then
    szLogfile = "pgSchema_" & App.ThreadID & ".log"
  Else
    szLogfile = Replace(vData, "%ID", App.ThreadID)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Logfile() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Logfile()", etFullDebug

  If szLogfile = "" Then szLogfile = "pgSchema_" & App.ThreadID & ".log"
  Logfile = szLogfile
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the UserLogOptions
Public Property Let UserLogOptions(ByVal vData As Long)
Attribute UserLogOptions.VB_Description = "Sets/Returns a bitmask describing the query types to include in the user log."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let UserLogOptions(" & vData & ")", etFullDebug

  lUserLogOptions = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get UserLogOptions() As Long
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get UserLogOptions()", etFullDebug

  UserLogOptions = lUserLogOptions
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the UserLogfile
Public Property Let UserLogfile(ByVal vData As String)
Attribute UserLogfile.VB_Description = "Sets/Returns the filename for the user log."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let UserLogfile(" & QUOTE & vData & QUOTE & ")", etFullDebug

  If vData = "" Then
    szUserLogfile = "pgSchema_User_" & App.ThreadID & ".log"
  Else
    szUserLogfile = Replace(vData, "%ID", App.ThreadID)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get UserLogfile() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get UserLogfile()", etFullDebug

  If szUserLogfile = "" Then szUserLogfile = "pgSchema_" & App.ThreadID & ".log"
  UserLogfile = szUserLogfile
      
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the UserLog
Public Property Let UserLog(ByVal vData As Boolean)
Attribute UserLog.VB_Description = "Sets/Returns a value determining whether or not the user log is enabled."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let UserLog(" & QUOTE & vData & QUOTE & ")", etFullDebug

  bUserLog = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get UserLog() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get UserLog()", etFullDebug

  UserLog = bUserLog
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the DeferConnection
Public Property Let DeferConnection(ByVal vData As Boolean)
Attribute DeferConnection.VB_Description = "Sets/Returns whether to defer connecting to databases until necessary. Setting this will prevent database comment & status retrieval initiating a connection."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let DeferConnection(" & QUOTE & vData & QUOTE & ")", etFullDebug

  bDeferConnection = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get DeferConnection() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get DeferConnection()", etFullDebug

  DeferConnection = bDeferConnection
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the MasterDB
Public Property Let MasterDB(ByVal vData As String)
Attribute MasterDB.VB_Description = "Sets/Returns the name of the master connection database."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let MasterDB(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szMasterDB = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get MasterDB() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get MasterDB()", etFullDebug

  MasterDB = szMasterDB
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Logging Property
Public Property Let LogLevel(ByVal vData As LogLevel)
Attribute LogLevel.VB_Description = "Sets/Returns the level of detail to include in the logfile."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Let LogLevel(" & QUOTE & vData & QUOTE & ")", etFullDebug

  lLogging = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get LogLevel() As LogLevel
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get LogLevel()", etFullDebug

  LogLevel = lLogging
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Friend Property Get Connection() As Connection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgServer.Property Get Connection()", etFullDebug

  Set Connection = cnPostgreSQL
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
