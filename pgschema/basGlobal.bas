Attribute VB_Name = "basGlobal"
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Global Variables
Public objVersion As Version
Public objServer As pgServer

'Constants
Public Const QUOTE = """"
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const ODBC_CONNECT_OPTIONS = "DRIVER={PostgreSQL};READONLY=0;PROTOCOL=6.4;FAKEOIDINDEX=0;SHOWOIDCOLUMN=0;ROWVERSIONING=0;SHOWSYSTEMTABLES=0;CONNSETTINGS=;FETCH=100;SOCKET=4096;UNKNOWNSIZES=0;MAXVARCHARSIZE=254;MAXLONGVARCHARSIZE=65536;OPTIMIZER=1;KSQO=1;USEDECLAREFETCH=0;TEXTASLONGVARCHAR=1;UNKNOWNSASLONGVARCHAR=1;BOOLSASCHAR=1;PARSE=0;CANCELASFREESTMT=0;EXTRASYSTABLEPREFIXES=dd_;"

'SQL constants
Public Const SQL_GET_DATABASES = "SELECT oid, *, pg_encoding_to_char(encoding) AS serverencodingname, pg_get_userbyid(datdba) AS datowner FROM pg_database"
Public Const SQL_GET_LANGUAGES = "SELECT now() AS ts, oid, * FROM pg_language"
Public Const SQL_GET_USERS = "SELECT * FROM pg_user"
Public Const SQL_GET_GROUPS = "SELECT * FROM pg_group"
Public Const SQL_GET_SEQUENCES = "SELECT now() AS ts, oid, relname, pg_get_userbyid(relowner) AS seqowner, relacl FROM pg_class WHERE relkind = 'S'"
Public Const SQL_GET_VIEWS = "SELECT now() AS ts, c.oid, c.relname, pg_get_userbyid(c.relowner) AS viewowner, c.relacl, pg_get_viewdef(c.relname) AS definition FROM pg_class c WHERE ((c.relhasrules AND (EXISTS (SELECT r.rulename FROM pg_rewrite r WHERE ((r.ev_class = c.oid) AND (bpchar(r.ev_type) = '1'::bpchar))))) OR (c.relkind = 'v'::" & QUOTE & "char" & QUOTE & "))"
Public Const SQL_GET_TYPES = "SELECT now() AS ts, oid, *, pg_get_userbyid(typowner) as typeowner FROM pg_type WHERE typrelid = 0"
Public Const SQL_GET_FUNCTIONS = "SELECT now() AS ts, oid, *, pg_get_userbyid(proowner) as funcowner FROM pg_proc"
Public Const SQL_GET_OPERATORS = "SELECT now() AS ts, oid, *, pg_get_userbyid(oprowner) as opowner FROM pg_operator"
Public Const SQL_GET_RULES = "SELECT now() AS ts, oid, rulename, pg_get_ruledef(rulename) as definition FROM pg_rewrite"
Public Const SQL_GET_TRIGGERS = "SELECT now() AS ts, t.oid, tgname, proname, tgargs, tgtype FROM pg_trigger t, pg_proc p WHERE t.tgfoid = p.oid AND tgisconstraint = FALSE"
Public Const SQL_GET_TABLES7_1 = "SELECT now() AS ts, oid, relname, pg_get_userbyid(relowner) as tableowner, relacl FROM pg_class WHERE ((relkind = 'r') OR (relkind = 's'))"
Public Const SQL_GET_TABLES7_2 = "SELECT now() AS ts, oid, relname, pg_get_userbyid(relowner) as tableowner, relacl, relhasoids FROM pg_class WHERE ((relkind = 'r') OR (relkind = 's'))"
Public Const SQL_GET_COLUMNS7_1 = "SELECT a.oid, a.attname, a.attnum, t.typname, CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS length, a.attnotnull, (SELECT adsrc FROM pg_attrdef d WHERE d.adrelid = a.attrelid AND d.adnum = a.attnum) AS default, (SELECT indisprimary FROM pg_index i, pg_class ic, pg_attribute ia  WHERE i.indrelid = a.attrelid AND i.indexrelid = ic.oid AND ic.oid = ia.attrelid AND ia.attname = a.attname LIMIT 1) AS primarykey FROM pg_attribute a, pg_type t WHERE a.atttypid = t.oid"
Public Const SQL_GET_COLUMNS7_2 = "SELECT 0::oid AS oid, a.attname, a.attnum, t.typname, CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS length, a.attnotnull, (SELECT adsrc FROM pg_attrdef d WHERE d.adrelid = a.attrelid AND d.adnum = a.attnum) AS default, (SELECT indisprimary FROM pg_index i, pg_class ic, pg_attribute ia  WHERE i.indrelid = a.attrelid AND i.indexrelid = ic.oid AND ic.oid = ia.attrelid AND ia.attname = a.attname LIMIT 1) AS primarykey FROM pg_attribute a, pg_type t WHERE a.atttypid = t.oid"
Public Const SQL_GET_INDEXES = "SELECT now() AS ts, i.oid, i.relname, x.indisunique, x.indisprimary, pg_get_indexdef(i.oid) AS definition FROM pg_index x, pg_class i WHERE i.oid = x.indexrelid"
Public Const SQL_GET_INDEX_COLUMNS = "SELECT attname FROM pg_attribute"
Public Const SQL_GET_CHECKS = "SELECT rcname, rcsrc FROM pg_relcheck WHERE NOT EXISTS (SELECT * FROM pg_relcheck AS c, pg_inherits AS i WHERE i.inhrelid = pg_relcheck.rcrelid AND (c.rcname = pg_relcheck.rcname OR (c.rcname[0] = '$' AND pg_relcheck.rcname[0] = '$')) AND c.rcsrc = pg_relcheck.rcsrc AND  c.rcrelid = i.inhparent)"
Public Const SQL_GET_INHERITED_TABLES = "SELECT c.relname FROM pg_class c, pg_inherits i WHERE c.oid = i.inhparent"
Public Const SQL_GET_AGGREGATES = "SELECT now() AS ts, oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate"
Public Const SQL_GET_FOREIGN_KEYS = "SELECT oid, tgrelid, tgconstrname, tgnargs, tgargs, tgdeferrable, tginitdeferred FROM pg_trigger WHERE tgisconstraint = TRUE AND tgtype = 21"

'SQL related to Revision Control.
'Note that the object OID is also logged to help create objects in dependency order when building scripts.
Public Const SQL_CREATE_REVLOG = "CREATE TABLE pgadmin_rclog(rc_timestamp timestamp DEFAULT now(), rc_user name DEFAULT current_user, rc_action varchar(1), rc_type varchar(32), rc_identifier varchar(256), rc_oid oid, rc_table varchar(64), rc_version int4, rc_definition text, rc_comment text); GRANT SELECT, INSERT ON pgadmin_rclog TO PUBLIC; COMMENT ON TABLE pgadmin_rclog IS 'pgAdmin II Revision Log'; CREATE INDEX pgadmin_rclog_idx ON pgadmin_rclog (rc_action, rc_type, rc_identifier, rc_table, rc_oid, rc_version);"
Public Const SQL_DROP_REVLOG = "DROP TABLE pgadmin_rclog;"
Public Const SQL_GRAVEYARD = "SELECT DISTINCT ON (rc_type, rc_identifier) * FROM pgadmin_rclog ORDER BY rc_type, rc_identifier, rc_version DESC"

'Type Declarations
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' NOTE: Some functions/subs here don't have logging or error handling
' for speed etc.

'Convert Strings for SQL
Public Function dbSZ(ByVal szData As String) As String
On Error Resume Next

  szData = Replace(szData, "\", "\\")
  szData = Replace(szData, "'", "''")
  szData = Replace(szData, vbCrLf, "\n")
  dbSZ = szData

End Function

'Convert Boolean field values to a Boolean
Public Function ToBool(ByVal vData As Variant) As Boolean
On Error Resume Next

  Select Case UCase(vData)
    Case "T"
      ToBool = True
    Case "F"
      ToBool = False
    Case "TRUE"
      ToBool = True
    Case "FALSE"
      ToBool = False
    Case 1
      ToBool = True
    Case 0
      ToBool = False
  End Select

End Function

'Encode case sensitive string into a non case sensitive string
Public Function ULEncode(ByVal szData As String) As String
On Error Resume Next

Dim X As Integer
Dim szChar As String
Dim szOutput As String

  For X = 1 To Len(szData)
    szChar = Mid(szData, X, 1)
    If (Asc(szChar) >= 65) And (Asc(szChar) <= 90) Then
      szOutput = szOutput & "U" & szChar
    Else
      szOutput = szOutput & "L" & szChar
    End If
  Next X

  ULEncode = szOutput

End Function

'Get a unique ID
Public Function GetUniqueID() As Long
On Error Resume Next

Static LastUniqueID As Long

  LastUniqueID = LastUniqueID + 1
  GetUniqueID = LastUniqueID
  
End Function

'Return the Database from a connection string
Public Function GetDatabase(ByVal szConnectionString As String) As String
On Error Resume Next

Dim X As Integer

  X = InStr(1, UCase(szConnectionString), "DATABASE=")
  If X <> 0 Then GetDatabase = Mid(szConnectionString, X + 9, InStr(X + 9, szConnectionString, ";") - (X + 9))

End Function

'Parse an ACL and return GRANT/REVOKE Statements
Public Function ParseACL(ByVal szObject As String, ByVal szACL As String) As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ParseACL(" & QUOTE & szObject & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim szEntries() As String
Dim szEntry As Variant
Dim szName As String
Dim szAccess As String
Dim szSQL As String
Dim szTemp As String
  
  szACL = Mid(szACL, 2, Len(szACL) - 2)
  szACL = Replace(szACL, QUOTE, "")
  szEntries = Split(szACL, ",")
  For Each szEntry In szEntries
  
    'Get the user/group name
    If UCase(Left(szEntry, 6)) = "GROUP " Then
      szName = "GROUP " & QUOTE & Mid(szEntry, 7, InStr(1, szEntry, "=") - 7) & QUOTE
    Else
      szName = QUOTE & Left(szEntry, InStr(1, szEntry, "=") - 1) & QUOTE
    End If
    If szName = QUOTE & QUOTE Then szName = "PUBLIC"
    
    'Get the Access
    szAccess = Mid(szEntry, InStr(1, szEntry, "=") + 1)
    
    'If the Access is "" then REVOKE all
    If szAccess = "" Then
      szSQL = szSQL & "REVOKE ALL ON " & QUOTE & szObject & QUOTE & " FROM " & szName & ";" & vbCrLf
    Else
    
      'Either grant ALL or individual privileges
      'Note that in 7.2+, Delete has been seperated from Update, and References/Trigger
      'have been added.
      If objServer.dbVersion.VersionNum >= 7.2 Then
        If szAccess = "arwdRxt" Then
          szAccess = "ALL"
        Else
          szTemp = ""
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "INSERT, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "SELECT, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "UPDATE, "
          If InStr(1, szAccess, "d") <> 0 Then szTemp = szTemp & "DELETE, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "RULE, "
          If InStr(1, szAccess, "x") <> 0 Then szTemp = szTemp & "REFERENCES, "
          If InStr(1, szAccess, "t") <> 0 Then szTemp = szTemp & "TRIGGER, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
        End If
      Else
        If szAccess = "arwR" Then
          szAccess = "ALL"
        Else
          szTemp = ""
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "INSERT, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "SELECT, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "UPDATE, DELETE, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "RULE, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
        End If
      End If
      
      szSQL = szSQL & "GRANT " & szAccess & " ON " & QUOTE & szObject & QUOTE & " TO " & szName & ";" & vbCrLf
    End If
  Next szEntry
  
  ParseACL = szSQL
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

Public Function WinVer() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!

Dim osVersion As OSVERSIONINFO

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  WinVer = osVersion.dwMajorVersion & "." & osVersion.dwMinorVersion

End Function

Public Function WinBuild() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!

Dim osVersion As OSVERSIONINFO

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  WinBuild = osVersion.dwBuildNumber And &HFFFF&
  
End Function

Public Function WinName() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!

Dim osVersion As OSVERSIONINFO

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  Select Case osVersion.dwPlatformId
    Case VER_PLATFORM_WIN32s
      WinName = "Windows 3.x"
    Case VER_PLATFORM_WIN32_WINDOWS
      If osVersion.dwMinorVersion = 0 Then WinName = "Windows 95"
      If osVersion.dwMinorVersion = 10 Then WinName = "Windows 98"
      If osVersion.dwMinorVersion = 90 Then WinName = "Windows ME"
    Case VER_PLATFORM_WIN32_NT
      If osVersion.dwMajorVersion < 5 Then WinName = "Windows NT"
      If osVersion.dwMajorVersion = 5 Then
        If osVersion.dwMinorVersion = 0 Then WinName = "Windows 2000"
        If osVersion.dwMinorVersion = 1 Then WinName = "Windows XP"
      End If
  End Select

End Function

Public Function WinInfo() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!bug

Dim osVersion As OSVERSIONINFO
Dim iLoc As Integer

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  iLoc = InStr(1, osVersion.szCSDVersion, Chr(0))
  
  If Len(osVersion.szCSDVersion) > 0 Then
    WinInfo = LTrim(Left(osVersion.szCSDVersion, iLoc - 1))
  Else
    WinInfo = vbNullString
  End If

End Function
