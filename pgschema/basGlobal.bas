Attribute VB_Name = "basGlobal"
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Global Variables
Public objVersion As Version
Public objServer As pgServer
Public inIDE As Boolean

'Constants
Public Const QUOTE = """"
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const ODBC_CONNECT_OPTIONS = "DRIVER={PostgreSQL};READONLY=0;PROTOCOL=6.4;FAKEOIDINDEX=0;SHOWOIDCOLUMN=0;ROWVERSIONING=0;SHOWSYSTEMTABLES=0;CONNSETTINGS=;FETCH=100;SOCKET=4096;UNKNOWNSIZES=0;MAXVARCHARSIZE=254;MAXLONGVARCHARSIZE=65536;OPTIMIZER=1;KSQO=1;USEDECLAREFETCH=0;TEXTASLONGVARCHAR=1;UNKNOWNSASLONGVARCHAR=0;BOOLSASCHAR=1;PARSE=0;CANCELASFREESTMT=0;EXTRASYSTABLEPREFIXES=dd_;"

'SQL constants
Public Const SQL_GET_DATABASES7_1 = "SELECT oid, datname, datpath, datallowconn, pg_encoding_to_char(encoding) AS serverencoding, pg_get_userbyid(datdba) AS datowner FROM pg_database"
Public Const SQL_GET_DATABASES7_3 = "SELECT oid, datname, datpath, datallowconn, datconfig, datacl, pg_encoding_to_char(encoding) AS serverencoding, pg_get_userbyid(datdba) AS datowner FROM pg_database"
Public Const SQL_GET_LANGUAGES = "SELECT oid, * FROM pg_language"
Public Const SQL_GET_USERS = "SELECT * FROM pg_user"
Public Const SQL_GET_GROUPS = "SELECT * FROM pg_group"
Public Const SQL_GET_SEQUENCES = "SELECT oid, relname, pg_get_userbyid(relowner) AS seqowner, relacl FROM pg_class WHERE relkind = 'S'"
Public Const SQL_GET_VIEWS7_1 = "SELECT c.oid, c.relname, pg_get_userbyid(c.relowner) AS viewowner, c.relacl, pg_get_viewdef(c.relname) AS definition FROM pg_class c WHERE ((c.relhasrules AND (EXISTS (SELECT r.rulename FROM pg_rewrite r WHERE ((r.ev_class = c.oid) AND (bpchar(r.ev_type) = '1'::bpchar))))) OR (c.relkind = 'v'::" & QUOTE & "char" & QUOTE & "))"
Public Const SQL_GET_VIEWS7_3 = "SELECT c.oid, c.relname, pg_get_userbyid(c.relowner) AS viewowner, c.relacl, pg_get_viewdef(c.oid) AS definition FROM pg_class c WHERE ((c.relhasrules AND (EXISTS (SELECT r.rulename FROM pg_rewrite r WHERE ((r.ev_class = c.oid) AND (bpchar(r.ev_type) = '1'::bpchar))))) OR (c.relkind = 'v'::" & QUOTE & "char" & QUOTE & "))"
Public Const SQL_GET_TYPES7_1 = "SELECT oid, *, pg_get_userbyid(typowner) as typeowner FROM pg_type WHERE typrelid = 0"
Public Const SQL_GET_TYPES7_3 = "SELECT oid, *, pg_get_userbyid(typowner) as typeowner FROM pg_type WHERE typtype != 'd' AND typtype != 'c'"
Public Const SQL_GET_DOMAINS = "SELECT oid, *, pg_get_userbyid(typowner) as domainowner FROM pg_type WHERE typtype = 'd'"
Public Const SQL_GET_FUNCTIONS7_1 = "SELECT oid, *, pg_get_userbyid(proowner) as funcowner FROM pg_proc"
Public Const SQL_GET_FUNCTIONS7_3 = "SELECT oid, *, pg_get_userbyid(proowner) as funcowner FROM pg_proc WHERE proisagg = FALSE"
Public Const SQL_GET_OPERATORS = "SELECT oid, *, pg_get_userbyid(oprowner) as opowner FROM pg_operator"
Public Const SQL_GET_RULES7_1 = "SELECT oid, rulename, pg_get_ruledef(rulename) as definition FROM pg_rewrite"
Public Const SQL_GET_RULES7_3 = "SELECT oid, rulename, pg_get_ruledef(oid) as definition FROM pg_rewrite"
Public Const SQL_GET_TRIGGERS = "SELECT t.oid, tgname, proname, tgargs, tgtype FROM pg_trigger t, pg_proc p WHERE t.tgfoid = p.oid AND tgisconstraint = FALSE"
Public Const SQL_GET_TABLES7_1 = "SELECT oid, relname, pg_get_userbyid(relowner) as tableowner, relacl FROM pg_class WHERE ((relkind = 'r') OR (relkind = 's'))"
Public Const SQL_GET_TABLES7_2 = "SELECT oid, relname, pg_get_userbyid(relowner) as tableowner, relacl, relhasoids FROM pg_class WHERE ((relkind = 'r') OR (relkind = 's'))"
Public Const SQL_GET_COLUMNS7_1 = "SELECT a.oid, a.attname, a.attnum, CASE WHEN (t.typlen = -1 AND t.typelem != 0) THEN (SELECT at.typname FROM pg_type at WHERE at.oid = t.typelem) || '[]' ELSE t.typname END AS typname, CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS length, a.attnotnull, (SELECT adsrc FROM pg_attrdef d WHERE d.adrelid = a.attrelid AND d.adnum = a.attnum) AS default, (SELECT indisprimary FROM pg_index i, pg_class ic, pg_attribute ia  WHERE i.indrelid = a.attrelid AND i.indexrelid = ic.oid AND ic.oid = ia.attrelid AND ia.attname = a.attname  AND indisprimary IS NOT NULL ORDER BY indisprimary DESC LIMIT 1) AS primarykey FROM pg_attribute a, pg_type t WHERE a.atttypid = t.oid"
Public Const SQL_GET_COLUMNS7_2 = "SELECT 0::oid AS oid, a.attname, a.attnum, CASE WHEN (t.typlen = -1 AND t.typelem != 0) THEN (SELECT at.typname FROM pg_type at WHERE at.oid = t.typelem) || '[]' ELSE t.typname END AS typname, CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS length, a.attnotnull, (SELECT adsrc FROM pg_attrdef d WHERE d.adrelid = a.attrelid AND d.adnum = a.attnum) AS default, (SELECT indisprimary FROM pg_index i, pg_class ic, pg_attribute ia  WHERE i.indrelid = a.attrelid AND i.indexrelid = ic.oid AND ic.oid = ia.attrelid AND ia.attname = a.attname  AND indisprimary IS NOT NULL ORDER BY indisprimary DESC LIMIT 1) AS primarykey, a.attstattarget FROM pg_attribute a, pg_type t WHERE a.atttypid = t.oid"
Public Const SQL_GET_COLUMNS7_3 = "SELECT 0::oid AS oid, a.attname, a.attnum, a.attstorage, CASE WHEN (t.typlen = -1 AND t.typelem != 0) THEN (SELECT at.typname FROM pg_type at WHERE at.oid = t.typelem) || '[]' ELSE t.typname END AS typname, CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS length, a.attnotnull, (SELECT adsrc FROM pg_attrdef d WHERE d.adrelid = a.attrelid AND d.adnum = a.attnum) AS default, (SELECT indisprimary FROM pg_index i, pg_class ic, pg_attribute ia  WHERE i.indrelid = a.attrelid AND i.indexrelid = ic.oid AND ic.oid = ia.attrelid AND ia.attname = a.attname  AND indisprimary IS NOT NULL ORDER BY indisprimary DESC LIMIT 1) AS primarykey, a.attstattarget FROM pg_attribute a, pg_type t WHERE a.atttypid = t.oid AND NOT attisdropped"
Public Const SQL_GET_INDEXES = "SELECT i.oid, i.relname, x.indisunique, x.indisprimary, pg_get_indexdef(i.oid) AS definition FROM pg_index x, pg_class i WHERE i.oid = x.indexrelid"
Public Const SQL_GET_INDEX_COLUMNS = "SELECT attname FROM pg_attribute"
Public Const SQL_GET_CHECKS7_2 = "SELECT rcname, rcsrc FROM pg_relcheck WHERE NOT EXISTS (SELECT * FROM pg_relcheck AS c, pg_inherits AS i WHERE i.inhrelid = pg_relcheck.rcrelid AND (c.rcname = pg_relcheck.rcname OR (c.rcname[0] = '$' AND pg_relcheck.rcname[0] = '$')) AND c.rcsrc = pg_relcheck.rcsrc AND  c.rcrelid = i.inhparent)"
Public Const SQL_GET_CHECKS7_3 = "SELECT conname, consrc FROM pg_constraint WHERE contype = 'c'"
Public Const SQL_GET_INHERITED_TABLES = "SELECT c.relname FROM pg_class c, pg_inherits i WHERE c.oid = i.inhparent"
Public Const SQL_GET_AGGREGATES7_1 = "SELECT oid, aggname, pg_get_userbyid(aggowner) AS owner, aggtransfn, aggfinalfn, aggbasetype, aggtranstype, aggfinaltype, agginitval FROM pg_aggregate"
Public Const SQL_GET_AGGREGATES7_3 = "SELECT oid, proname AS aggname, pg_get_userbyid(proowner) AS owner, aggtransfn, aggfinalfn, proargtypes[0] AS aggbasetype, aggtranstype, prorettype AS aggfinaltype, agginitval FROM pg_aggregate, pg_proc WHERE pg_proc.oid = pg_aggregate.aggfnoid"
Public Const SQL_GET_FOREIGN_KEYS = "SELECT oid, tgrelid, tgconstrname, tgnargs, tgargs, tgdeferrable, tginitdeferred FROM pg_trigger WHERE tgisconstraint = TRUE AND tgtype = 21"
Public Const SQL_GET_NAMESPACES = "SELECT oid, nspname, pg_get_userbyid(nspowner) AS namespaceowner, nspacl FROM pg_namespace"
Public Const SQL_GET_CASTS = "SELECT c.oid, t1.typname AS castsource, t2.typname AS casttarget, p.proname AS castfunc, castcontext FROM pg_cast c, pg_type t1, pg_type t2, pg_proc p WHERE c.castsource = t1.oid AND c.casttarget = t2.oid AND c.castfunc = p.oid"
Public Const SQL_GET_CONVERSIONS = "SELECT c.oid, c.conname, c.condefault, pg_get_userbyid(c.conowner) AS conowner, pg_encoding_to_char(c.conforencoding) as forencoding, pg_encoding_to_char(c.contoencoding) as toencoding, (select quote_ident(n.nspname) FROM pg_namespace n WHERE n.oid=p.pronamespace) || '.' || quote_ident(p.proname) AS procconv FROM pg_conversion c, pg_proc p WHERE p.oid = c.conproc"

'Type Declarations
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

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

'Format an identifier as required
'This code is based on fmtID from the pg_dump code
Public Function fmtID(ByVal szData As String) As String
On Error Resume Next

Dim X As Integer
Dim iVal As Integer

  'Replace double quotes
  szData = Replace(szData, QUOTE, QUOTE & QUOTE)

  If IsNumeric(szData) Then
    szData = QUOTE & szData & QUOTE
  Else
    For X = 1 To Len(szData)
      iVal = Asc(Mid(szData, X, 1))
      If Not ((iVal >= 48) And (iVal <= 57)) And _
         Not ((iVal >= 97) And (iVal <= 122)) And _
         Not (iVal = 95) Then
        szData = QUOTE & szData & QUOTE
        Exit For
      End If
    Next X
  End If

  fmtID = szData

End Function

Public Function fmtTypeID(ByVal szData As String) As String
On Error Resume Next

Dim iLen As Integer
Dim bArray As Boolean
Dim X As Integer
Dim iVal As Integer

  'Replace double quotes
  szData = Replace(szData, QUOTE, QUOTE & QUOTE)
  
  'Dirty hack - if the last 2 chars are [], then this is probably an array specifier
  'so get rid of it.
  If Right(szData, 2) = "[]" Then
    bArray = True
    szData = Left(szData, Len(szData) - 2)
  End If
  
  For X = 1 To Len(szData)
    iVal = Asc(Mid(szData, X, 1))
    If Not ((iVal >= 48) And (iVal <= 57)) And _
       Not ((iVal >= 97) And (iVal <= 122)) And _
       Not (iVal = 95) Then
      szData = QUOTE & szData & QUOTE
      Exit For
    End If
  Next X
  
  If bArray Then
    fmtTypeID = szData & "[]"
  Else
    fmtTypeID = szData
  End If

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
Public Function ParseACL(ByVal szObject As String, ByVal szACL As String, Optional iType As aclType = aclClass) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":ParseACL(" & QUOTE & szObject & QUOTE & ", " & QUOTE & szACL & QUOTE & ")", etFullDebug

Dim szEntries() As String
Dim szEntry As Variant
Dim szName As String
Dim szAccess As String
Dim szSQL As String
Dim szFullObject As String
Dim szTemp As String
  
  szACL = Mid(szACL, 2, Len(szACL) - 2)
  szACL = Replace(szACL, QUOTE, "")
  szEntries = Split(szACL, ",")
  Select Case iType
    Case aclClass
      szFullObject = "TABLE " & szObject
    Case aclDatabase
      szFullObject = "DATABASE " & szObject
    Case aclFunction
      szFullObject = "FUNCTION " & szObject
    Case aclLanguage
      szFullObject = "LANGUAGE " & szObject
    Case aclSchema
      szFullObject = "SCHEMA " & szObject
  End Select
  For Each szEntry In szEntries
  
    'Get the user/group name
    If UCase(Left(szEntry, 6)) = "GROUP " Then
      szName = "GROUP " & fmtID(Mid(szEntry, 7, InStr(1, szEntry, "=") - 7))
    Else
      szName = fmtID(Left(szEntry, InStr(1, szEntry, "=") - 1))
    End If
    If szName = "" Then szName = "PUBLIC"
    
    'Get the Access
    szAccess = Mid(szEntry, InStr(1, szEntry, "=") + 1)
    
    'If the Access is "" then REVOKE all
    If szAccess = "" Then
      szSQL = szSQL & "REVOKE ALL ON " & szFullObject & " FROM " & szName & ";" & vbCrLf
    Else
    
      'Either grant ALL or individual privileges
      'Note that in 7.2+, Delete has been seperated from Update, and References/Trigger
      'have been added.
      If objVersion.VersionNum >= 7.2 Then
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
          If InStr(1, szAccess, "X") <> 0 Then szTemp = szTemp & "EXECUTE, "
          If InStr(1, szAccess, "C") <> 0 Then szTemp = szTemp & "CREATE, "
          If InStr(1, szAccess, "T") <> 0 Then szTemp = szTemp & "TEMP, "
          If InStr(1, szAccess, "U") <> 0 Then szTemp = szTemp & "USAGE, "
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
      
      szSQL = szSQL & "GRANT " & szAccess & " ON " & szFullObject & " TO " & szName & ";" & vbCrLf
    End If
  Next szEntry
  
  ParseACL = szSQL
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
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
