Attribute VB_Name = "basAPI"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basAPI.bas - Contains Windows API Declarations.

Option Explicit

'API Declarations
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Public Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv As Long, ByVal fDirection As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, pcbDescription As Integer) As Integer
Public Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv As Long, phdbc As Long) As Integer
Public Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) As Integer
Public Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc As Long, phstmt As Long) As Integer
Public Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As Long) As Integer
Public Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Public Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal hWnd As Long, ByVal szCSIn As String, ByVal cbCSIn As Integer, ByVal szCSOut As String, ByVal cbCSMax As Integer, cbCSOut As Integer, ByVal fDrvrComp As Integer) As Integer
Public Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hdbc As Long, ByVal fInfoType As Integer, ByRef rgbInfoValue As Any, ByVal cbInfoMax As Integer, cbInfoOut As Integer) As Integer
Public Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hdbc As Long, ByVal fInfoType As Integer, ByVal rgbInfoValue As String, ByVal cbInfoMax As Integer, cbInfoOut As Integer) As Integer
Public Declare Function SQLError Lib "odbc32.dll" (ByVal henv As Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As String, pfNativeError As Long, ByVal szErrorMsg As String, ByVal cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Public Declare Function SQLExecDirect Lib "odbc32.dll" (ByVal hstmt As Long, ByVal szSqlStr As String, ByVal cbSqlStr As Long) As Integer
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberofBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Constants
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const SW_SHOWNORMAL = 1
Public Const ODBC_ADD_DSN = 1            ' Add data source
Public Const ODBC_CONFIG_DSN = 2         ' Configure (edit) data source
Public Const ODBC_REMOVE_DSN = 3         ' Remove data source
Public Const ODBC_ADD_SYS_DSN = 4        ' Add a system DSN
Public Const ODBC_CONFIG_SYS_DSN = 5     ' Configure a system DSN
Public Const ODBC_REMOVE_SYS_DSN = 6     ' Remove a system DSN
Public Const ODBC_REMOVE_DEFAULT_DSN = 7 ' Remove the default DSN

'HTML Help
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_TP_HELP_CONTEXTMENU = &H10
Public Const HH_TP_HELP_WM_HELP = &H11


'SQL Retcodes
Public Const SQL_ERROR As Long = -1
Public Const SQL_INVALID_HANDLE As Long = -2
Public Const SQL_NO_DATA_FOUND As Long = 100
Public Const SQL_SUCCESS As Long = 0
Public Const SQL_SUCCESS_WITH_INFO As Long = 1

'Fetch direction option masks
Public Const SQL_FD_FETCH_NEXT As Long = &H1&
Public Const SQL_FD_FETCH_FIRST As Long = &H2&
Public Const SQL_FD_FETCH_LAST As Long = &H4&
Public Const SQL_FD_FETCH_PRIOR As Long = &H8&
Public Const SQL_FD_FETCH_ABSOLUTE As Long = &H10&
Public Const SQL_FD_FETCH_RELATIVE As Long = &H20&
Public Const SQL_FD_FETCH_RESUME As Long = &H40&
Public Const SQL_FD_FETCH_BOOKMARK As Long = &H80&

'Options for SQLDriverConnect
Public Const SQL_DRIVER_NOPROMPT As Long = 0
Public Const SQL_DRIVER_COMPLETE As Long = 1
Public Const SQL_DRIVER_PROMPT As Long = 2
Public Const SQL_DRIVER_COMPLETE_REQUIRED As Long = 3

'Constants for SQLGetInfo
Public Const SQL_INFO_FIRST As Long = 0
Public Const SQL_ACTIVE_CONNECTIONS As Long = 0
Public Const SQL_ACTIVE_STATEMENTS As Long = 1
Public Const SQL_DATA_SOURCE_NAME As Long = 2
Public Const SQL_DRIVER_HDBC As Long = 3
Public Const SQL_DRIVER_HENV As Long = 4
Public Const SQL_DRIVER_HSTMT As Long = 5
Public Const SQL_DRIVER_NAME As Long = 6
Public Const SQL_DRIVER_VER As Long = 7
Public Const SQL_FETCH_DIRECTION As Long = 8
Public Const SQL_ODBC_API_CONFORMANCE As Long = 9
Public Const SQL_ODBC_VER As Long = 10
Public Const SQL_ROW_UPDATES As Long = 11
Public Const SQL_ODBC_SAG_CLI_CONFORMANCE As Long = 12
Public Const SQL_SERVER_NAME As Long = 13
Public Const SQL_SEARCH_PATTERN_ESCAPE As Long = 14
Public Const SQL_ODBC_SQL_CONFORMANCE As Long = 15
Public Const SQL_DBMS_NAME As Long = 17
Public Const SQL_DBMS_VER As Long = 18
Public Const SQL_ACCESSIBLE_TABLES As Long = 19
Public Const SQL_ACCESSIBLE_PROCEDURES As Long = 20
Public Const SQL_PROCEDURES As Long = 21
Public Const SQL_CONCAT_NULL_BEHAVIOR As Long = 22
Public Const SQL_CURSOR_COMMIT_BEHAVIOR As Long = 23
Public Const SQL_CURSOR_ROLLBACK_BEHAVIOR As Long = 24
Public Const SQL_DATA_SOURCE_READ_ONLY As Long = 25
Public Const SQL_DEFAULT_TXN_ISOLATION As Long = 26
Public Const SQL_EXPRESSIONS_IN_ORDERBY As Long = 27
Public Const SQL_IDENTIFIER_CASE As Long = 28
Public Const SQL_IDENTIFIER_QUOTE_CHAR As Long = 29
Public Const SQL_MAX_COLUMN_NAME_LEN As Long = 30
Public Const SQL_MAX_CURSOR_NAME_LEN As Long = 31
Public Const SQL_MAX_OWNER_NAME_LEN As Long = 32
Public Const SQL_MAX_PROCEDURE_NAME_LEN As Long = 33
Public Const SQL_MAX_QUALIFIER_NAME_LEN As Long = 34
Public Const SQL_MAX_TABLE_NAME_LEN As Long = 35
Public Const SQL_MULT_RESULT_SETS As Long = 36
Public Const SQL_MULTIPLE_ACTIVE_TXN As Long = 37
Public Const SQL_OUTER_JOINS As Long = 38
Public Const SQL_OWNER_TERM As Long = 39
Public Const SQL_PROCEDURE_TERM As Long = 40
Public Const SQL_QUALIFIER_NAME_SEPARATOR As Long = 41
Public Const SQL_QUALIFIER_TERM As Long = 42
Public Const SQL_SCROLL_CONCURRENCY As Long = 43
Public Const SQL_SCROLL_OPTIONS As Long = 44
Public Const SQL_TABLE_TERM As Long = 45
Public Const SQL_TXN_CAPABLE As Long = 46
Public Const SQL_USER_NAME As Long = 47
Public Const SQL_CONVERT_FUNCTIONS As Long = 48
Public Const SQL_NUMERIC_FUNCTIONS As Long = 49
Public Const SQL_STRING_FUNCTIONS As Long = 50
Public Const SQL_SYSTEM_FUNCTIONS As Long = 51
Public Const SQL_TIMEDATE_FUNCTIONS As Long = 52
Public Const SQL_CONVERT_BIGINT As Long = 53
Public Const SQL_CONVERT_BINARY As Long = 54
Public Const SQL_CONVERT_BIT As Long = 55
Public Const SQL_CONVERT_CHAR As Long = 56
Public Const SQL_CONVERT_DATE As Long = 57
Public Const SQL_CONVERT_DECIMAL As Long = 58
Public Const SQL_CONVERT_DOUBLE As Long = 59
Public Const SQL_CONVERT_FLOAT As Long = 60
Public Const SQL_CONVERT_INTEGER As Long = 61
Public Const SQL_CONVERT_LONGVARCHAR As Long = 62
Public Const SQL_CONVERT_NUMERIC As Long = 63
Public Const SQL_CONVERT_REAL As Long = 64
Public Const SQL_CONVERT_SMALLINT As Long = 65
Public Const SQL_CONVERT_TIME As Long = 66
Public Const SQL_CONVERT_TIMESTAMP As Long = 67
Public Const SQL_CONVERT_TINYINT As Long = 68
Public Const SQL_CONVERT_VARBINARY As Long = 69
Public Const SQL_CONVERT_VARCHAR As Long = 70
Public Const SQL_CONVERT_LONGVARBINARY As Long = 71
Public Const SQL_TXN_ISOLATION_OPTION As Long = 72
Public Const SQL_ODBC_SQL_OPT_IEF As Long = 73
Public Const SQL_CORRELATION_NAME As Long = 74
Public Const SQL_NON_NULLABLE_COLUMNS As Long = 75
Public Const SQL_DRIVER_HLIB As Long = 76
Public Const SQL_DRIVER_ODBC_VER As Long = 77
Public Const SQL_LOCK_TYPES As Long = 78
Public Const SQL_POS_OPERATIONS As Long = 79
Public Const SQL_POSITIONED_STATEMENTS As Long = 80
Public Const SQL_GETDATA_EXTENSIONS As Long = 81
Public Const SQL_BOOKMARK_PERSISTENCE As Long = 82
Public Const SQL_STATIC_SENSITIVITY As Long = 83
Public Const SQL_FILE_USAGE As Long = 84
Public Const SQL_NULL_COLLATION As Long = 85
Public Const SQL_ALTER_TABLE As Long = 86
Public Const SQL_COLUMN_ALIAS As Long = 87
Public Const SQL_GROUP_BY As Long = 88
Public Const SQL_KEYWORDS As Long = 89
Public Const SQL_ORDER_BY_COLUMNS_IN_SELECT As Long = 90
Public Const SQL_OWNER_USAGE As Long = 91
Public Const SQL_QUALIFIER_USAGE As Long = 92
Public Const SQL_QUOTED_IDENTIFIER_CASE As Long = 93
Public Const SQL_SPECIAL_CHARACTERS As Long = 94
Public Const SQL_SUBQUERIES As Long = 95
Public Const SQL_UNION As Long = 96
Public Const SQL_MAX_COLUMNS_IN_GROUP_BY As Long = 97
Public Const SQL_MAX_COLUMNS_IN_INDEX As Long = 98
Public Const SQL_MAX_COLUMNS_IN_ORDER_BY As Long = 99
Public Const SQL_MAX_COLUMNS_IN_SELECT As Long = 100
Public Const SQL_MAX_COLUMNS_IN_TABLE As Long = 101
Public Const SQL_MAX_INDEX_SIZE As Long = 102
Public Const SQL_MAX_ROW_SIZE_INCLUDES_LONG As Long = 103
Public Const SQL_MAX_ROW_SIZE As Long = 104
Public Const SQL_MAX_STATEMENT_LEN As Long = 105
Public Const SQL_MAX_TABLES_IN_SELECT As Long = 106
Public Const SQL_MAX_USER_NAME_LEN As Long = 107
Public Const SQL_MAX_CHAR_LITERAL_LEN As Long = 108
Public Const SQL_TIMEDATE_ADD_INTERVALS As Long = 109
Public Const SQL_TIMEDATE_DIFF_INTERVALS As Long = 110
Public Const SQL_NEED_LONG_DATA_LEN As Long = 111
Public Const SQL_MAX_BINARY_LITERAL_LEN As Long = 112
Public Const SQL_LIKE_ESCAPE_CLAUSE As Long = 113
Public Const SQL_QUALIFIER_LOCATION As Long = 114
Public Const SQL_INFO_LAST As Long = SQL_QUALIFIER_LOCATION
Public Const SQL_INFO_DRIVER_START As Long = 1000

'Enumerated Constants
Public Enum RegTypes
   regNull = 0
   regString = 1
   regXString = 2
   regBinary = 3
   regDWord = 4
   regLink = 6
   regMultiString = 7
   regResList = 8
End Enum

Public Enum RegHives
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

'WinInet constants
Public Const scUserAgent = "http sample"
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const INTERNET_OPTION_VERSION = 40
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

'WinInet DLL Version Structure
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

'Listviews
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2 'Note: On last column, its width fills remaining width
                                                   'of list-view according to Micro$oft. This does not
                                                   'appear to be the case when I do it.
