Attribute VB_Name = "basAPI"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Functions
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal S As String) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Public Declare Function IcmpCreateFile Lib "ICMP.DLL" () As Long
Public Declare Function IcmpSendEcho Lib "ICMP.DLL" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, RequestOptions As ICMP_OPTIONS, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Declare Function IcmpCloseHandle Lib "ICMP.DLL" (ByVal IcmpHandle As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal nBytes As Long)
Public Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Public Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Public Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv&, phdbc&) As Integer
Public Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv&) As Integer
Public Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc&, ByVal hWnd As Long, ByVal szCSIn$, ByVal cbCSIn%, ByVal szCSOut$, ByVal cbCSMax%, cbCSOut%, ByVal fDrvrComp%) As Integer
Public Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hdbc&, ByVal fInfoType%, ByVal rgbInfoValue As String, ByVal cbInfoMax%, cbInfoOut%) As Integer
Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hdbc&, ByVal fInfoType%, ByRef rgbInfoValue As Any, ByVal cbInfoMax%, cbInfoOut%) As Integer
Public Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv&) As Integer
Public Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Public Declare Function SQLError Lib "odbc32.dll" (ByVal henv&, ByVal hdbc&, ByVal hstmt&, ByVal szSqlState$, pfNativeError&, ByVal szErrorMsg$, ByVal cbErrorMsgMax%, pcbErrorMsg%) As Integer

'Constants
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128
Public Const AF_INET = 2

'SQL Retcodes
Public Const SQL_ERROR As Long = -1
Public Const SQL_INVALID_HANDLE As Long = -2
Public Const SQL_NO_DATA_FOUND As Long = 100
Public Const SQL_SUCCESS As Long = 0
Public Const SQL_SUCCESS_WITH_INFO As Long = 1

'Fetch direction option masks
Global Const SQL_FD_FETCH_NEXT As Long = &H1&
Global Const SQL_FD_FETCH_FIRST As Long = &H2&
Global Const SQL_FD_FETCH_LAST As Long = &H4&
Global Const SQL_FD_FETCH_PRIOR As Long = &H8&
Global Const SQL_FD_FETCH_ABSOLUTE As Long = &H10&
Global Const SQL_FD_FETCH_RELATIVE As Long = &H20&
Global Const SQL_FD_FETCH_RESUME As Long = &H40&
Global Const SQL_FD_FETCH_BOOKMARK As Long = &H80&

'Options for SQLDriverConnect
Public Const SQL_DRIVER_NOPROMPT As Long = 0
Public Const SQL_DRIVER_COMPLETE As Long = 1
Public Const SQL_DRIVER_PROMPT As Long = 2
Public Const SQL_DRIVER_COMPLETE_REQUIRED As Long = 3

'Defines for SQLGetInfo
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

'UDTs
Public Type ICMP_OPTIONS
    TTL             As Byte
    tos             As Byte
    flags           As Byte
    OptionsSize     As Long
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    datasize        As Integer
    reserved        As Integer
    DataPointer     As Long
    options         As ICMP_OPTIONS
    ReturnedData    As String * 256
End Type

Public Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  iMaxSockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type
