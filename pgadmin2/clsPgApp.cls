VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPgApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' clsPgApp.cls - class export application function/form

Public frmAbout As pgAdmin2.frmAbout                        'frmAbout.frm
Public frmAggregate As pgAdmin2.frmAggregate                'frmAggregate.frm
Public frmCast As pgAdmin2.frmCast                          'frmCast.frm
Public frmClone As pgAdmin2.frmClone                        'frmClone.frm
Public frmColumn As pgAdmin2.frmColumn                      'frmColumn.frm
Public frmConnect As pgAdmin2.frmConnect                    'frmConnect.frm
Public frmConversion As pgAdmin2.frmConversion              'frmConversion.frm
Public frmDatabase As pgAdmin2.frmDatabase                  'frmDatabase.frm
Public frmDomain As pgAdmin2.frmDomain                      'frmDomain.frm
Public frmDummy As pgAdmin2.frmDummy                        'frmDummy.frm
Public frmError As pgAdmin2.frmError                        'frmError.frm
Public frmFind As pgAdmin2.frmFind                          'frmFind.frm
Public frmForeignKey As pgAdmin2.frmForeignKey              'frmForeignKey.frm
Public frmFunction As pgAdmin2.frmFunction                  'frmFunction.frm
Public frmGroup As pgAdmin2.frmGroup                        'frmGroup.frm
Public frmImportColumn As pgAdmin2.frmImportColumn          'frmImportColumn.frm
Public frmIndex As pgAdmin2.frmIndex                        'frmIndex.frm
Public frmLanguage As pgAdmin2.frmLanguage                  'frmLanguage.frm
Public frmLog As pgAdmin2.frmLog                            'frmLog.frm
Public frmMain As pgAdmin2.frmMain                          'frmMain.frm
Public frmNamespace As pgAdmin2.frmNamespace                'frmNamespace.frm
Public frmOperator As pgAdmin2.frmOperator                  'frmOperator.frm
Public frmOptions As pgAdmin2.frmOptions                    'frmOptions.frm
Public frmPassword As pgAdmin2.frmPassword                  'frmPassword.frm
Public frmRecordLog As pgAdmin2.frmRecordLog                'frmRecordLog.frm
Public frmRule As pgAdmin2.frmRule                          'frmRule.frm
Public frmSequence As pgAdmin2.frmSequence                  'frmSequence.frm
Public frmServer As pgAdmin2.frmServer                      'frmServer.frm
Public frmSplash As pgAdmin2.frmSplash                      'frmSplash.frm
Public frmSQLExplain As pgAdmin2.frmSQLExplain              'frmSQLExplain.frm
Public frmSQLInput As pgAdmin2.frmSQLInput                  'frmSQLInput.frm
Public frmSQLOutput As pgAdmin2.frmSQLOutput                'frmSQLOutput.frm
Public frmSQLWizard As pgAdmin2.frmSQLWizard                'frmSQLWizard.frm
Public frmTable As pgAdmin2.frmTable                        'frmTable.frm
Public frmTip As pgAdmin2.frmTip                            'frmTip.frm
Public frmTrigger As pgAdmin2.frmTrigger                    'frmTrigger.frm
Public frmType As pgAdmin2.frmType                          'frmType.frm
Public frmUpgradeWizard As pgAdmin2.frmUpgradeWizard        'frmUpgradeWizard.frm
Public frmUser As pgAdmin2.frmUser                          'frmUser.frm
Public frmView As pgAdmin2.frmView                          'frmView.frm

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basActions
'File: basActions.bas
Public Sub Vacuum(bAnalyse As Boolean)
  Call basActions.Vacuum(bAnalyse)
End Sub

Public Sub Reindex()
  Call basActions.Reindex()

End Sub

Public Sub Drop()
  Call basActions.Drop()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basAPI
'File: basAPI.bas
''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basClone
'File: basClone.bas
'Inizialize clone object
Public Sub InitClone()
  Call basClone.InitClone()

End Sub

'clear object database
Public Sub ClearObjDb()
  Call basClone.ClearObjDb()

End Sub

'copy object database
Public Sub CopyObjDb()
  Call basClone.CopyObjDb()

End Sub

'paste object database
Public Sub PasteObjDb()
  Call basClone.PasteObjDb()

End Sub

'clone type
Public Function CloneType(szNewName As String, szDatabase As String, szNamespace As String) 'As pgType
'  CloneType = basClone.CloneType(szNewName,szDatabase,szNamespace)
End Function

'clone cast
Public Function CloneCast(szDatabase As String) 'As pgCast
'  CloneCast = basClone.CloneCast(szDatabase)
End Function

'clone operator
Public Function CloneOperator(szNewName As String, szDatabase As String, szNamespace As String) 'As pgOperator
'  CloneOperator = basClone.CloneOperator(szNewName,szDatabase,szNamespace)
End Function

'clone aggregate
Public Function CloneAggregate(szNewName As String, szDatabase As String, szNamespace As String) 'As pgAggregate
'  CloneAggregate = basClone.CloneAggregate(szNewName,szDatabase,szNamespace)
End Function

'clone function
Public Function CloneFunction(szNewName As String, szDatabase As String, szNamespace As String) 'As pgFunction
'  CloneFunction = basClone.CloneFunction(szNewName,szDatabase,szNamespace)
End Function

'clone domain
Public Function CloneDomain(szNewName As String, szDatabase As String, szNamespace As String) 'As pgDomain
'  CloneDomain = basClone.CloneDomain(szNewName,szDatabase,szNamespace)
End Function

'clone user
Public Function CloneUser(szNewName As String) 'As pgUser
'  CloneUser = basClone.CloneUser(szNewName)
End Function

'clone group
Public Function CloneGroup(szNewName As String) 'As pgGroup
'  CloneGroup = basClone.CloneGroup(szNewName)
End Function

'clone view
Public Function CloneView(szNewName As String, szDatabase As String, szNamespace As String) 'As pgView
'  CloneView = basClone.CloneView(szNewName,szDatabase,szNamespace)
End Function

'clone table
Public Function CloneTable(szNewName As String, szDatabase As String, szNamespace As String, Optional bCopyData As Boolean = False) 'As pgTable
'  CloneTable = basClone.CloneTable(szNewName,szDatabase,szNamespace,bCopyData,=,False)
End Function

Public Sub CloneAcl(objDb As Variant)
  Call basClone.CloneAcl(objDb)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basGlobal
'File: basGlobal.bas
''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basMisc
'File: basMisc.bas
Public Sub Main()
  Call basMisc.Main()

End Sub

Public Function GetID() As String
  GetID = basMisc.GetID()
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
  SetTopMostWindow = basMisc.SetTopMostWindow(hwnd,Topmost)
End Function

Public Sub BuildConnectionMenu()
  Call basMisc.BuildConnectionMenu()

End Sub

Public Sub BuildPluginsMenu()
  Call basMisc.BuildPluginsMenu()

End Sub

Public Sub LogError(lError As Long, szError As String, szRoutine As String)
  Call basMisc.LogError(lError,szError,szRoutine)
End Sub

Public Sub StartMsg(ByVal szMsg As String)
  Call basMisc.StartMsg(szMsg)
End Sub

Public Sub EndMsg(Optional ByVal szErr As String)
  Call basMisc.EndMsg(szErr)
End Sub

Public Function dbSZ(szData As String) As String
  dbSZ = basMisc.dbSZ(szData)
End Function

'This code is based on fmtID from the pg_dump code
Public Function fmtID(ByVal szData As String) As String
  fmtID = basMisc.fmtID(szData)
End Function

Public Function Bool2Bin(bData As Boolean) As Integer
  Bool2Bin = basMisc.Bool2Bin(bData)
End Function

Public Function Bin2Bool(iData As Integer) As Boolean
  Bin2Bool = basMisc.Bin2Bool(iData)
End Function

'Parse an ACL and return | delimited User/Access lists
Public Sub ParseACL(ByVal szACL As String, ByRef szUserlist As String, ByRef szAccesslist As String)
  Call basMisc.ParseACL(szACL,szUserlist,szAccesslist)
End Sub

'Format a typename
Public Function fmtTypeName(objType As pgType) As String
  fmtTypeName = basMisc.fmtTypeName(objType)
End Function

Public Sub AutoSizeColumnLv(lv As ListView)
  Call basMisc.AutoSizeColumnLv(lv)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basPatch
'File: basPatch.bas
'Patch form
Public Sub PatchForm(objForm As Form)
  Call basPatch.PatchForm(objForm)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basRegistry
'File: basRegistry.bas
Public Function RegRead(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, Optional Default As Variant) As String
  RegRead = basRegistry.RegRead(Hive,Section,Key,Default)
End Function

Public Sub RegWrite(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, ByVal ValType As RegTypes, ByVal Value As Variant)
  Call basRegistry.RegWrite(Hive,Section,Key,ValType,Value)
End Sub

Public Function RegGetSubkey(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As String
  RegGetSubkey = basRegistry.RegGetSubkey(Hive,Section,Idx)
End Function

Public Function RegReadAll(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As Variant
  RegReadAll = basRegistry.RegReadAll(Hive,Section,Idx)
End Function

Public Sub RegDelSubkey(ByVal Hive As RegHives, ByVal Section As String)
  Call basRegistry.RegDelSubkey(Hive,Section)
End Sub

Public Sub RegDelValue(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String)
  Call basRegistry.RegDelValue(Hive,Section,Key)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basVarDb
'File: basVarDb.bas
'Initalization
Public Sub InitVarDb()
  Call basVarDb.InitVarDb()

End Sub

'return the definition var db from Collection
Public Function GetVarDb(szName As String) 'As VarDb
'  GetVarDb = basVarDb.GetVarDb(szName)
End Function

'return the image name of value
Public Function GetImageFromVal(szValue As String, TypeVar As TypeVarDb) As String
  GetImageFromVal = basVarDb.GetImageFromVal(szValue,TypeVar)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal routine
'create collection forms activate
Public Function FormsActivate() As Collection
Dim objCol As New Collection
Dim objTmp

  For Each objTmp In VB.Forms
    objCol.Add objTmp
  Next
  Set FormsActivate = objCol
End Function

'get forms by name
Public Function FormByName(vData As String) As Form
Dim objTmp As Form
Dim objFrm As Form

  For Each objTmp In VB.Forms
    If LCase(objTmp.Name) = LCase(vData) Then
      Set objFrm = objTmp
      Exit For
    End If
  Next
  Set GetFormByName = objFrm
End Function

'return value of the const/enumerate by name
Public Function ConstByName(szName As String)

  Select Case LCase(szName)
    case LCase("SWP_NOMOVE")
      ConstByName=SWP_NOMOVE

    case LCase("SWP_NOSIZE")
      ConstByName=SWP_NOSIZE

    case LCase("SWP_NOACTIVATE")
      ConstByName=SWP_NOACTIVATE

    case LCase("FLAGS")
      ConstByName=FLAGS

    case LCase("CW_USEDEFAULT")
      ConstByName=CW_USEDEFAULT

    case LCase("WM_USER")
      ConstByName=WM_USER

    case LCase("TTF_CENTERTIP")
      ConstByName=TTF_CENTERTIP

    case LCase("TTF_SUBCLASS")
      ConstByName=TTF_SUBCLASS

    case LCase("TTM_ADDTOOLA")
      ConstByName=TTM_ADDTOOLA

    case LCase("TTM_GETTEXTA")
      ConstByName=TTM_GETTEXTA

    case LCase("TTM_SETMAXTIPWIDTH")
      ConstByName=TTM_SETMAXTIPWIDTH

    case LCase("TTM_SETTIPBKCOLOR")
      ConstByName=TTM_SETTIPBKCOLOR

    case LCase("TTM_SETTIPTEXTCOLOR")
      ConstByName=TTM_SETTIPTEXTCOLOR

    case LCase("TTM_UPDATETIPTEXTA")
      ConstByName=TTM_UPDATETIPTEXTA

    case LCase("TTS_ALWAYSTIP")
      ConstByName=TTS_ALWAYSTIP

    case LCase("TTS_BALLOON")
      ConstByName=TTS_BALLOON

    case LCase("HWND_TOPMOST")
      ConstByName=HWND_TOPMOST

    case LCase("HWND_NOTOPMOST")
      ConstByName=HWND_NOTOPMOST

    case LCase("READ_CONTROL")
      ConstByName=READ_CONTROL

    case LCase("SYNCHRONIZE")
      ConstByName=SYNCHRONIZE

    case LCase("STANDARD_RIGHTS_ALL")
      ConstByName=STANDARD_RIGHTS_ALL

    case LCase("STANDARD_RIGHTS_READ")
      ConstByName=STANDARD_RIGHTS_READ

    case LCase("STANDARD_RIGHTS_WRITE")
      ConstByName=STANDARD_RIGHTS_WRITE

    case LCase("KEY_QUERY_VALUE")
      ConstByName=KEY_QUERY_VALUE

    case LCase("KEY_SET_VALUE")
      ConstByName=KEY_SET_VALUE

    case LCase("KEY_CREATE_SUB_KEY")
      ConstByName=KEY_CREATE_SUB_KEY

    case LCase("KEY_ENUMERATE_SUB_KEYS")
      ConstByName=KEY_ENUMERATE_SUB_KEYS

    case LCase("KEY_NOTIFY")
      ConstByName=KEY_NOTIFY

    case LCase("KEY_CREATE_LINK")
      ConstByName=KEY_CREATE_LINK

    case LCase("KEY_ALL_ACCESS")
      ConstByName=KEY_ALL_ACCESS

    case LCase("KEY_READ")
      ConstByName=KEY_READ

    case LCase("KEY_EXECUTE")
      ConstByName=KEY_EXECUTE

    case LCase("KEY_WRITE")
      ConstByName=KEY_WRITE

    case LCase("REG_NONE")
      ConstByName=REG_NONE

    case LCase("REG_SZ")
      ConstByName=REG_SZ

    case LCase("REG_EXPAND_SZ")
      ConstByName=REG_EXPAND_SZ

    case LCase("REG_BINARY")
      ConstByName=REG_BINARY

    case LCase("REG_DWORD")
      ConstByName=REG_DWORD

    case LCase("REG_LINK")
      ConstByName=REG_LINK

    case LCase("REG_MULTI_SZ")
      ConstByName=REG_MULTI_SZ

    case LCase("REG_RESOURCE_LIST")
      ConstByName=REG_RESOURCE_LIST

    case LCase("ERROR_NONE")
      ConstByName=ERROR_NONE

    case LCase("ERROR_BADDB")
      ConstByName=ERROR_BADDB

    case LCase("ERROR_BADKEY")
      ConstByName=ERROR_BADKEY

    case LCase("ERROR_CANTOPEN")
      ConstByName=ERROR_CANTOPEN

    case LCase("ERROR_CANTREAD")
      ConstByName=ERROR_CANTREAD

    case LCase("ERROR_CANTWRITE")
      ConstByName=ERROR_CANTWRITE

    case LCase("ERROR_OUTOFMEMORY")
      ConstByName=ERROR_OUTOFMEMORY

    case LCase("ERROR_INVALID_PARAMETER")
      ConstByName=ERROR_INVALID_PARAMETER

    case LCase("ERROR_ACCESS_DENIED")
      ConstByName=ERROR_ACCESS_DENIED

    case LCase("ERROR_INVALID_PARAMETERS")
      ConstByName=ERROR_INVALID_PARAMETERS

    case LCase("ERROR_NO_MORE_ITEMS")
      ConstByName=ERROR_NO_MORE_ITEMS

    case LCase("SW_SHOWNORMAL")
      ConstByName=SW_SHOWNORMAL

    case LCase("SW_SHOW")
      ConstByName=SW_SHOW

    case LCase("ODBC_ADD_DSN")
      ConstByName=ODBC_ADD_DSN

    case LCase("ODBC_CONFIG_DSN")
      ConstByName=ODBC_CONFIG_DSN

    case LCase("ODBC_REMOVE_DSN")
      ConstByName=ODBC_REMOVE_DSN

    case LCase("ODBC_ADD_SYS_DSN")
      ConstByName=ODBC_ADD_SYS_DSN

    case LCase("ODBC_CONFIG_SYS_DSN")
      ConstByName=ODBC_CONFIG_SYS_DSN

    case LCase("ODBC_REMOVE_SYS_DSN")
      ConstByName=ODBC_REMOVE_SYS_DSN

    case LCase("ODBC_REMOVE_DEFAULT_DSN")
      ConstByName=ODBC_REMOVE_DEFAULT_DSN

    case LCase("HH_DISPLAY_TOPIC")
      ConstByName=HH_DISPLAY_TOPIC

    case LCase("HH_SET_WIN_TYPE")
      ConstByName=HH_SET_WIN_TYPE

    case LCase("HH_GET_WIN_TYPE")
      ConstByName=HH_GET_WIN_TYPE

    case LCase("HH_GET_WIN_HANDLE")
      ConstByName=HH_GET_WIN_HANDLE

    case LCase("HH_DISPLAY_TEXT_POPUP")
      ConstByName=HH_DISPLAY_TEXT_POPUP

    case LCase("HH_HELP_CONTEXT")
      ConstByName=HH_HELP_CONTEXT

    case LCase("HH_TP_HELP_CONTEXTMENU")
      ConstByName=HH_TP_HELP_CONTEXTMENU

    case LCase("HH_TP_HELP_WM_HELP")
      ConstByName=HH_TP_HELP_WM_HELP

    case LCase("SQL_ERROR")
      ConstByName=SQL_ERROR

    case LCase("SQL_INVALID_HANDLE")
      ConstByName=SQL_INVALID_HANDLE

    case LCase("SQL_NO_DATA_FOUND")
      ConstByName=SQL_NO_DATA_FOUND

    case LCase("SQL_SUCCESS")
      ConstByName=SQL_SUCCESS

    case LCase("SQL_SUCCESS_WITH_INFO")
      ConstByName=SQL_SUCCESS_WITH_INFO

    case LCase("SQL_FD_FETCH_NEXT")
      ConstByName=SQL_FD_FETCH_NEXT

    case LCase("SQL_FD_FETCH_FIRST")
      ConstByName=SQL_FD_FETCH_FIRST

    case LCase("SQL_FD_FETCH_LAST")
      ConstByName=SQL_FD_FETCH_LAST

    case LCase("SQL_FD_FETCH_PRIOR")
      ConstByName=SQL_FD_FETCH_PRIOR

    case LCase("SQL_FD_FETCH_ABSOLUTE")
      ConstByName=SQL_FD_FETCH_ABSOLUTE

    case LCase("SQL_FD_FETCH_RELATIVE")
      ConstByName=SQL_FD_FETCH_RELATIVE

    case LCase("SQL_FD_FETCH_RESUME")
      ConstByName=SQL_FD_FETCH_RESUME

    case LCase("SQL_FD_FETCH_BOOKMARK")
      ConstByName=SQL_FD_FETCH_BOOKMARK

    case LCase("SQL_DRIVER_NOPROMPT")
      ConstByName=SQL_DRIVER_NOPROMPT

    case LCase("SQL_DRIVER_COMPLETE")
      ConstByName=SQL_DRIVER_COMPLETE

    case LCase("SQL_DRIVER_PROMPT")
      ConstByName=SQL_DRIVER_PROMPT

    case LCase("SQL_DRIVER_COMPLETE_REQUIRED")
      ConstByName=SQL_DRIVER_COMPLETE_REQUIRED

    case LCase("SQL_INFO_FIRST")
      ConstByName=SQL_INFO_FIRST

    case LCase("SQL_ACTIVE_CONNECTIONS")
      ConstByName=SQL_ACTIVE_CONNECTIONS

    case LCase("SQL_ACTIVE_STATEMENTS")
      ConstByName=SQL_ACTIVE_STATEMENTS

    case LCase("SQL_DATA_SOURCE_NAME")
      ConstByName=SQL_DATA_SOURCE_NAME

    case LCase("SQL_DRIVER_HDBC")
      ConstByName=SQL_DRIVER_HDBC

    case LCase("SQL_DRIVER_HENV")
      ConstByName=SQL_DRIVER_HENV

    case LCase("SQL_DRIVER_HSTMT")
      ConstByName=SQL_DRIVER_HSTMT

    case LCase("SQL_DRIVER_NAME")
      ConstByName=SQL_DRIVER_NAME

    case LCase("SQL_DRIVER_VER")
      ConstByName=SQL_DRIVER_VER

    case LCase("SQL_FETCH_DIRECTION")
      ConstByName=SQL_FETCH_DIRECTION

    case LCase("SQL_ODBC_API_CONFORMANCE")
      ConstByName=SQL_ODBC_API_CONFORMANCE

    case LCase("SQL_ODBC_VER")
      ConstByName=SQL_ODBC_VER

    case LCase("SQL_ROW_UPDATES")
      ConstByName=SQL_ROW_UPDATES

    case LCase("SQL_ODBC_SAG_CLI_CONFORMANCE")
      ConstByName=SQL_ODBC_SAG_CLI_CONFORMANCE

    case LCase("SQL_SERVER_NAME")
      ConstByName=SQL_SERVER_NAME

    case LCase("SQL_SEARCH_PATTERN_ESCAPE")
      ConstByName=SQL_SEARCH_PATTERN_ESCAPE

    case LCase("SQL_ODBC_SQL_CONFORMANCE")
      ConstByName=SQL_ODBC_SQL_CONFORMANCE

    case LCase("SQL_DBMS_NAME")
      ConstByName=SQL_DBMS_NAME

    case LCase("SQL_DBMS_VER")
      ConstByName=SQL_DBMS_VER

    case LCase("SQL_ACCESSIBLE_TABLES")
      ConstByName=SQL_ACCESSIBLE_TABLES

    case LCase("SQL_ACCESSIBLE_PROCEDURES")
      ConstByName=SQL_ACCESSIBLE_PROCEDURES

    case LCase("SQL_PROCEDURES")
      ConstByName=SQL_PROCEDURES

    case LCase("SQL_CONCAT_NULL_BEHAVIOR")
      ConstByName=SQL_CONCAT_NULL_BEHAVIOR

    case LCase("SQL_CURSOR_COMMIT_BEHAVIOR")
      ConstByName=SQL_CURSOR_COMMIT_BEHAVIOR

    case LCase("SQL_CURSOR_ROLLBACK_BEHAVIOR")
      ConstByName=SQL_CURSOR_ROLLBACK_BEHAVIOR

    case LCase("SQL_DATA_SOURCE_READ_ONLY")
      ConstByName=SQL_DATA_SOURCE_READ_ONLY

    case LCase("SQL_DEFAULT_TXN_ISOLATION")
      ConstByName=SQL_DEFAULT_TXN_ISOLATION

    case LCase("SQL_EXPRESSIONS_IN_ORDERBY")
      ConstByName=SQL_EXPRESSIONS_IN_ORDERBY

    case LCase("SQL_IDENTIFIER_CASE")
      ConstByName=SQL_IDENTIFIER_CASE

    case LCase("SQL_IDENTIFIER_QUOTE_CHAR")
      ConstByName=SQL_IDENTIFIER_QUOTE_CHAR

    case LCase("SQL_MAX_COLUMN_NAME_LEN")
      ConstByName=SQL_MAX_COLUMN_NAME_LEN

    case LCase("SQL_MAX_CURSOR_NAME_LEN")
      ConstByName=SQL_MAX_CURSOR_NAME_LEN

    case LCase("SQL_MAX_OWNER_NAME_LEN")
      ConstByName=SQL_MAX_OWNER_NAME_LEN

    case LCase("SQL_MAX_PROCEDURE_NAME_LEN")
      ConstByName=SQL_MAX_PROCEDURE_NAME_LEN

    case LCase("SQL_MAX_QUALIFIER_NAME_LEN")
      ConstByName=SQL_MAX_QUALIFIER_NAME_LEN

    case LCase("SQL_MAX_TABLE_NAME_LEN")
      ConstByName=SQL_MAX_TABLE_NAME_LEN

    case LCase("SQL_MULT_RESULT_SETS")
      ConstByName=SQL_MULT_RESULT_SETS

    case LCase("SQL_MULTIPLE_ACTIVE_TXN")
      ConstByName=SQL_MULTIPLE_ACTIVE_TXN

    case LCase("SQL_OUTER_JOINS")
      ConstByName=SQL_OUTER_JOINS

    case LCase("SQL_OWNER_TERM")
      ConstByName=SQL_OWNER_TERM

    case LCase("SQL_PROCEDURE_TERM")
      ConstByName=SQL_PROCEDURE_TERM

    case LCase("SQL_QUALIFIER_NAME_SEPARATOR")
      ConstByName=SQL_QUALIFIER_NAME_SEPARATOR

    case LCase("SQL_QUALIFIER_TERM")
      ConstByName=SQL_QUALIFIER_TERM

    case LCase("SQL_SCROLL_CONCURRENCY")
      ConstByName=SQL_SCROLL_CONCURRENCY

    case LCase("SQL_SCROLL_OPTIONS")
      ConstByName=SQL_SCROLL_OPTIONS

    case LCase("SQL_TABLE_TERM")
      ConstByName=SQL_TABLE_TERM

    case LCase("SQL_TXN_CAPABLE")
      ConstByName=SQL_TXN_CAPABLE

    case LCase("SQL_USER_NAME")
      ConstByName=SQL_USER_NAME

    case LCase("SQL_CONVERT_FUNCTIONS")
      ConstByName=SQL_CONVERT_FUNCTIONS

    case LCase("SQL_NUMERIC_FUNCTIONS")
      ConstByName=SQL_NUMERIC_FUNCTIONS

    case LCase("SQL_STRING_FUNCTIONS")
      ConstByName=SQL_STRING_FUNCTIONS

    case LCase("SQL_SYSTEM_FUNCTIONS")
      ConstByName=SQL_SYSTEM_FUNCTIONS

    case LCase("SQL_TIMEDATE_FUNCTIONS")
      ConstByName=SQL_TIMEDATE_FUNCTIONS

    case LCase("SQL_CONVERT_BIGINT")
      ConstByName=SQL_CONVERT_BIGINT

    case LCase("SQL_CONVERT_BINARY")
      ConstByName=SQL_CONVERT_BINARY

    case LCase("SQL_CONVERT_BIT")
      ConstByName=SQL_CONVERT_BIT

    case LCase("SQL_CONVERT_CHAR")
      ConstByName=SQL_CONVERT_CHAR

    case LCase("SQL_CONVERT_DATE")
      ConstByName=SQL_CONVERT_DATE

    case LCase("SQL_CONVERT_DECIMAL")
      ConstByName=SQL_CONVERT_DECIMAL

    case LCase("SQL_CONVERT_DOUBLE")
      ConstByName=SQL_CONVERT_DOUBLE

    case LCase("SQL_CONVERT_FLOAT")
      ConstByName=SQL_CONVERT_FLOAT

    case LCase("SQL_CONVERT_INTEGER")
      ConstByName=SQL_CONVERT_INTEGER

    case LCase("SQL_CONVERT_LONGVARCHAR")
      ConstByName=SQL_CONVERT_LONGVARCHAR

    case LCase("SQL_CONVERT_NUMERIC")
      ConstByName=SQL_CONVERT_NUMERIC

    case LCase("SQL_CONVERT_REAL")
      ConstByName=SQL_CONVERT_REAL

    case LCase("SQL_CONVERT_SMALLINT")
      ConstByName=SQL_CONVERT_SMALLINT

    case LCase("SQL_CONVERT_TIME")
      ConstByName=SQL_CONVERT_TIME

    case LCase("SQL_CONVERT_TIMESTAMP")
      ConstByName=SQL_CONVERT_TIMESTAMP

    case LCase("SQL_CONVERT_TINYINT")
      ConstByName=SQL_CONVERT_TINYINT

    case LCase("SQL_CONVERT_VARBINARY")
      ConstByName=SQL_CONVERT_VARBINARY

    case LCase("SQL_CONVERT_VARCHAR")
      ConstByName=SQL_CONVERT_VARCHAR

    case LCase("SQL_CONVERT_LONGVARBINARY")
      ConstByName=SQL_CONVERT_LONGVARBINARY

    case LCase("SQL_TXN_ISOLATION_OPTION")
      ConstByName=SQL_TXN_ISOLATION_OPTION

    case LCase("SQL_ODBC_SQL_OPT_IEF")
      ConstByName=SQL_ODBC_SQL_OPT_IEF

    case LCase("SQL_CORRELATION_NAME")
      ConstByName=SQL_CORRELATION_NAME

    case LCase("SQL_NON_NULLABLE_COLUMNS")
      ConstByName=SQL_NON_NULLABLE_COLUMNS

    case LCase("SQL_DRIVER_HLIB")
      ConstByName=SQL_DRIVER_HLIB

    case LCase("SQL_DRIVER_ODBC_VER")
      ConstByName=SQL_DRIVER_ODBC_VER

    case LCase("SQL_LOCK_TYPES")
      ConstByName=SQL_LOCK_TYPES

    case LCase("SQL_POS_OPERATIONS")
      ConstByName=SQL_POS_OPERATIONS

    case LCase("SQL_POSITIONED_STATEMENTS")
      ConstByName=SQL_POSITIONED_STATEMENTS

    case LCase("SQL_GETDATA_EXTENSIONS")
      ConstByName=SQL_GETDATA_EXTENSIONS

    case LCase("SQL_BOOKMARK_PERSISTENCE")
      ConstByName=SQL_BOOKMARK_PERSISTENCE

    case LCase("SQL_STATIC_SENSITIVITY")
      ConstByName=SQL_STATIC_SENSITIVITY

    case LCase("SQL_FILE_USAGE")
      ConstByName=SQL_FILE_USAGE

    case LCase("SQL_NULL_COLLATION")
      ConstByName=SQL_NULL_COLLATION

    case LCase("SQL_ALTER_TABLE")
      ConstByName=SQL_ALTER_TABLE

    case LCase("SQL_COLUMN_ALIAS")
      ConstByName=SQL_COLUMN_ALIAS

    case LCase("SQL_GROUP_BY")
      ConstByName=SQL_GROUP_BY

    case LCase("SQL_KEYWORDS")
      ConstByName=SQL_KEYWORDS

    case LCase("SQL_ORDER_BY_COLUMNS_IN_SELECT")
      ConstByName=SQL_ORDER_BY_COLUMNS_IN_SELECT

    case LCase("SQL_OWNER_USAGE")
      ConstByName=SQL_OWNER_USAGE

    case LCase("SQL_QUALIFIER_USAGE")
      ConstByName=SQL_QUALIFIER_USAGE

    case LCase("SQL_QUOTED_IDENTIFIER_CASE")
      ConstByName=SQL_QUOTED_IDENTIFIER_CASE

    case LCase("SQL_SPECIAL_CHARACTERS")
      ConstByName=SQL_SPECIAL_CHARACTERS

    case LCase("SQL_SUBQUERIES")
      ConstByName=SQL_SUBQUERIES

    case LCase("SQL_UNION")
      ConstByName=SQL_UNION

    case LCase("SQL_MAX_COLUMNS_IN_GROUP_BY")
      ConstByName=SQL_MAX_COLUMNS_IN_GROUP_BY

    case LCase("SQL_MAX_COLUMNS_IN_INDEX")
      ConstByName=SQL_MAX_COLUMNS_IN_INDEX

    case LCase("SQL_MAX_COLUMNS_IN_ORDER_BY")
      ConstByName=SQL_MAX_COLUMNS_IN_ORDER_BY

    case LCase("SQL_MAX_COLUMNS_IN_SELECT")
      ConstByName=SQL_MAX_COLUMNS_IN_SELECT

    case LCase("SQL_MAX_COLUMNS_IN_TABLE")
      ConstByName=SQL_MAX_COLUMNS_IN_TABLE

    case LCase("SQL_MAX_INDEX_SIZE")
      ConstByName=SQL_MAX_INDEX_SIZE

    case LCase("SQL_MAX_ROW_SIZE_INCLUDES_LONG")
      ConstByName=SQL_MAX_ROW_SIZE_INCLUDES_LONG

    case LCase("SQL_MAX_ROW_SIZE")
      ConstByName=SQL_MAX_ROW_SIZE

    case LCase("SQL_MAX_STATEMENT_LEN")
      ConstByName=SQL_MAX_STATEMENT_LEN

    case LCase("SQL_MAX_TABLES_IN_SELECT")
      ConstByName=SQL_MAX_TABLES_IN_SELECT

    case LCase("SQL_MAX_USER_NAME_LEN")
      ConstByName=SQL_MAX_USER_NAME_LEN

    case LCase("SQL_MAX_CHAR_LITERAL_LEN")
      ConstByName=SQL_MAX_CHAR_LITERAL_LEN

    case LCase("SQL_TIMEDATE_ADD_INTERVALS")
      ConstByName=SQL_TIMEDATE_ADD_INTERVALS

    case LCase("SQL_TIMEDATE_DIFF_INTERVALS")
      ConstByName=SQL_TIMEDATE_DIFF_INTERVALS

    case LCase("SQL_NEED_LONG_DATA_LEN")
      ConstByName=SQL_NEED_LONG_DATA_LEN

    case LCase("SQL_MAX_BINARY_LITERAL_LEN")
      ConstByName=SQL_MAX_BINARY_LITERAL_LEN

    case LCase("SQL_LIKE_ESCAPE_CLAUSE")
      ConstByName=SQL_LIKE_ESCAPE_CLAUSE

    case LCase("SQL_QUALIFIER_LOCATION")
      ConstByName=SQL_QUALIFIER_LOCATION

    case LCase("SQL_INFO_LAST")
      ConstByName=SQL_INFO_LAST

    case LCase("SQL_INFO_DRIVER_START")
      ConstByName=SQL_INFO_DRIVER_START

    case LCase("regNull")
      ConstByName=regNull

    case LCase("regString")
      ConstByName=regString

    case LCase("regXString")
      ConstByName=regXString

    case LCase("regBinary")
      ConstByName=regBinary

    case LCase("regDWord")
      ConstByName=regDWord

    case LCase("regLink")
      ConstByName=regLink

    case LCase("regMultiString")
      ConstByName=regMultiString

    case LCase("regResList")
      ConstByName=regResList

    case LCase("HKEY_CLASSES_ROOT")
      ConstByName=HKEY_CLASSES_ROOT

    case LCase("HKEY_CURRENT_USER")
      ConstByName=HKEY_CURRENT_USER

    case LCase("HKEY_LOCAL_MACHINE")
      ConstByName=HKEY_LOCAL_MACHINE

    case LCase("HKEY_USERS")
      ConstByName=HKEY_USERS

    case LCase("HKEY_PERFORMANCE_DATA")
      ConstByName=HKEY_PERFORMANCE_DATA

    case LCase("HKEY_CURRENT_CONFIG")
      ConstByName=HKEY_CURRENT_CONFIG

    case LCase("HKEY_DYN_DATA")
      ConstByName=HKEY_DYN_DATA

    case LCase("scUserAgent")
      ConstByName=scUserAgent

    case LCase("INTERNET_OPEN_TYPE_PRECONFIG")
      ConstByName=INTERNET_OPEN_TYPE_PRECONFIG

    case LCase("INTERNET_DEFAULT_FTP_PORT")
      ConstByName=INTERNET_DEFAULT_FTP_PORT

    case LCase("INTERNET_DEFAULT_GOPHER_PORT")
      ConstByName=INTERNET_DEFAULT_GOPHER_PORT

    case LCase("INTERNET_DEFAULT_HTTP_PORT")
      ConstByName=INTERNET_DEFAULT_HTTP_PORT

    case LCase("INTERNET_DEFAULT_HTTPS_PORT")
      ConstByName=INTERNET_DEFAULT_HTTPS_PORT

    case LCase("INTERNET_DEFAULT_SOCKS_PORT")
      ConstByName=INTERNET_DEFAULT_SOCKS_PORT

    case LCase("INTERNET_SERVICE_FTP")
      ConstByName=INTERNET_SERVICE_FTP

    case LCase("INTERNET_SERVICE_GOPHER")
      ConstByName=INTERNET_SERVICE_GOPHER

    case LCase("INTERNET_SERVICE_HTTP")
      ConstByName=INTERNET_SERVICE_HTTP

    case LCase("INTERNET_FLAG_RELOAD")
      ConstByName=INTERNET_FLAG_RELOAD

    case LCase("HTTP_QUERY_CONTENT_TYPE")
      ConstByName=HTTP_QUERY_CONTENT_TYPE

    case LCase("HTTP_QUERY_CONTENT_LENGTH")
      ConstByName=HTTP_QUERY_CONTENT_LENGTH

    case LCase("HTTP_QUERY_EXPIRES")
      ConstByName=HTTP_QUERY_EXPIRES

    case LCase("HTTP_QUERY_LAST_MODIFIED")
      ConstByName=HTTP_QUERY_LAST_MODIFIED

    case LCase("HTTP_QUERY_PRAGMA")
      ConstByName=HTTP_QUERY_PRAGMA

    case LCase("HTTP_QUERY_VERSION")
      ConstByName=HTTP_QUERY_VERSION

    case LCase("HTTP_QUERY_STATUS_CODE")
      ConstByName=HTTP_QUERY_STATUS_CODE

    case LCase("HTTP_QUERY_STATUS_TEXT")
      ConstByName=HTTP_QUERY_STATUS_TEXT

    case LCase("HTTP_QUERY_RAW_HEADERS")
      ConstByName=HTTP_QUERY_RAW_HEADERS

    case LCase("HTTP_QUERY_RAW_HEADERS_CRLF")
      ConstByName=HTTP_QUERY_RAW_HEADERS_CRLF

    case LCase("HTTP_QUERY_FORWARDED")
      ConstByName=HTTP_QUERY_FORWARDED

    case LCase("HTTP_QUERY_SERVER")
      ConstByName=HTTP_QUERY_SERVER

    case LCase("HTTP_QUERY_USER_AGENT")
      ConstByName=HTTP_QUERY_USER_AGENT

    case LCase("HTTP_QUERY_SET_COOKIE")
      ConstByName=HTTP_QUERY_SET_COOKIE

    case LCase("HTTP_QUERY_REQUEST_METHOD")
      ConstByName=HTTP_QUERY_REQUEST_METHOD

    case LCase("HTTP_QUERY_FLAG_REQUEST_HEADERS")
      ConstByName=HTTP_QUERY_FLAG_REQUEST_HEADERS

    case LCase("INTERNET_OPTION_VERSION")
      ConstByName=INTERNET_OPTION_VERSION

    case LCase("HTTP_ADDREQ_FLAG_ADD_IF_NEW")
      ConstByName=HTTP_ADDREQ_FLAG_ADD_IF_NEW

    case LCase("HTTP_ADDREQ_FLAG_ADD")
      ConstByName=HTTP_ADDREQ_FLAG_ADD

    case LCase("HTTP_ADDREQ_FLAG_REPLACE")
      ConstByName=HTTP_ADDREQ_FLAG_REPLACE

    case LCase("LVM_FIRST")
      ConstByName=LVM_FIRST

    case LCase("LVM_SETCOLUMNWIDTH")
      ConstByName=LVM_SETCOLUMNWIDTH

    case LCase("LVSCW_AUTOSIZE")
      ConstByName=LVSCW_AUTOSIZE

    case LCase("LVSCW_AUTOSIZE_USEHEADER")
      ConstByName=LVSCW_AUTOSIZE_USEHEADER

    case LCase("VER_PLATFORM_WIN32s")
      ConstByName=VER_PLATFORM_WIN32s

    case LCase("VER_PLATFORM_WIN32_WINDOWS")
      ConstByName=VER_PLATFORM_WIN32_WINDOWS

    case LCase("VER_PLATFORM_WIN32_NT")
      ConstByName=VER_PLATFORM_WIN32_NT

    case LCase("SUPPORT_EMAIL")
      ConstByName=SUPPORT_EMAIL

    case LCase("QUOTE")
      ConstByName=QUOTE

    case LCase("DEFAULT_AUTOHIGHLIGHT")
      ConstByName=DEFAULT_AUTOHIGHLIGHT

  End Select
End Function


