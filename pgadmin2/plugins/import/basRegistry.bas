Attribute VB_Name = "basRegistry"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basRegistry.bas - Contains Registry manipulation routines.

Option Explicit

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

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

Public Function RegRead(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, Optional Default As Variant) As String
On Error Resume Next

Dim lResult As Long
Dim lKeyValue As Long
Dim lDataTypeValue As Long
Dim lValueLength As Long
Dim szValue As String
Dim td As Double
Dim TStr1 As String
Dim TStr2 As String
Dim i As Integer
  lResult = RegOpenKey(Hive, Section, lKeyValue)
  szValue = Space(2048)
  lValueLength = Len(szValue)
  lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, szValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
    If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid(szValue, 1, 1)) + &H100& * Asc(Mid(szValue, 2, 1)) + &H10000 * Asc(Mid(szValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(szValue, 4, 1)))
      szValue = Format(td, "000")
    End If
    If lDataTypeValue = REG_BINARY Then
      ' Return a binary field as a hex string (2 chars per byte)
      TStr2 = ""
      For i = 1 To lValueLength
        TStr1 = Hex(Asc(Mid(szValue, i, 1)))
        If Len(TStr1) = 1 Then TStr1 = "0" & TStr1
        TStr2 = TStr2 + TStr1
      Next
      szValue = TStr2
    Else
      szValue = Left(szValue, lValueLength - 1)
    End If
  Else
    szValue = Default
  End If
  lResult = RegCloseKey(lKeyValue)
  RegRead = szValue
End Function

Public Sub RegWrite(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, ByVal ValType As RegTypes, ByVal Value As Variant)
On Error Resume Next

Dim lResult As Long
Dim lKeyValue As Long
Dim InLen As Long
Dim lNewVal As Long
Dim szNewVal As String
  lResult = RegCreateKey(Hive, Section, lKeyValue)
  If ValType = regDWord Then
    lNewVal = CLng(Value)
    InLen = 4
    lResult = RegSetValueExLong(lKeyValue, Key, 0&, ValType, lNewVal, InLen)
  Else
    If ValType = regString Then Value = Value + Chr(0)
    szNewVal = Value
    InLen = Len(szNewVal)
    lResult = RegSetValueExString(lKeyValue, Key, 0&, 1&, szNewVal, InLen)
  End If
  lResult = RegFlushKey(lKeyValue)
  lResult = RegCloseKey(lKeyValue)
End Sub

Public Sub RegDelSubkey(ByVal Hive As RegHives, ByVal Section As String)
On Error Resume Next

Dim lKeyValue As Long
  RegOpenKeyEx Hive, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue
  RegDeleteKey lKeyValue, Section
  RegCloseKey lKeyValue
End Sub
