Attribute VB_Name = "basRegistry"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basRegistry.bas - Contains Registry manipulation routines.

Option Explicit

Public Function RegRead(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, Optional Default As Variant) As String
On Error Resume Next
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegRead(" & Hive & ", " & QUOTE & Section & QUOTE & ", " & QUOTE & Key & QUOTE & ")", etFullDebug

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
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegWrite(" & Hive & ", " & QUOTE & Section & QUOTE & ", " & QUOTE & Key & QUOTE & ", " & ValType & ", " & QUOTE & Value & QUOTE & ")", etFullDebug

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

Public Function RegGetSubkey(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As String
On Error Resume Next
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegGetSubKey(" & Hive & ", " & QUOTE & Section & QUOTE & ", " & Idx & ")", etFullDebug

Dim lResult As Long
Dim lKeyValue As Long
Dim lDataTypeValue As Long
Dim lValueLength As Long
Dim szValue As String
Dim td As Double
  lResult = RegOpenKey(Hive, Section, lKeyValue)
  szValue = Space(2048)
  lValueLength = Len(szValue)
  lResult = RegEnumKey(lKeyValue, Idx, szValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
    szValue = Left(szValue, InStr(szValue, Chr(0)) - 1)
  Else
    szValue = ""
  End If
  lResult = RegCloseKey(lKeyValue)
  RegGetSubkey = szValue
End Function

Public Function RegReadAll(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As Variant
On Error Resume Next
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegReadAll(" & Hive & ", " & QUOTE & Section & QUOTE & ", " & Idx & ")", etFullDebug

Dim lResult As Long
Dim lKeyValue As Long
Dim lDataTypeValue As Long
Dim lValueLength As Long
Dim lValueNameLength As Long
Dim szValueName As String
Dim szValue As String
Dim td As Double
  lResult = RegOpenKey(Hive, Section, lKeyValue)
  szValue = Space(2048)
  szValueName = Space(2048)
  lValueLength = Len(szValue)
  lValueNameLength = Len(szValueName)
  lResult = RegEnumValue(lKeyValue, Idx, szValueName, lValueNameLength, 0&, lDataTypeValue, szValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
    If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid(szValue, 1, 1)) + &H100& * Asc(Mid(szValue, 2, 1)) + &H10000 * Asc(Mid(szValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(szValue, 4, 1)))
      szValue = Format(td, "000")
    End If
    szValue = Left(szValue, lValueLength - 1)
    szValueName = Left(szValueName, lValueNameLength)
  Else
    szValue = ""
  End If
  lResult = RegCloseKey(lKeyValue)
  RegReadAll = Array(lDataTypeValue, szValueName, szValue)
End Function

Public Sub RegDelSubkey(ByVal Hive As RegHives, ByVal Section As String)
On Error Resume Next
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegDelSubKey(" & Hive & ", " & QUOTE & Section & QUOTE & ")", etFullDebug

Dim lKeyValue As Long
  RegOpenKeyEx Hive, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue
  RegDeleteKey lKeyValue, Section
  RegCloseKey lKeyValue
End Sub

Public Sub RegDelValue(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String)
On Error Resume Next
frmMain.svr.LogEvent "Entering " & App.Title & ":basRegistry.RegDelValue(" & Hive & ", " & QUOTE & Section & QUOTE & ", " & QUOTE & Key & QUOTE & ")", etFullDebug

Dim lKeyValue As Long
  RegOpenKey Hive, Section, lKeyValue
  RegDeleteValue lKeyValue, Key
  RegCloseKey lKeyValue
End Sub
