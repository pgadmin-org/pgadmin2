VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' clsError.bas - Contains gestion Error

Option Explicit

Public Enum TypeInfoError
  TIE_SYSTEM
  TIE_APPLICATION
  TIE_DATABASE
  TIE_DRIVER_ODBC
  TIE_ERROR
End Enum

Const FileTroubleshooting As String = "Troubleshooting.txt"
Const TagError As String = "<ERROR>"
Dim szDesError As String
Dim lNumError As Long
Dim szRoutineError As String

'Error description
Public Property Let Description(ByVal vData As String)
  szDesError = vData
End Property
Public Property Get Description() As String
Attribute Description.VB_UserMemId = 0
  Description = szDesError
End Property

'Error number
Public Property Let Number(ByVal vData As Long)
  lNumError = vData
End Property
Public Property Get Number() As Long
  Number = lNumError
End Property

'Error routine
Public Property Let Routine(ByVal vData As String)
  szRoutineError = vData
End Property
Public Property Get Routine() As String
  Routine = szRoutineError
End Property

Public Function GetInfo(TIE As TypeInfoError) As String
On Error Resume Next
Dim szTemp As String
Dim szSep As String

  szSep = String(60, "-")
  
  If TIE = TIE_SYSTEM Then
    'System. info
    szTemp = "System Information" & vbCrLf
    szTemp = szTemp & szSep & vbCrLf
    szTemp = szTemp & "Platform: " & vbTab & vbTab & WinName & vbCrLf
    szTemp = szTemp & "Version: " & vbTab & vbTab & WinVer & vbCrLf
    szTemp = szTemp & "Build: " & vbTab & vbTab & WinBuild & " " & WinInfo & vbCrLf
    
  ElseIf TIE = TIE_APPLICATION Then
    'app. info
    szTemp = "Application Information" & vbCrLf
    szTemp = szTemp & szSep & vbCrLf
    
    'PgAdmin2
    szTemp = szTemp & "Name: " & vbTab & vbTab & App.Title & vbCrLf
    szTemp = szTemp & "Version: " & vbTab & vbTab & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    szTemp = szTemp & "Descripton: " & vbTab & App.Comments & vbCrLf
    
    'PgSchema
    szTemp = szTemp & szSep & vbCrLf
    szTemp = szTemp & "Name: " & vbTab & vbTab & "pgSchema" & vbCrLf
    szTemp = szTemp & "Version: " & vbTab & vbTab & frmMain.svr.Version.Major & "." & frmMain.svr.Version.Minor & "." & frmMain.svr.Version.Revision & vbCrLf
    szTemp = szTemp & "Descripton: " & vbTab & frmMain.svr.Version.Description & vbCrLf
    
  ElseIf TIE = TIE_DATABASE Then
    'Database
    szTemp = "Database Information" & vbCrLf
    szTemp = szTemp & szSep & vbCrLf
    szTemp = szTemp & "Version: " & vbTab & vbTab & frmMain.svr.dbVersion.Major & "." & frmMain.svr.dbVersion.Minor & "." & frmMain.svr.dbVersion.Revision & vbCrLf
    szTemp = szTemp & "Descripton: " & vbTab & frmMain.svr.dbVersion.Description & vbCrLf
    
  ElseIf TIE = TIE_DRIVER_ODBC Then
    'DriverVersion
    szTemp = "Driver Information" & vbCrLf
    szTemp = szTemp & szSep & vbCrLf
    szTemp = szTemp & "Name: " & vbTab & vbTab & frmMain.svr.DriverName & vbCrLf
    szTemp = szTemp & "Version: " & vbTab & vbTab & frmMain.svr.DriverVersion.Major & "." & frmMain.svr.DriverVersion.Minor & "." & frmMain.svr.DriverVersion.Revision & vbCrLf
    szTemp = szTemp & "Descripton: " & vbTab & frmMain.svr.DriverVersion.Description & vbCrLf
    
  ElseIf TIE = TIE_ERROR Then
    'Error. info
    szTemp = "Error Information" & vbCrLf
    szTemp = szTemp & szSep & vbCrLf
    szTemp = szTemp & "Description: " & vbTab & szDesError & vbCrLf
    szTemp = szTemp & "Number: " & vbTab & vbTab & lNumError & vbCrLf
    szTemp = szTemp & "Routine: " & vbTab & vbTab & szRoutineError & vbCrLf
  End If
  GetInfo = szTemp
End Function

Private Function WinVer() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!

Dim osVersion As OSVERSIONINFO

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  WinVer = osVersion.dwMajorVersion & "." & osVersion.dwMinorVersion

End Function

Private Function WinBuild() As String
On Error Resume Next
'No logging or error handling here because these are called by the error handle
'and we all know how recursive function calls are a Bad Thing!

Dim osVersion As OSVERSIONINFO

  osVersion.dwOSVersionInfoSize = Len(osVersion)
  GetVersionEx osVersion
  WinBuild = osVersion.dwBuildNumber And &HFFFF&
  
End Function

Private Function WinName() As String
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

Private Function WinInfo() As String
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

'retun string possible causes of this error
Public Function Troubleshooting() As String
Dim szTemp As String
Dim vData
Dim lfile As Long
Dim ii As Integer
Dim szError As String
Dim bEnd As Boolean

  szTemp = ""
  If Dir(App.Path & "\" & FileTroubleshooting) <> "" Then
    lfile = FreeFile
    Open App.Path & "\" & FileTroubleshooting For Input As #lfile
    szTemp = Input(LOF(lfile), #lfile)
    Close #lfile
    vData = Split(szTemp, vbCrLf)
  
    'find error in file
    szTemp = ""
    For ii = 0 To UBound(vData)
      If Left(vData(ii), Len(TagError)) = TagError Then
        szError = Mid(vData(ii), Len(TagError) + 1)
        If Mid(szDesError, 1, Len(szError)) = szError Then
          'get description
          szTemp = vbCrLf & vbCrLf
          ii = ii + 1
          bEnd = True
          While ii <= UBound(vData) And bEnd
            If Left(vData(ii), Len(TagError)) <> TagError Then
              szTemp = szTemp & vData(ii) & vbCrLf
            Else
              bEnd = False
            End If
            ii = ii + 1
          Wend
          Exit For
        End If
      End If
    Next
  End If
  Troubleshooting = szTemp
End Function
