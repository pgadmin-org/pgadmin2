VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' clsPlugins.cls - Manages the Plugins Collection

Option Explicit
Dim mCol As Collection

Private Sub Class_Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Class_Initialize()", etFullDebug

  LoadPlugins
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Class_Initialize"
End Sub

Public Sub LoadPlugins()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.LoadPlugins()", etFullDebug

Dim szPlugin As String
Dim objPlugin As pgPlugin
Dim X As Long

  StartMsg "Loading Plugins..."
  Set mCol = New Collection
  szPlugin = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Do Until szPlugin = ""
    If InStr(1, szPlugin, "pgAdmin2_Plugin") Then
      Set objPlugin = CreateObject(szPlugin)
      frmMain.svr.LogEvent "Loading Plugin: " & szPlugin & " (" & objPlugin.Description & " v" & objPlugin.Version & ")", etMiniDebug
      mCol.Add objPlugin, objPlugin.Description
Continue:
    End If
    X = X + 1
    szPlugin = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Loop
  
  EndMsg
  Exit Sub
  
Err_Handler:
  If Err.Number = -2147024770 Then
    frmMain.svr.LogEvent "Plugin: " & szPlugin & " is registered but could not be found!", etErrors
    GoTo Continue
  ElseIf Err.Number = 13 Or Err.Number = 429 Then
    frmMain.svr.LogEvent "Plugin: " & szPlugin & " is corrupt or invalid!", etErrors
    GoTo Continue
  Else
    EndMsg
    If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.LoadPlugins"
  End If
End Sub

'Return the specified Item
Public Property Get Item(IndexOrKey As Variant) As pgPlugin
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Property Get Item(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

  Set Item = mCol(IndexOrKey)
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Property Get Item"
End Property

'Return the Item count
Public Property Get Count() As Long
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Property Get Count"
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Property Get NewEnum"
End Property

Public Sub Install(szFilename As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Install(" & QUOTE & szFilename & QUOTE & ")", etFullDebug

Dim lHdc As Long

  StartMsg "Installing Plugin..."
  frmMain.svr.LogEvent "Installing: " & szFilename, etMiniDebug
  lHdc = GetDesktopWindow()
  ShellExecute lHdc, "Open", "regsvr32", " /s " & QUOTE & szFilename & QUOTE, "C:\", SW_SHOWNORMAL
  LoadPlugins
  
  EndMsg
  Exit Sub
    
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Install"
End Sub

Public Sub Uninstall(szPlugin As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsPlugins.Uninstall(" & QUOTE & szPlugin & QUOTE & ")", etFullDebug

Dim lHdc As Long
Dim szKey As String
Dim szDll As String
Dim szClassID As String
Dim objPlugin As pgPlugin
Dim X As Long

  StartMsg "Uninstalling Plugin..."
  frmMain.svr.LogEvent "Uninstalling: " & szPlugin, etMiniDebug
  
  szKey = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Do Until szKey = ""
    If InStr(1, szKey, "pgAdmin2_Plugin") Then
      Set objPlugin = CreateObject(szKey)
      If objPlugin.Description = szPlugin Then
        szClassID = RegRead(HKEY_CLASSES_ROOT, szKey & "\Clsid", "")
        szDll = RegRead(HKEY_CLASSES_ROOT, "CLSID\" & szClassID & "\InProcServer32", "")
        lHdc = GetDesktopWindow()
        ShellExecute lHdc, "Open", "regsvr32", " /s /u " & QUOTE & szDll & QUOTE, "C:\", SW_SHOWNORMAL
        mCol.Remove szPlugin
      End If
    End If
  X = X + 1
  szKey = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Loop
  
  EndMsg
  Exit Sub
    
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Uninstall"
End Sub
