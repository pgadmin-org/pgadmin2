VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExporters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' clsExporters.cls - Manages the Exporters Collection

Option Explicit
Dim mCol As Collection

Private Sub Class_Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Class_Initialize()", etFullDebug

  LoadExporters
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Class_Initialize"
End Sub

Public Sub LoadExporters()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.LoadExporters()", etFullDebug

Dim szExporter As String
Dim objExporter As pgExporter
Dim X As Long

  StartMsg §§TrasLang§§("Loading Exporters...")
  Set mCol = New Collection
  szExporter = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Do Until szExporter = ""
    If InStr(1, szExporter, "pgAdmin2_Exporter") Then
      Set objExporter = CreateObject(szExporter)
      frmMain.svr.LogEvent "Loading Exporter: " & szExporter & " (" & objExporter.Description & " v" & objExporter.Version & ")", etMiniDebug
      mCol.Add objExporter, objExporter.Description
Continue:
    End If
    X = X + 1
    szExporter = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Loop
  
  EndMsg
  Exit Sub
  
Err_Handler:
  If Err.Number = -2147024770 Then
    frmMain.svr.LogEvent "Exporter: " & szExporter & " is registered but could not be found!", etErrors
    GoTo Continue
  ElseIf Err.Number = 13 Or Err.Number = 429 Then
    frmMain.svr.LogEvent "Exporter: " & szExporter & " is corrupt or invalid!", etErrors
    GoTo Continue
  Else
    EndMsg
    If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.LoadExporters"
  End If
End Sub

'Return the specified Item
Public Property Get Item(IndexOrKey As Variant) As pgExporter
Attribute Item.VB_UserMemId = 0
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Property Get Item(" & QUOTE & IndexOrKey & QUOTE & ")", etFullDebug

  Set Item = mCol(IndexOrKey)
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Property Get Item"
End Property

'Return the Item count
Public Property Get Count() As Long
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Property Get Count"
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Property Get NewEnum"
End Property

Public Sub Install(szFilename As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Install(" & QUOTE & szFilename & QUOTE & ")", etFullDebug

Dim lHdc As Long

  StartMsg §§TrasLang§§("Installing Exporter...")
  frmMain.svr.LogEvent "Installing: " & szFilename, etMiniDebug
  lHdc = GetDesktopWindow()
  ShellExecute lHdc, "Open", "regsvr32", " /s " & QUOTE & szFilename & QUOTE, "C:\", SW_SHOWNORMAL
  LoadExporters
  
  EndMsg
  Exit Sub
    
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Install"
End Sub

Public Sub Uninstall(szExporter As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":clsExporters.Uninstall(" & QUOTE & szExporter & QUOTE & ")", etFullDebug

Dim lHdc As Long
Dim szKey As String
Dim szDll As String
Dim szClassID As String
Dim objExporter As pgExporter
Dim X As Long

  StartMsg §§TrasLang§§("Uninstalling Exporter...")
  frmMain.svr.LogEvent "Uninstalling: " & szExporter, etMiniDebug
  
  szKey = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Do Until szKey = ""
    If InStr(1, szKey, "pgAdmin2_Exporter") Then
      Set objExporter = CreateObject(szKey)
      If objExporter.Description = szExporter Then
        szClassID = RegRead(HKEY_CLASSES_ROOT, szKey & "\Clsid", "")
        szDll = RegRead(HKEY_CLASSES_ROOT, "CLSID\" & szClassID & "\InProcServer32", "")
        lHdc = GetDesktopWindow()
        ShellExecute lHdc, "Open", "regsvr32", " /s /u " & QUOTE & szDll & QUOTE, "C:\", SW_SHOWNORMAL
        mCol.Remove szExporter
      End If
    End If
  X = X + 1
  szKey = RegGetSubkey(HKEY_CLASSES_ROOT, "", X)
  Loop
  
  EndMsg
  Exit Sub
    
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsExporters.Uninstall"
End Sub
