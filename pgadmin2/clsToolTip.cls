VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' clsToolTip.cls - Create custom tooltip

Option Explicit

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Public Sub AssignToolTip(ByRef hTarget As Object, ByRef sMessage As String)
    Dim TipWindow As Long
    Dim ti As TOOLINFO
    Dim uid As Long
    Dim ToolTipText As String * 50
    Dim RECT As RECT
    uid = 0
    
    TipWindow = CreateWindowEx(0&, TOOLTIPS_CLASSA, "", TTS_ALWAYSTIP, 0, 0, 0, 0, hTarget.hwnd, 0&, App.hInstance, 0&)

    SetWindowPos TipWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    GetClientRect hTarget.hwnd, RECT

    'Fill the TOOLINFO structure with info about
    'the target tool
    With ti
        .cbSize = Len(ti)
        .uFlags = TTF_SUBCLASS
        .hwnd = hTarget.hwnd
        .hinst = App.hInstance
        .uid = uid
        .lpszText = sMessage
        .RECT = RECT
        .lpszText = sMessage
    End With

    SendMessage TipWindow, TTM_ADDTOOLA, 0, ti
    SendMessage TipWindow, TTM_SETMAXTIPWIDTH, 0, 80
End Sub
