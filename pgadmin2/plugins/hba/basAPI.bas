Attribute VB_Name = "basAPI"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Constants
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2 'Note: On last column, its width fills remaining width
                                                   'of list-view according to Micro$oft. This does not
                                                   'appear to be the case when I do it.
