Attribute VB_Name = "basAPI"
' pgAdmin II Migration Wizard
' This code is based on code from the original pgAdmin Project
' Copyright (C) 1998 - 2002, Dave Page & others

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit

Public Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv&, phdbc&) As Integer
Public Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv&) As Integer
Public Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc&, ByVal hWnd As Long, ByVal szCSIn$, ByVal cbCSIn%, ByVal szCSOut$, ByVal cbCSMax%, cbCSOut%, ByVal fDrvrComp%) As Integer
Public Declare Function SQLGetInfoString Lib "odbc32.dll" Alias "SQLGetInfo" (ByVal hdbc&, ByVal fInfoType%, ByVal rgbInfoValue As String, ByVal cbInfoMax%, cbInfoOut%) As Integer
Public Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Public Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv&) As Integer
Public Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer

Public Const SQL_SUCCESS As Long = 0
Public Const SQL_DRIVER_NOPROMPT As Long = 0
Public Const SQL_IDENTIFIER_QUOTE_CHAR As Long = 29
Public Const SQL_FD_FETCH_NEXT As Long = &H1&
