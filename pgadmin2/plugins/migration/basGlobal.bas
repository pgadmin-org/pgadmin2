Attribute VB_Name = "basGlobal"
' pgAdmin II Migration Wizard
' This code is based on code from the original pgAdmin Project
' Copyright (C) 1998 - 2003, Dave Page & others

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

'Are we already running?
Global bRunning As Boolean

'The global Server object
Global svr As pgServer

'Reference to the pgAdmin Status Bar
Global sb As Variant

'Msg Timer start value.
Global sTimer As Single

'ODBC Handles
Global lEnv As Long
Global lDBC As Long

'The intermediate collection of object
Global colDB As Collection

Global Const QUOTE = """"
