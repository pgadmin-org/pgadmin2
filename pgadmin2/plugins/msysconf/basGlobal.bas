Attribute VB_Name = "basGlobal"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Are we already running?
Global bRunning As Boolean

'The global Server object
Global svr As pgServer

'Reference to the pgAdmin Status Bar
Global sb As StatusBar

'Msg Timer start value.
Global sTimer As Single
