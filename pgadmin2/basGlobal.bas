Attribute VB_Name = "basGlobal"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basGlobal.bas - Contains global declarations and constants.

Option Explicit

'Running Environment
Global inIDE As Boolean

'Support email address
Global Const SUPPORT_EMAIL = "pgadmin-support@postgresql.org"

'Makes life easier...
Global Const Quote = """"

'Global Context object. This contains Globals.
Global ctx As New clsContext

'Global Exporters Class
Global exp As New clsExporters

'Global Plugins Class
Global plg As New clsPlugins

'Msg Timer start value.
Global sTimer As Single

'Default HighLight Colours
Global szDefaultAutoHighlight As String

'error to ignore
Global ColIgnoreError As New Collection

'template file lenguage
Public Const TEMPLATE_FILE_LANG As String = "Language.tmp"

