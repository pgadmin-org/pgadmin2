VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgAdmin2_Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Note:
' pgAdmin Plugins must
' a) Have 'pgAdmin2_Plugin' (without the quotes) somewhere in the class name
' b) Reference 'pgAdmin II Abstract Plugin', 'PostgreSQL Schema Objects' & 'Microsoft ActiveX Data Objects 2.0'

'Implement the Plugin Interface
Implements pgPlugin

'The pgPlugin_Author function returns info about the Plugin's Author
Private Property Get pgPlugin_Author() As String
  pgPlugin_Author = "Written by Dave Page" & vbCrLf & _
                      "http://www.pgadmin.org/"
End Property

'The pgPlugin_Description function simply returns the description that is used
'in menus etc.
Private Property Get pgPlugin_Description() As String
  pgPlugin_Description = "MSysConf Wizard"
End Property

'The pgPlugin_Version function simply returns the Plugin Version
'This code should never need to be changed.
Private Property Get pgPlugin_Version() As String
  pgPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

'The Plugin type can currently be either 1 or 2.
'A type 1 plugin will only be shown by pgAdmin when logged on to a server.
'A type 2 plugin is always available to the user in pgAdmin
Private Property Get pgPlugin_PluginType() As Long
  pgPlugin_PluginType = 1
End Property

'The pgPlugin_Export subroutine is called to do the actual job. The sub is
'passed a reference to the main pgSchema object.
Private Sub pgPlugin_Execute(objStatusBar As Variant, objSchema As Variant)

  'Abort if we're already running
  If bRunning = True Then Exit Sub
  bRunning = True
  
  'Reference the Server/Status Bar object
  Set svr = objSchema
  Set sb = objStatusBar

  'Display the Wizard
  Load frmWizard
  frmWizard.Initialise
  frmWizard.Show
  
End Sub


