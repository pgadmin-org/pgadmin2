VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmSQLExplain 
   Caption         =   "Query Plan"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmSQLExplain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5670
   Begin HighlightBox.HBX txtQuery 
      Height          =   1860
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Displays the SQL Query."
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3281
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "SQL Query"
   End
   Begin HighlightBox.HBX txtPlan 
      Height          =   1860
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Displays the Query Execution Plan."
      Top             =   1890
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3281
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "Query Execution Plan"
   End
End
Attribute VB_Name = "frmSQLExplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmSQLExplain.frm - Display an SQL Query Plan

Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLExplain.Form_Load()", etFullDebug

  Me.Width = 5790
  Me.Height = 4200
  
  Set txtQuery.Font = ctx.Font
  Set txtPlan.Font = ctx.Font
  txtQuery.Wordlist = ctx.AutoHighlight
  txtPlan.Wordlist = ctx.AutoHighlight
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLExplain.Form_Resize()", etFullDebug

  txtQuery.Minimise
  txtPlan.Minimise
  If Me.WindowState <> 1 Then
    If Me.Width < 5790 Then Me.Width = 5790
    If Me.Height < 4200 Then Me.Height = 4200
    txtQuery.Width = Me.ScaleWidth
    txtPlan.Width = Me.ScaleWidth
    txtQuery.Height = (Me.ScaleHeight / 5) * 2
    txtPlan.Height = ((Me.ScaleHeight / 5) * 3) - 50
    txtPlan.Top = txtQuery.Height + 50
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_Resize"
End Sub

Public Sub Explain(szSQL As String, szDatabase As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLExplain.Form_Resize()", etFullDebug

Dim lEnv As Long
Dim lDBC As Long
Dim lRet As Long
Dim lStmt As Long
Dim lErr As Long
Dim iSize As Integer
Dim szConnect As String
Dim szResult As String * 256
Dim szSqlState As String * 1024
Dim szErrorMsg As String * 1024
Dim szPlan As String
Dim rsPlan As New Recordset

  Me.Caption = "Query Plan (Database: " & szDatabase & ")"
  txtQuery.Text = szSQL
  txtQuery.ColourText

  'Execute the statement. In theory, the ADO connection object can does this, and the plan can
  'be picked up as a series of 512Byte strings in the Errors collection. This is unreliable though, so
  'we'll use ODBC directly insted <shudder>
  
  StartMsg "Requesting Query Execution Plan..."
  
  'Query plans are returned as resultsets in 7.3+
  If ctx.dbVer >= 7.3 Then
    Set rsPlan = frmMain.svr.Databases(szDatabase).Execute("EXPLAIN " & szSQL)
    While Not rsPlan.EOF
      txtPlan.Text = txtPlan.Text & rsPlan.Fields(0).Value & vbCrLf
      rsPlan.MoveNext
    Wend
    txtPlan.ColourText
    If rsPlan.State <> adStateClosed Then rsPlan.Close
    Set rsPlan = Nothing
    
    If txtPlan.Text = "" Then
      frmMain.svr.LogEvent "A Query Execution Plan could not be calculated for the specified SQL query.", etMiniDebug
      txtPlan.Text = "A Query Execution Plan could not be calculated for the specified SQL query."
    End If
  
  Else
    'Initialisze the ODBC subsystem
    If SQLAllocEnv(lEnv) <> 0 Then
      frmMain.svr.LogEvent "Unable to initialize ODBC API drivers!", etMiniDebug
      MsgBox "Unable to initialize ODBC API drivers!", vbCritical, "Error"
      GoTo Cleanup
    End If
  
    If SQLAllocConnect(lEnv, lDBC) <> 0 Then
      frmMain.svr.LogEvent "Could not allocate memory for connection Handle!", etMiniDebug
      MsgBox "Could not allocate memory for connection Handle!", vbCritical, "Error"
      GoTo Cleanup
    End If
  
    szConnect = "DRIVER=" & frmMain.svr.DriverName & ";DATABASE=" & szDatabase & ";UID=" & ctx.Username & ";PWD=" & ctx.Password & ";SERVER=" & frmMain.svr.Server & ";PORT=" & frmMain.svr.Port
    lRet = SQLDriverConnect(lDBC, Me.hWnd, szConnect, Len(szConnect), szResult, Len(szResult), iSize, 1)
    If lRet <> SQL_SUCCESS Then
      frmMain.svr.LogEvent "Could not establish connection to ODBC driver! Error: " & lRet, etMiniDebug
      MsgBox "Could not establish connection to ODBC driver!" & vbCrLf & "Error: " & lRet, vbCritical, "Error"
      GoTo Cleanup
    End If
    
    'Check the ODBC Driver version. EXPLAIN will only work with 07.01.0006 or higher.
    SQLGetInfoString lDBC, SQL_DBMS_VER, szResult, Len(szResult), vbNull
    frmMain.svr.LogEvent "ODBC Driver Version: " & szResult, etMiniDebug
    If Val(Mid(szResult, 1, 2)) < 7 Then
       frmMain.svr.LogEvent "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", etMiniDebug
       MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
       GoTo Cleanup
    Else
      If Val(Mid(szResult, 1, 2)) = 7 Then
        If Val(Mid(szResult, 4, 2)) < 1 Then
          frmMain.svr.LogEvent "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", etMiniDebug
          MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
          GoTo Cleanup
        Else
          If Val(Mid(szResult, 4, 2)) = 1 Then
            If Val(Mid(szResult, 7, 4)) < 6 Then
              frmMain.svr.LogEvent "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", etMiniDebug
              MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
              GoTo Cleanup
            End If
          End If
        End If
      End If
    End If
    
    'Allocate memory for the statement handle.
    If SQLAllocStmt(lDBC, lStmt) <> 0 Then
      frmMain.svr.LogEvent "Could not allocate memory for a statement handle!", etMiniDebug
      MsgBox "Could not allocate memory for a statement handle!", vbCritical, "Error"
      Exit Sub
    End If
    
    szSQL = "EXPLAIN " & szSQL
    frmMain.svr.LogEvent "SQLExecDirect: " & szSQL, etMiniDebug
    If SQLExecDirect(lStmt, szSQL, Len(szSQL)) = SQL_SUCCESS_WITH_INFO Then
      While SQLError(lEnv, lDBC, lStmt, szSqlState, lErr, szErrorMsg, 1024, iSize) <> SQL_NO_DATA_FOUND
        If iSize > 512 Then iSize = 512
        szPlan = szPlan & Left(szErrorMsg, iSize)
      Wend
    End If
    
    If Len(szPlan) > 22 Then szPlan = Mid(szPlan, 23)
    If szPlan <> "" Then
      txtPlan.Text = szPlan
      txtPlan.ColourText
    Else
      frmMain.svr.LogEvent "A Query Execution Plan could not be calculated for the specified SQL query.", etMiniDebug
      txtPlan.Text = "A Query Execution Plan could not be calculated for the specified SQL query."
    End If
  
Cleanup:
    'Log out and cleanup
    If lDBC <> 0 Then
      SQLDisconnect lDBC
    End If
    SQLFreeConnect lDBC
    If lEnv <> 0 Then
      SQLFreeEnv lEnv
    End If
  End If
 
  EndMsg
  Exit Sub
Err_Handler:
  If Err.Number = -2147467259 Then 'Query cannot be EXPLAINed or is invalid
    frmMain.svr.LogEvent "A Query Execution Plan could not be calculated for the specified SQL query.", etMiniDebug
    txtPlan.Text = "A Query Execution Plan could not be calculated for the specified SQL query."
    EndMsg
    If rsPlan.State <> adStateClosed Then rsPlan.Close
    Set rsPlan = Nothing
    Exit Sub
  End If
  If rsPlan.State <> adStateClosed Then rsPlan.Close
  Set rsPlan = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Explain"
End Sub
