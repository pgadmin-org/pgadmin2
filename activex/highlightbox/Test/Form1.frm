VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin HighlightBox.TBX TBX1 
      Height          =   6180
      Left            =   4185
      TabIndex        =   3
      Top             =   0
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   10901
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test HBX"
      Height          =   525
      Left            =   45
      TabIndex        =   1
      Top             =   6255
      Width           =   4095
   End
   Begin VB.CommandButton command1 
      Caption         =   "Test TXT"
      Height          =   510
      Left            =   4230
      TabIndex        =   0
      Top             =   6255
      Width           =   4215
   End
   Begin HighlightBox.HBX HBX1 
      Height          =   6180
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   10901
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RightMargin     =   10000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Start As Single
Dim X As Long
  TBX1.Text = ""
  Start = Timer
  For X = 0 To 10
    TBX1.Text = TBX1.Text & "SELECT \n * FROM pg_class WHERE relname NOT LIKE 'pg_%'" & vbCrLf
  Next X
  command1.Caption = "Textbox: " & Fix((Timer - Start) * 100) / 100 & " Seconds"
End Sub

Private Sub Command2_Click()
Dim Start As Single
Dim X As Long
  HBX1.Font.Name = "Comic Sans MS"
  HBX1.Text = ""
  Start = Timer
  HBX1.AutoColour = False
  For X = 0 To 10
    HBX1.Text = HBX1.Text & "SELECT \n * FROM pg_class WHERE relname NOT LIKE 'pg_%'" & vbCrLf
  Next X
  'HBX1.AutoColour = True
  HBX1.ColourText
  Command2.Caption = "Highlightbox: " & Fix((Timer - Start) * 100) / 100 & " Seconds"
End Sub

Private Sub Form_Load()
  HBX1.Wordlist = "ALTER|0|0|16711680;COMMENT|0|0|16711680;CREATE|0|0|16711680;DELETE|0|0|16711680;DROP|0|0|16711680;EXPLAIN|0|0|16711680;GRANT|0|0|16711680;INSERT|0|0|16711680;REVOKE|0|0|16711680;" & _
                  "SELECT|0|0|16711680;UPDATE|0|0|16711680;VACUUM|0|0|16711680;AGGREGATE|0|0|255;CONSTRAINT|0|0|255;DATABASE|0|0|255;FUNCTION|0|0|255;GROUP|0|0|255;INDEX|0|0|255;" & _
                  "LANGUAGE|0|0|255;OPERATOR|0|0|255;RULE|0|0|255;SEQUENCE|0|0|255;TABLE|0|0|255;TRIGGER|0|0|255;ABORT|0|0|11998061;BEGIN|0|0|11998061;" & _
                  "CHECKPOINT|0|0|11998061;CLOSE|0|0|11998061;CLUSTER|0|0|11998061;COMMIT|0|0|11998061;COPY|0|0|11998061;DECLARE|0|0|11998061;FETCH|0|0|11998061;LISTEN|0|0|11998061;" & _
                  "LOAD|0|0|11998061;LOCK|0|0|11998061;MOVE|0|0|11998061;NOTIFY|0|0|11998061;REINDEX|0|0|11998061;RESET|0|0|11998061;ROLLBACK|0|0|11998061;SET|0|0|11998061;SHOW|0|0|11998061;TRUNCATE|0|0|11998061;" & _
                  "UNLISTEN|0|0|11998061;AS|0|0|32768;ASC|0|0|32768;ASCENDING|0|0|32768;BY|0|0|32768;CASE|0|0|32768;DESC|0|0|32768;DESCENDING|0|0|32768;ELSE|0|0|32768;FROM|0|0|32768;END|0|0|32768;HAVING|0|0|32768;INTO|0|0|32768;" & _
                  "ON|0|0|32768;ORDER|0|0|32768;THEN|0|0|32768;USING|0|0|32768;WHEN|0|0|32768;WHERE|0|0|32768;"
End Sub

