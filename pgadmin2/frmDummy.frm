VERSION 5.00
Begin VB.Form frmDummy 
   Caption         =   "Dummy Form"
   ClientHeight    =   1104
   ClientLeft      =   5580
   ClientTop       =   3516
   ClientWidth     =   2316
   LinkTopic       =   "Form1"
   ScaleHeight     =   1104
   ScaleWidth      =   2316
   Begin pgAdmin2.ScrollObjDb ScrollObjDb1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _extentx        =   445
      _extenty        =   656
   End
   Begin VB.Image imgChecked 
      Height          =   228
      Left            =   480
      Picture         =   "frmDummy.frx":0000
      Top             =   120
      Width           =   216
   End
   Begin VB.Image imgUnchecked 
      Height          =   216
      Left            =   480
      Picture         =   "frmDummy.frx":046A
      Top             =   372
      Width           =   228
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
