VERSION 5.00
Begin VB.Form frmDummy 
   Caption         =   "Dummy Form"
   ClientHeight    =   2628
   ClientLeft      =   6360
   ClientTop       =   5652
   ClientWidth     =   4368
   LinkTopic       =   "Form1"
   ScaleHeight     =   2628
   ScaleWidth      =   4368
   Begin pgAdmin2.ScrollObjDb ScrollObjDb1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   445
      _ExtentY        =   656
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
