VERSION 5.00
Begin VB.Form frmTypeMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type Map"
   ClientHeight    =   4440
   ClientLeft      =   1230
   ClientTop       =   1425
   ClientWidth     =   7845
   Icon            =   "frmTypeMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7845
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   33
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   45
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   3330
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   24
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   59
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   450
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   25
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   60
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   810
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   26
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   61
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1170
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   27
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   62
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1530
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   28
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   63
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1890
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   29
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   64
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2250
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   30
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   65
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2610
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   31
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   66
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2970
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   32
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   67
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   3330
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   12
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   36
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   90
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   13
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   37
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   450
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   14
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   38
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   810
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   15
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   39
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1170
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   16
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   40
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1530
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   17
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   41
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1890
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   18
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   42
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2250
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   19
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   43
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2610
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   20
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   44
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2970
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   21
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   46
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   3690
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   22
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   47
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   4050
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   23
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   48
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Exit"
      Height          =   330
      Left            =   6210
      TabIndex        =   68
      ToolTipText     =   "Save changes and Exit."
      Top             =   4050
      Width           =   1545
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   11
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   4050
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   10
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   22
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   3690
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   9
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   21
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   3330
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   8
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   20
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2970
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   7
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   19
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2610
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   6
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   18
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   2250
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   5
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1890
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   4
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1530
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   3
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   1170
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   2
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   14
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   810
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   1
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   450
      Width           =   1200
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   0
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Select the PostgreSQL Datatype to map to."
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numeric"
      Height          =   195
      Index           =   33
      Left            =   2670
      TabIndex        =   58
      Top             =   3375
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TinyInt"
      Height          =   195
      Index           =   24
      Left            =   5265
      TabIndex        =   49
      Top             =   495
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UnsignedBigInt"
      Height          =   195
      Index           =   25
      Left            =   5265
      TabIndex        =   50
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UnsignedSmallInt"
      Height          =   195
      Index           =   26
      Left            =   5265
      TabIndex        =   51
      Top             =   1215
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UnsignedTinyInt"
      Height          =   195
      Index           =   27
      Left            =   5265
      TabIndex        =   52
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UserDefined"
      Height          =   195
      Index           =   28
      Left            =   5265
      TabIndex        =   53
      Top             =   1935
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VarBinary"
      Height          =   195
      Index           =   29
      Left            =   5265
      TabIndex        =   54
      Top             =   2295
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VarChar"
      Height          =   195
      Index           =   30
      Left            =   5265
      TabIndex        =   55
      Top             =   2655
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VarWChar"
      Height          =   195
      Index           =   31
      Left            =   5265
      TabIndex        =   56
      Top             =   3015
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WChar"
      Height          =   195
      Index           =   32
      Left            =   5265
      TabIndex        =   57
      Top             =   3375
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Double"
      Height          =   195
      Index           =   12
      Left            =   2700
      TabIndex        =   24
      Top             =   135
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empty"
      Height          =   195
      Index           =   13
      Left            =   2700
      TabIndex        =   25
      Top             =   495
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Error"
      Height          =   195
      Index           =   14
      Left            =   2700
      TabIndex        =   26
      Top             =   855
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FileTime"
      Height          =   195
      Index           =   15
      Left            =   2700
      TabIndex        =   27
      Top             =   1215
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GUID"
      Height          =   195
      Index           =   16
      Left            =   2700
      TabIndex        =   28
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Integer"
      Height          =   195
      Index           =   17
      Left            =   2700
      TabIndex        =   29
      Top             =   1935
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LongVarBinary"
      Height          =   195
      Index           =   18
      Left            =   2700
      TabIndex        =   30
      Top             =   2295
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LongVarChar"
      Height          =   195
      Index           =   19
      Left            =   2700
      TabIndex        =   31
      Top             =   2655
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LongVarWChar"
      Height          =   195
      Index           =   20
      Left            =   2700
      TabIndex        =   32
      Top             =   3015
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PropVariant"
      Height          =   195
      Index           =   21
      Left            =   2700
      TabIndex        =   33
      Top             =   3735
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Single"
      Height          =   195
      Index           =   22
      Left            =   2700
      TabIndex        =   34
      Top             =   4095
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SmallInt"
      Height          =   195
      Index           =   23
      Left            =   5310
      TabIndex        =   35
      Top             =   135
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Decimal"
      Height          =   195
      Index           =   11
      Left            =   90
      TabIndex        =   11
      Top             =   4095
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DBTimestamp"
      Height          =   195
      Index           =   10
      Left            =   90
      TabIndex        =   10
      Top             =   3735
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DBTime"
      Height          =   195
      Index           =   9
      Left            =   90
      TabIndex        =   9
      Top             =   3375
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DBDate"
      Height          =   195
      Index           =   8
      Left            =   90
      TabIndex        =   8
      Top             =   3015
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   2655
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Currency"
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   6
      Top             =   2295
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Char"
      Height          =   195
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   1935
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chapter"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   1575
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BSTR"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1215
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Boolean"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   855
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Binary"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   495
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BigInt"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   405
   End
End
Attribute VB_Name = "frmTypeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim X As Integer

  StartMsg "Saving Type Map..."
  For X = 0 To 33   ' AM 20020110 Added adNumeric
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Migration Wizard\Type Map", Label1(X).Caption, regString, cboType(X).Text
  Next
  
  EndMsg
  
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTypeMap.cmdSave_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim objType As pgType
Dim X As Integer
Dim Y As Integer
Dim Temp As String
Dim Current As String

  StartMsg "Loading Type Definitions..."

  For X = 0 To 33  ' AM 20020110 Added adNumeric
    cboType(X).Clear
  Next
  For Each objType In svr.Databases(frmWizard.lstDatabase.SelectedItem.Text).Types

    For X = 0 To 33   ' AM 20020110  Added adNumeric
      If Left(objType.Identifier, 1) <> "_" Then cboType(X).AddItem objType.Identifier
    Next
  Next objType

  For X = 0 To 33   ' AM 20020110 Added adNumeric
    Select Case Label1(X).Caption
      Case "BigInt"
        Temp = "int8"
      Case "Binary"
        Temp = "text"
      Case "Boolean"
        Temp = "bool"
      Case "BSTR"
        Temp = "bytea"
      Case "Chapter"
        Temp = "int4"
      Case "Char"
        Temp = "char"
      Case "Currency"
        Temp = "money"
      Case "Date"
        Temp = "date"
      Case "DBDate"
        Temp = "date"
      Case "DBTime"
        Temp = "time"
      Case "DBTimestamp"
        Temp = "timestamp"
      Case "Decimal"
        Temp = "numeric"
      Case "Numeric"       ' AM 20020110  Added adNumeric
        Temp = "numeric"
      Case "Double"
        Temp = "float8"
      Case "Empty"
        Temp = "text"
      Case "Error"
        Temp = "int4"
      Case "FileTime"
        Temp = "timestamp"
      Case "GUID"
        Temp = "text"
      Case "Integer"
        Temp = "int4"
      Case "LongVarBinary"
        Temp = "lo"
      Case "LongVarChar"
        Temp = "text"
      Case "LongVarWChar"
        Temp = "text"
       Case "PropVariant"
        Temp = "text"
       Case "Single"
        Temp = "float4"
       Case "SmallInt"
        Temp = "int2"
       Case "TinyInt"
        Temp = "int2"
       Case "UnsignedBigInt"
        Temp = "int8"
       Case "UnsignedInt"
        Temp = "int4"
       Case "UnsignedSmallInt"
        Temp = "int2"
       Case "UnsignedTinyInt"
        Temp = "int2"
       Case "UserDefined"
        Temp = "text"
       Case "VarBinary"
        Temp = "lo"
       Case "VarChar"
        '1/16/2001 Rod Childers
        'Changed VarChar to default to VarChar
        'Text in Access is = VarChar in PostgreSQL
        'Memo in Access is = text in PostgreSQL
        'Temp = "text"
        Temp = "varchar"
       Case "VarWChar"
          '1/16/2001 Rod Childers
          'Changed VarWChar to default to VarChar
          'Text in Access is = VarChar in PostgreSQL
          'Memo in Access is = text in PostgreSQL
         Temp = "varchar"
       Case "WVar"
        Temp = "text"
    End Select
    Current = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\Migration Wizard\Type Map", Label1(X).Caption, Temp)
    For Y = 0 To cboType(X).ListCount - 1
      If cboType(X).List(Y) = Current Then cboType(X).ListIndex = Y
    Next
  Next
  EndMsg

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTypeMap.Form_Load"
End Sub
