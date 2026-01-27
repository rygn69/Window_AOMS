VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SigAccom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   7905
   ClientLeft      =   750
   ClientTop       =   1335
   ClientWidth     =   11940
   Icon            =   "frm_SigAccom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11940
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   393216
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   169869315
      CurrentDate     =   40393
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   169869315
      CurrentDate     =   40393
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&View"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_SigAccom.frx":076A
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   10800
      TabIndex        =   8
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_SigAccom.frx":1164
      cBack           =   16777215
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Signatory Accomplishment Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Define the criteria then click the view button."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   4995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frm_SigAccom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Exec [fmis].[dbo].[MPproc_LoadSigAccom] @from = '" & DTPicker1.Value & "',@to = '" & DTPicker2.Value & "'")
Set MSHFlexGrid1.DataSource = rec
    MSHFlexGrid1.TextMatrix(0, 1) = "ID"
    MSHFlexGrid1.TextMatrix(0, 2) = "Name"
    MSHFlexGrid1.TextMatrix(0, 3) = "Audited"
    MSHFlexGrid1.TextMatrix(0, 4) = "Approved"
    
    MSHFlexGrid1.ColWidth(0) = 0
    MSHFlexGrid1.ColWidth(1) = 0
    MSHFlexGrid1.ColWidth(2) = 5000
    MSHFlexGrid1.ColWidth(3) = 2000
    MSHFlexGrid1.ColWidth(4) = 2000
    MSHFlexGrid1.ColWidth(5) = 2000
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
DTPicker2.Value = Now
End Sub
