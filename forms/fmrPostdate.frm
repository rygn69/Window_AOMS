VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPOstdate 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Date Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin lvButton.lvButtons_H Command1 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         Caption         =   "&OK"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483635
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "fmrPostdate.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use Default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   121896963
         UpDown          =   -1  'True
         CurrentDate     =   40596
      End
      Begin lvButton.lvButtons_H Command2 
         Height          =   735
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483635
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "fmrPostdate.frx":1052
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frmPOstdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DatePost = Format(DTPicker1.Value, "MM/DD/YYYY")
If Check1.Value = 1 Then
    DefaultPost = DTPicker1.Value
End If
JevOk = True
Unload Me
End Sub

'Private Sub Command2_Click()
'If MsgBox("Are You Sure You Want To Cancel Your Posting? ", vbYesNo, "System Confirmation") = vbYes Then
'JevOk = False
'Unload Me
'End If
'End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If IsEmpty(DefaultPost) = True Then
DTPicker1.Value = Now
Else
DTPicker1.Value = DefaultPost
End If
Check1.Value = 1
End Sub

Private Sub lvButtons_H1_Click()

End Sub
