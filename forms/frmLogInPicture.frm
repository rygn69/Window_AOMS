VERSION 5.00
Begin VB.Form frmLogInPicture 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   7080
   ClientTop       =   2505
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   5550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -15
      TabIndex        =   1
      Top             =   6990
      Width           =   5550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   -15
      TabIndex        =   0
      Top             =   6585
      Width           =   5550
   End
   Begin VB.Image img_userpic 
      Height          =   6435
      Left            =   -30
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   5685
   End
End
Attribute VB_Name = "frmLogInPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.ZOrder 0
End Sub
Private Sub Form_Load()
Dim lR As Long
lR = SetTopMostWindow(Me.hwnd, True)
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.BackColor = MDIFrm_MAIN.BackColor
End Sub

