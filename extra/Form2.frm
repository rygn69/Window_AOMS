VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.8#0"; "FLATBTN2.OCX"
Object = "{CEE11B38-2F29-11D3-B64F-444553540000}#2.0#0"; "BlinkingLabel.ocx"
Begin VB.Form frmErr 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin BlinkingLabel.ctlblink ctlblink1 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1085
      Caption         =   "Note: If the error Persist Contact your system            Administrator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      Interval        =   300
   End
   Begin DevPowerFlatBttn.FlatBttn FlatBttn1 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3360
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      AutoSize        =   0   'False
      Caption         =   "Ignore"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Error Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtdes 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Form2.frx":3AFA
         Top             =   1680
         Width           =   6615
      End
      Begin VB.Label lblsource 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   4980
      End
      Begin VB.Label lblno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Source:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Errno, Errsource, errdes As String

Private Sub FlatBttn1_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblno.Caption = Errno
lblsource.Caption = Errsource
txtdes.Text = errdes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
opndbaseFMIS.Execute "insert into tbl_ErrorDetail ([no],[source],[description],[userID],[datetime],[Status])" & _
" values (" & lblno.Caption & ",'" & lblsource.Caption & "','" & txtdes.Text & "','" & Trim(ActiveUserID) & "','" & Now & "',0) "
End Sub
