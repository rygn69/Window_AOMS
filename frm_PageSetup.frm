VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frm_PageSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Setup"
   ClientHeight    =   3345
   ClientLeft      =   3405
   ClientTop       =   2955
   ClientWidth     =   4575
   Icon            =   "frm_PageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4575
   Begin VB.CheckBox Check1 
      Caption         =   "Use Default Margin"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   2160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   945
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      FrameControl    =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1605
         TabIndex        =   4
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1140
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   2160
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   1
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Top Margin(Inch)"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left Margin(Inch)"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   915
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bottom Margin(Inch)"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   6
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Right Margin(Inch)"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frm_PageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RPTreport As Object
Private Sub Check1_Click()
If Check1.Value = 1 Then
    WritePrivateProfileString "usepagesetup", "use", "Yes", ReportLocation & "\pagesetup.ini"
Else
    WritePrivateProfileString "usepagesetup", "use", "No", ReportLocation & "\pagesetup.ini"
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo bad
        RPTreport.TopMargin = Format(Text1.Text * 1440, "#0.00")
        RPTreport.LeftMargin = Format(Text2.Text * 1440, "#0.00")
        RPTreport.BottomMargin = Format(Text3.Text * 1440, "#0.00")
        RPTreport.RightMargin = Format(Text4.Text * 1440, "#0.00")
Unload Me
Exit Sub
bad:
MsgBox "Noted: " & err.description
End Sub
Private Sub Form_Load()
Dim usedef As Variant
'------------------------------------------------
            Text1.Text = Format(RPTreport.TopMargin / 1440, "#0.00")
            Text2.Text = Format(RPTreport.LeftMargin / 1440, "#0.00")
            Text3.Text = Format(RPTreport.BottomMargin / 1440, "#0.00")
            Text4.Text = Format(RPTreport.RightMargin / 1440, "#0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_PageSetup = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub
