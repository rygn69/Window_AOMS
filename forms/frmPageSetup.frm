VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmPageSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Setup"
   ClientHeight    =   3360
   ClientLeft      =   3405
   ClientTop       =   2955
   ClientWidth     =   3690
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3690
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
      Left            =   2640
      TabIndex        =   14
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
      Left            =   1560
      TabIndex        =   13
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
      Width           =   3435
      Begin VB.CheckBox Check1 
         Caption         =   "Use Default Margin"
         Height          =   300
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   2160
      End
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
         Text            =   "Text1"
         Top             =   270
         Width           =   1020
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
         Left            =   1605
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   780
         Width           =   1020
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
         Left            =   1605
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1290
         Width           =   1020
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
         Left            =   1605
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Top Margin"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   435
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left Margin"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   915
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bottom Margin"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1425
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Right Margin"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "in"
         Height          =   195
         Index           =   0
         Left            =   2745
         TabIndex        =   8
         Top             =   405
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "in"
         Height          =   195
         Index           =   1
         Left            =   2745
         TabIndex        =   7
         Top             =   915
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "in"
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   6
         Top             =   1410
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "in"
         Height          =   195
         Index           =   3
         Left            =   2715
         TabIndex        =   5
         Top             =   1950
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

cmdOK.Enabled = False
If Check1.Value = 1 Then
    WritetoIniNewValue
End If
'--------Applying new setting--------------
Select Case ReportName
    Case "AcctAdvice"
        rptAccntAdvice.TopMargin = Text1.Text * 1440
        rptAccntAdvice.LeftMargin = Text2.Text * 1440
        rptAccntAdvice.BottomMargin = Text3.Text * 1440
        rptAccntAdvice.RightMargin = Text4.Text * 1440
    Case "JEV"
        rptJEV.TopMargin = Text1.Text * 1440
        rptJEV.LeftMargin = Text2.Text * 1440
        rptJEV.BottomMargin = Text3.Text * 1440
        rptJEV.RightMargin = Text4.Text * 1440
            
End Select

frmViewer.CRViewer1.Refresh
cmdOK.Enabled = True
Unload Me
End Sub


Private Sub Form_Load()
Dim usedef As Variant
'reading default --------
usedef = readTXTDATA("usepagesetup", "use", ReportLocation & "\pagesetup.ini")

Select Case usedef
    Case "No"
        Check1.Value = 0
    Case "Yes"
        Check1.Value = 1
End Select
WindowsXPC1.InitSubClassing

'------------------------------------------------
If Check1.Value = 1 Then
    Call LoadSavePageValue
Else
    Select Case ReportName
        Case "AcctAdvice"
            Text1.Text = rptAccntAdvice.TopMargin / 1440
            Text2.Text = rptAccntAdvice.LeftMargin / 1440
            Text3.Text = rptAccntAdvice.BottomMargin / 1440
            Text4.Text = rptAccntAdvice.RightMargin / 1440
        Case "JEV"
            Text1.Text = rptJEV.TopMargin / 1440
            Text2.Text = rptJEV.LeftMargin / 1440
            Text3.Text = rptJEV.BottomMargin / 1440
            Text4.Text = rptJEV.RightMargin / 1440
      
    End Select
End If
End Sub
Private Sub LoadSavePageValue()
Dim newval(0 To 3) As Variant

newval(0) = readTXTDATA("Page Setup", "Top", ReportLocation & "\pagesetup.ini")
newval(1) = readTXTDATA("Page Setup", "Left", ReportLocation & "\pagesetup.ini")
newval(2) = readTXTDATA("Page Setup", "Bottom", ReportLocation & "\pagesetup.ini")
newval(3) = readTXTDATA("Page Setup", "Right", ReportLocation & "\pagesetup.ini")

Text1.Text = newval(0)
Text2.Text = newval(1)
Text3.Text = newval(2)
Text4.Text = newval(3)
End Sub

Private Sub WritetoIniNewValue()
WritePrivateProfileString "Page Setup", "Top", Text1.Text, ReportLocation & "\pagesetup.ini"
WritePrivateProfileString "Page Setup", "Left", Text2.Text, ReportLocation & "\pagesetup.ini"
WritePrivateProfileString "Page Setup", "Bottom", Text3.Text, ReportLocation & "\pagesetup.ini"
WritePrivateProfileString "Page Setup", "Right", Text4.Text, ReportLocation & "\pagesetup.ini"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPageSetup = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub
