VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmPageSetup 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2745
   ClientLeft      =   4785
   ClientTop       =   2955
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_PageSetUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4395
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   3330
      TabIndex        =   2
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   405
      Left            =   2325
      TabIndex        =   1
      Top             =   2310
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Margin"
      Height          =   2145
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4215
      Begin VB.TextBox txtRight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2910
         TabIndex        =   10
         Text            =   "0"
         Top             =   1335
         Width           =   975
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2910
         TabIndex        =   9
         Text            =   "0"
         Top             =   615
         Width           =   975
      End
      Begin VB.TextBox txtBottom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   855
         TabIndex        =   8
         Text            =   "0"
         Top             =   1335
         Width           =   975
      End
      Begin VB.TextBox txtTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   855
         TabIndex        =   7
         Text            =   "0"
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Right"
         Height          =   210
         Left            =   2280
         TabIndex        =   6
         Top             =   1410
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Left"
         Height          =   210
         Left            =   2280
         TabIndex        =   5
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Bottom"
         Height          =   210
         Left            =   180
         TabIndex        =   4
         Top             =   1410
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Top"
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   600
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   0
      Top             =   0
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      MsgBox_InputBox =   0   'False
      Common_Dialog   =   0   'False
      FrameControl    =   0   'False
      OptionControl   =   0   'False
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Select Case UCase$(strReportName)
        Case "POREG"
            Call SetPageMargin(crptPORegistry)
        Case "CKDJRECAP"
            Call SetPageMargin(CrystalReportRecapCKDJ)
        Case "CDJRECAP"
            Call SetPageMargin(CrystalReportRecapCDJ)
        Case "CRJRECAP"
            Call SetPageMargin(CrystalReportRecapCRJ)
        Case "TB"
            Call SetPageMargin(CrystalReportTB)
        Case "LRC"
            Call SetPageMargin(CrystalReportResponCenter)
        Case "SMRFS"
            Call SetPageMargin(CrystalReportSMRFS)
        Case "CRJ"
            Call SetPageMargin(crpCRJ)
        Case "CDJ"
            Call SetPageMargin(crpCDJ)
        Case "CKDJ"
            Call SetPageMargin(crpCKDJ)
        Case "JEV", "JEVPRINT"
            Call SetPageMargin(crpJEV)
            'Call SetPageMargin(crpJEVDetails)
        Case "ADV"
            Call SetPageMargin(CrystalReportADVICE)
        Case "SL"
            Call SetPageMargin(CrystalReportSL2)
        Case "GL"
            Call SetPageMargin(crpGL2)
        Case "GJRECAP"
            Call SetPageMargin(CrystalReportRecapGJ)
        Case "GJ"
            Call SetPageMargin(crpGJ)
        Case "BS"
            Call SetPageMargin(CrystalReportBS)
        Case "SIAE"
            Call SetPageMargin(CrystalReportSIE)
        Case "SCHEDULE"
            Call SetPageMargin(CrystalReportSked)
        Case "SOCF"
            Call SetPageMargin(CrystalReportSCF)
        Case "SOGE"
            Call SetPageMargin(CrystalReportSGE)
        Case "OTHERSCHEDULE"
            Call SetPageMargin(CrystalReportOtherSked)
        Case "LORPDO"
            Call SetPageMargin(crptListOfReportNoPerDisbursing)
        Case "RRR"
            Call SetPageMargin(CrystalReportRRR)
        Case "LIQREG"
            Call SetPageMargin(crptLiquidationReg)
        Case "LIQCA"
            Call SetPageMargin(CrystalReportCA)
        Case Else
        
    End Select
End Sub
Private Sub SetPageMargin(ByVal crptName As Report)
    crptName.TopMargin = CDbl(txtTop.Text) * 1440
    crptName.BottomMargin = CDbl(txtBottom.Text) * 1440
    crptName.LeftMargin = CDbl(txtLeft.Text) * 1440
    crptName.RightMargin = CDbl(txtRight.Text) * 1440
    dblTop = CDbl(txtTop.Text) * 1440
    dblBottom = CDbl(txtBottom.Text) * 1440
    dblLeft = CDbl(txtLeft.Text) * 1440
    dblRight = CDbl(txtRight.Text) * 1440
    PreviewForm.CRViewer1.Refresh
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'WindowsXPC1.InitSubClassing

    Select Case UCase$(strReportName)
        Case "POREG"
            Call GetCurrentPageMargin(crptPORegistry)
        Case "CKDJRECAP"
            Call GetCurrentPageMargin(CrystalReportRecapCKDJ)
        Case "CDJRECAP"
            Call GetCurrentPageMargin(CrystalReportRecapCDJ)
        Case "CRJRECAP"
            Call GetCurrentPageMargin(CrystalReportRecapCRJ)
        Case "TB"
            Call GetCurrentPageMargin(CrystalReportTB)
        Case "LRC"
            Call GetCurrentPageMargin(CrystalReportResponCenter)
        Case "SMRFS"
            Call GetCurrentPageMargin(CrystalReportSMRFS)
        Case "CRJ"
            Call GetCurrentPageMargin(crpCRJ)
        Case "CDJ"
            Call GetCurrentPageMargin(crpCDJ)
        Case "CKDJ"
            Call GetCurrentPageMargin(crpCKDJ)
        Case "JEV", "JEVPRINT"
            Call GetCurrentPageMargin(crpJEV)
            'Call GetCurrentPageMargin(crpJEVDetails)
        Case "ADV"
            Call GetCurrentPageMargin(CrystalReportADVICE)
        Case "SL"
            Call GetCurrentPageMargin(CrystalReportSL2)
        Case "GL"
            Call GetCurrentPageMargin(crpGL2)
        Case "GJRECAP"
            Call GetCurrentPageMargin(CrystalReportRecapGJ)
        Case "GJ"
            Call GetCurrentPageMargin(crpGJ)
        Case "BS"
            Call GetCurrentPageMargin(CrystalReportBS)
        Case "SIAE"
            Call GetCurrentPageMargin(CrystalReportSIE)
        Case "SCHEDULE"
            Call GetCurrentPageMargin(CrystalReportSked)
        Case "SOCF"
            Call GetCurrentPageMargin(CrystalReportSCF)
        Case "SOGE"
            Call GetCurrentPageMargin(CrystalReportSGE)
        Case "OTHERSCHEDULE"
            Call GetCurrentPageMargin(CrystalReportOtherSked)
        Case "LORPDO"
            Call GetCurrentPageMargin(crptListOfReportNoPerDisbursing)
        Case Else
        
    End Select
End Sub
Private Sub GetCurrentPageMargin(ByVal crptName As Report)
    txtTop.Text = crptName.TopMargin / 1440
    txtBottom.Text = crptName.BottomMargin / 1440
    txtLeft.Text = crptName.LeftMargin / 1440
    txtRight.Text = crptName.RightMargin / 1440
End Sub
