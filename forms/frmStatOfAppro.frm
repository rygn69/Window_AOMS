VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmStatOfAppro 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   6330
   ClientTop       =   3630
   ClientWidth     =   5625
   Icon            =   "frmStatOfAppro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5625
   Begin VB.CommandButton cmdPreview2 
      Caption         =   "Preview"
      Height          =   840
      Left            =   3480
      Picture         =   "frmStatOfAppro.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2805
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1710
      Index           =   2
      Left            =   3030
      TabIndex        =   6
      Top             =   1005
      Width           =   2550
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   8
         Top             =   1065
         Width           =   2400
      End
      Begin VB.OptionButton optDateRange 
         Caption         =   "Date Range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   7
         Top             =   135
         Value           =   -1  'True
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   9
         Top             =   405
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   100073475
         CurrentDate     =   38838
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   10
         Top             =   765
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   100073475
         CurrentDate     =   38868
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   435
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4575
      Picture         =   "frmStatOfAppro.frx":0BAC
      TabIndex        =   5
      Top             =   2805
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fund Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   1005
      Width           =   2970
      Begin VB.ComboBox cboFundType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   435
         Width           =   2850
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.ComboBox cboEco 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   405
         TabIndex        =   2
         Top             =   1155
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.CheckBox chkEco 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   930
      Top             =   8565
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   525
      Left            =   45
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   926
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   39
      FullHeight      =   35
   End
   Begin VB.Frame Frame7 
      Height          =   35
      Left            =   -90
      TabIndex        =   14
      Top             =   840
      Width           =   11220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Define criteria prior to preview a Status of Appropriation Report."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   90
      TabIndex        =   16
      Top             =   360
      Width           =   5520
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status of Appropriation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      TabIndex        =   15
      Top             =   90
      Width           =   2310
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -30
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "frmStatOfAppro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dblPriorAmount As Double
Public dblCurrentOp As Double
Public dblTransferPI As Double
Public dblGovernmentJan As Double
Public dblPriorDebit As Double
Public dblPriorCredit As Double
Public dblGEJanDebit As Double
Public dblGEJanCredit As Double
Private Sub cboEco_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    KeyAscii = AutoFind(cboEco, KeyAscii, False)
    Exit Sub
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboEco_KeyPress"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub
Private Sub cboFundType_KeyPress(KeyAscii As Integer)

    On Error GoTo errHandler
    KeyAscii = AutoFind(cboFundType, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboFundType_KeyPress"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Sub cmdPreview2_Click()
'Call LoadReport
End Sub

Private Sub DTPicker1_Change()
    DTPicker1.Value = Month(DTPicker1.Value) & "/" & "1" & "/" & Year(DTPicker1.Value)
End Sub



Private Sub Form_Load()

    On Error GoTo errHandler
    DTPicker2.Value = Now
    Call LoadFundType(cboFundType)
    Exit Sub
 
errHandler:
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Load"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo errHandler
    WindowsXPC1.EndWinXPCSubClassing
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Function MantYerRange() As String
Dim strMantYerRange As String
Dim xxx As Integer
     
    strMantYerRange = ""
    For xxx = 0 To DateDiff("m", DTPicker1.Value, DTPicker2.Value)
        If xxx = 0 Then
            strMantYerRange = "'" & "" & Format(DateAdd("m", xxx, DTPicker1.Value), "M-YY") & "" & "'"
        Else
            strMantYerRange = strMantYerRange & ",'" & "" & Format(DateAdd("m", xxx, DTPicker1.Value), "M-YY") & "" & "'"
        End If
    Next
    MantYerRange = strMantYerRange
End Function

Private Sub optDateRange_Click()

    On Error GoTo errHandler
    If optDateRange.Value = True Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".optDateRange_Click"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Sub LoadReport()
'    strReportName = "SOA"
'    Call SetAnimation(frmStatOfAppro.Animation1)
'    crptPPAallotment.Database.SetDataSource opndbaseFMIS.Execute("SELECT * from [vwPPA_MonthlyAllotment_Report] where ppayear >= '" & DTPicker1.Year & "' and ppayear <= '" & DTPicker2.Year & "' and ppamonth >= '" & DTPicker1.Month & "' and ppamonth <= '" & DTPicker2.Month & "'")
'    crptPPAallotment.txtmonth.SetText "As of " & Format(DTPicker1.Value, "MMMM dd,yyyy") & "to " & Format(DTPicker2.Value, "MMMM dd,yyyy")
'    crptPPAallotment.txtPrintedby.SetText "Printed by: " & getUserName(ActiveUserID, "FullName")
'    'Call TransactionLogging("Print Preview", "Status of Appropriation", Me.Caption)
'    Call UnsetAnimation(frmStatOfAppro.Animation1)
'    PreviewForm.Show vbModal
End Sub
