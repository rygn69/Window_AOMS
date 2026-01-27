VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmTrialBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Balance"
   ClientHeight    =   4965
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   3960
   Icon            =   "frmtrialbalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmtrialbalance.frx":076A
   ScaleHeight     =   4965
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6465
      ScaleWidth      =   8610
      TabIndex        =   13
      Top             =   9840
      Visible         =   0   'False
      Width           =   8635
      Begin VB.TextBox txtfind 
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
         Left            =   1440
         TabIndex        =   16
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtdetails 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   75
         Width           =   7215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Many"
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
         Left            =   120
         TabIndex        =   14
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5175
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   9128
         _Version        =   393216
         BackColor       =   16777215
         BackColorSel    =   8454143
         ForeColorSel    =   0
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   8160
         TabIndex        =   18
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmtrialbalance.frx":AE19
         cBack           =   16777215
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Search Name:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   5925
         Width           =   1335
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Press ENTER "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4920
         TabIndex        =   20
         Top             =   5835
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Details:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
      Begin VB.CheckBox Check2 
         Caption         =   "Cash Flow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cmbRC 
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
         ItemData        =   "frmtrialbalance.frx":E923
         Left            =   120
         List            =   "frmtrialbalance.frx":E930
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2160
         Width           =   3420
      End
      Begin VB.CheckBox chkRC 
         Caption         =   "Responsibility Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkPostClosing 
         Caption         =   "Post Closing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1530
      End
      Begin VB.ComboBox cmb_Accountcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmtrialbalance.frx":E966
         Left            =   0
         List            =   "frmtrialbalance.frx":E968
         TabIndex        =   5
         Top             =   3480
         Width           =   3465
      End
      Begin VB.ComboBox cmb_FundType 
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
         ItemData        =   "frmtrialbalance.frx":E96A
         Left            =   120
         List            =   "frmtrialbalance.frx":E977
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3420
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   360
         TabIndex        =   8
         Top             =   3840
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   172490753
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   635
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
         CustomFormat    =   "MMMM dd yyyy"
         Format          =   172490755
         UpDown          =   -1  'True
         CurrentDate     =   40574
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Accountcode"
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
         Left            =   0
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fund type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton dsf 
      Caption         =   "&View Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frmtrialbalance.frx":E9AD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   4320
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
      Image           =   "frmtrialbalance.frx":F117
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   4320
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
      Image           =   "frmtrialbalance.frx":FB11
      cBack           =   16777215
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Balance"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   3675
   End
End
Attribute VB_Name = "frmTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Private Sub LoadSavedReport()
Dim frm As New frma_RPTfinancialViewer
On Error GoTo bad
Dim sql As String

If Check2.Value = 1 Then
    Report9 = "Trial Balance"
    With frm
    If chkConsolidated.Value = 1 Then
        .Conso = 1
        .FundType = Trim(cmb_fundtype.Text) & " - Consolidated (" & "Cash Flow)"
    Else
        .Conso = 0
        .FundType = Trim(cmb_fundtype.Text) & " (" & "Cash Flow)"
    End If
        .query = "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "',@reports = 'Trialbalance-Cashflow'"
        .dated = Format(DTPicker1.Value, "MMMM dd, YYYY")
        .Show
    End With
Else
    Report9 = "Trial Balance"
    With frm
    
        
        
        If chkConsolidated.Value = 1 Then
            .Conso = 1
            .FundType = Trim(cmb_fundtype.Text) & " - Consolidated"
        Else
            .Conso = 0
            .FundType = Trim(cmb_fundtype.Text)
        End If
        
        If chkPostClosing.Value = 1 Then
            .preclosing = False
            If cmb_fundtype.ItemData(cmb_fundtype.ListIndex) = 0 Then
                .AllFunds = True
                .query = "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "',@reports = 'Trialbalance_Postclosing_AllFunds'"
            Else
                .AllFunds = False
                .query = "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "',@reports = 'Trialbalance-Posting'"
            End If
                
        Else
            .preclosing = True
            If cmb_fundtype.ItemData(cmb_fundtype.ListIndex) = 0 Then
                .AllFunds = True
                .query = "EXECUTE [fmis].[dbo].[MPproc_New_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "',@reports = 'Trialbalance-ALLfunds'"
            Else
                .AllFunds = False
                .query = "EXECUTE [fmis].[dbo].[MPproc_New_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "',@reports = 'Trialbalance-PreClosing'"
            End If
        End If
        
        
        
    .dated = Format(DTPicker1.Value, "MMMM dd, YYYY")
    .Show
    End With
End If
Exit Sub
bad:
    If err.Number = 364 Then
    MsgBox "No Record Found..", vbInformation, "System Message"
    Else
    MsgBox err.description
    End If
End Sub


Public Sub LoadFund()
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
                
opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 
If opnfund.RecordCount <> 0 Then
    cmb_fundtype.Clear
    Do Until opnfund.EOF
        cmb_fundtype.AddItem (opnfund!FundName)
        cmb_fundtype.ItemData(cc) = opnfund!fundcode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_fundtype.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub Check2_Click()
If Check2.Value = 1 Then
chkPostClosing.Value = 0
End If
End Sub

Private Sub chkConsolidated_Click()
If chkConsolidated.Value = 1 Then
    Call LoadMotherFund(cmb_fundtype)
Else
    Call Form_Load
End If
End Sub

Private Sub chkRC_Click()
If chkRC.Value = 1 Then
    cmbrc.Enabled = True
    Call LoadRC(cmbrc)
Else
    cmbrc.Enabled = False
    cmbrc.Clear
End If
End Sub

Private Sub cmb_Accountcode_KeyPress(KeyAscii As Integer)
KeyAscii = AutoFind(cmb_Accountcode, KeyAscii, True)
End Sub

Private Sub cmb_FundType_Change()
'Call loadChildAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub Command1_Click()
Call PlayAVI(Me.Animation1, "Refresh.avi")
  Call LoadSavedReport
Call StopAvi(Me.Animation1)
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_fundtype)
DTPicker1.Value = Now
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
