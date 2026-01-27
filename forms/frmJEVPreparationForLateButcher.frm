VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJEVPreparation 
   Caption         =   "JEV Preparation"
   ClientHeight    =   9585
   ClientLeft      =   -150
   ClientTop       =   2865
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   15120
   Begin VB.CommandButton btnReturn 
      Caption         =   "Return To PA"
      Enabled         =   0   'False
      Height          =   975
      Left            =   11040
      TabIndex        =   44
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   12840
      TabIndex        =   43
      Top             =   600
      Width           =   1665
   End
   Begin VB.TextBox txtClaimantCode 
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      Height          =   255
      Left            =   12240
      TabIndex        =   39
      Top             =   8640
      Width           =   1575
   End
   Begin VB.ComboBox cmb_month 
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
      Left            =   13230
      TabIndex        =   31
      Top             =   4350
      Width           =   1230
   End
   Begin VB.TextBox txtDate 
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
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1110
      Width           =   2565
   End
   Begin VB.ComboBox cmb_trnYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13230
      TabIndex        =   27
      Top             =   3930
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      Height          =   1845
      Left            =   420
      TabIndex        =   14
      Top             =   1680
      Width           =   14115
      Begin VB.ComboBox cmbrc 
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
         Left            =   9840
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton btnParticular 
         Caption         =   "..."
         Height          =   255
         Left            =   9120
         TabIndex        =   42
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   40
         ToolTipText     =   "Click here to select claimant..."
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtFund 
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1305
         Width           =   1860
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
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
         Left            =   12030
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1275
         Width           =   1860
      End
      Begin VB.TextBox txtParticular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   5160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   540
         Width           =   4290
      End
      Begin VB.TextBox txtAlobs 
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   4260
      End
      Begin VB.TextBox txtClaimant 
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1305
         Width           =   4260
      End
      Begin VB.TextBox txtRC 
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
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   4050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
         Height          =   195
         Left            =   9900
         TabIndex        =   26
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross)"
         Height          =   195
         Left            =   12060
         TabIndex        =   24
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   195
         Left            =   5100
         TabIndex        =   22
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alobs/OBR No:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
         Height          =   195
         Left            =   9780
         TabIndex        =   17
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   12240
      TabIndex        =   12
      Top             =   5145
      Width           =   2265
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      Height          =   360
      Left            =   12225
      TabIndex        =   11
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      Height          =   720
      Left            =   435
      TabIndex        =   6
      Top             =   3705
      Width           =   7830
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   195
         Left            =   6405
         TabIndex        =   10
         Tag             =   "04"
         Top             =   300
         Width           =   1230
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash Disbursement"
         Height          =   195
         Left            =   4245
         TabIndex        =   9
         Tag             =   "03"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCheck 
         Caption         =   "Check Disbursement"
         Height          =   195
         Left            =   1965
         TabIndex        =   8
         Tag             =   "02"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCollection 
         Caption         =   "Collection"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Tag             =   "01"
         Top             =   285
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   405
      ScaleHeight     =   4200
      ScaleWidth      =   10800
      TabIndex        =   3
      Top             =   4890
      Width           =   10830
      Begin VB.ComboBox cmbEntry 
         Height          =   315
         Left            =   3120
         TabIndex        =   38
         Text            =   "cmbEntry"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   3720
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   7408
         _Version        =   393216
         FixedCols       =   0
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
      End
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   405
      TabIndex        =   1
      Top             =   975
      Width           =   4845
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disapprove"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtJEVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8880
      TabIndex        =   37
      Top             =   1050
      Width           =   825
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9795
      TabIndex        =   36
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Trn Year :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   33
      Top             =   4005
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Month of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   32
      Top             =   4425
      Width           =   675
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5670
      TabIndex        =   29
      Top             =   885
      Width           =   1035
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   885
      Left            =   12285
      Top             =   3870
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Vouchers Prepared with JEV"
      Height          =   225
      Left            =   12270
      TabIndex        =   13
      Top             =   4875
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
      Height          =   195
      Left            =   435
      TabIndex        =   4
      Top             =   4620
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter DV Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   2
      Top             =   720
      Width           =   1290
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   8640
      Top             =   600
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   960
      Left            =   -15
      Top             =   615
      Width           =   8625
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Disbursement Voucher No :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmJEVPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Edited As Boolean
Dim xDebit As Currency
Dim xCredit As Currency
Dim xObR As String
Dim xNAcode As String
Dim CUFlag As Boolean           'Claimant Update Flag
Dim XFlag As Boolean
Public isfrom_jevNumbering As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double


Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnParticular_Click()
    CUFlag = True
    txtParticular.Locked = False
End Sub

Private Sub btnPrtJEV_Click()
Dim sql As String

If Edited = True Then
'    SQL = "SELECT dbo.tblAMIS_IncomingDVTrns.RCenterCode, dbo.tblAMIS_JournalEntry.TransDate, dbo.tblAMIS_JournalEntry.TransType," & _
'            "dbo.tblAMIS_JournalEntry.FmisAccntCode, dbo.tblREF_AIS_ChartofAccounts.AccountNameFull, dbo.tblREF_AIS_ChartofAccounts.ChildAccountCode," & _
'            "dbo.tblAMIS_JournalEntry.Amount, dbo.tblAMIS_JournalEntry.DebitCredit, dbo.tblAMIS_JournalEntry.Actioncode," & _
'            "dbo.tblAMIS_IncomingDVTrns.Particular , dbo.tblAMIS_IncomingDVTrns.ClaimantCode FROM dbo.tblAMIS_JournalEntry INNER JOIN " & _
'            "dbo.tblAMIS_IncomingDVTrns ON dbo.tblAMIS_JournalEntry.DVNo = dbo.tblAMIS_IncomingDVTrns.DVNo AND " & _
'            "dbo.tblAMIS_JournalEntry.Actioncode = dbo.tblAMIS_IncomingDVTrns.Actioncode INNER JOIN " & _
'            "dbo.tblREF_AIS_ChartofAccounts ON dbo.tblAMIS_JournalEntry.FmisAccntCode = dbo.tblREF_AIS_ChartofAccounts.FMISAccountCode AND " & _
'            "dbo.tblAMIS_JournalEntry.ActionCode = dbo.tblREF_AIS_ChartofAccounts.Active " & _
'            "WHERE (dbo.tblREF_AIS_ChartofAccounts.FundType ='" & GetFundName(txtFund.Text) & "') AND (dbo.tblAMIS_JournalEntry.DVNo ='" & List1.Text & "')"
    
sql = "SELECT dbo.tblAMIS_IncomingDVTrns.RCenterCode, dbo.tblAMIS_JournalEntry.TransDate, dbo.tblAMIS_JournalEntry.TransType," & _
            "dbo.tblAMIS_JournalEntry.FmisAccntCode, dbo.tblREF_AIS_ChartofAccounts.AccountNameFull, dbo.tblREF_AIS_ChartofAccounts.ChildAccountCode," & _
            "dbo.tblAMIS_JournalEntry.Amount, dbo.tblAMIS_JournalEntry.DebitCredit, dbo.tblAMIS_JournalEntry.Actioncode," & _
            "dbo.tblAMIS_IncomingDVTrns.Particular , dbo.tblAMIS_IncomingDVTrns.ClaimantCode FROM dbo.tblAMIS_JournalEntry INNER JOIN " & _
            "dbo.tblAMIS_IncomingDVTrns ON dbo.tblAMIS_JournalEntry.DVNo = dbo.tblAMIS_IncomingDVTrns.DVNo AND " & _
            "dbo.tblAMIS_JournalEntry.Actioncode = dbo.tblAMIS_IncomingDVTrns.Actioncode INNER JOIN " & _
            "dbo.tblREF_AIS_ChartofAccounts ON dbo.tblAMIS_JournalEntry.FmisAccntCode = dbo.tblREF_AIS_ChartofAccounts.FMISAccountCode AND " & _
            "(dbo.tblAMIS_JournalEntry.ActionCode = dbo.tblREF_AIS_ChartofAccounts.Active or dbo.tblAMIS_JournalEntry.ActionCode=5 )" & _
            "WHERE (dbo.tblREF_AIS_ChartofAccounts.FundType ='" & GetFundName(txtFund.Text) & "') AND (dbo.tblAMIS_JournalEntry.DVNo ='" & List1.Text & "')"
    
    'Debug.Print sql
    ReportName = "JEV"
    rptJEV.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEV.txtRC.SetText txtRC.Text
    rptJEV.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    
    If chkSTP.Value = 1 Then
        rptJEV.Line1.Suppress = True
        rptJEV.Line2.Suppress = True
        rptJEV.Line3.Suppress = True
        rptJEV.Line4.Suppress = True
        rptJEV.Line5.Suppress = True
        rptJEV.Line6.Suppress = True
        rptJEV.Line8.Suppress = True
        rptJEV.Line9.Suppress = True
        rptJEV.Line10.Suppress = True
        rptJEV.Line11.Suppress = True
        rptJEV.Line12.Suppress = True
        rptJEV.Line13.Suppress = True
        rptJEV.Line14.Suppress = True
        rptJEV.Line15.Suppress = True
        rptJEV.Line16.Suppress = True
        rptJEV.Line17.Suppress = True
        rptJEV.Line18.Suppress = True
        rptJEV.Line19.Suppress = True
        
        rptJEV.Text1.Suppress = True
        rptJEV.Text2.Suppress = True
        rptJEV.Text3.Suppress = True
        rptJEV.Text4.Suppress = True
        rptJEV.Text8.Suppress = True
        rptJEV.Text9.Suppress = True
        rptJEV.Text12.Suppress = True
        rptJEV.Text13.Suppress = True
        rptJEV.Text15.Suppress = True
        rptJEV.Text16.Suppress = True
        rptJEV.Text17.Suppress = True
        rptJEV.Text18.Suppress = True
        rptJEV.Text19.Suppress = True
        rptJEV.Text20.Suppress = True
        rptJEV.Text21.Suppress = True
        rptJEV.Text22.Suppress = True
        rptJEV.Text25.Suppress = True
        
    End If
    
    rptJEV.Database.SetDataSource opndbaseFMIS.Execute(sql)
    rptJEV.Database.Verify
    frmViewer.Show 1
End If

End Sub

Private Sub btnReturn_Click()
    If MsgBox("Are you sure you want to return DV No.: " & txtDVNo.Text & " to Pre-Audit?", vbQuestion + vbYesNo, "System Security") = vbYes Then
        If ChkIfAlreadyJEV(txtDVNo.Text) = "" Then
            opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [PAout]=0, [PAoutDate]=null, [PADesc]=null, [OutBy]=null where [DVNo]='" & txtDVNo.Text & "' and actioncode=1"
        End If
        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    End If
End Sub

Private Sub btnSearch_Click()
    frmDVSearch.Show 1
End Sub

Private Sub cmb_month_Click()
    Call LoadPrevTrans
End Sub

Private Sub LoadPrevTrans()
Dim PRec As New ADODB.Recordset
Dim X As Integer

    List1.Clear
    List1.Enabled = False
    PRec.Open ("Select DVNo, min(trnno) as trnno From tblAMIS_JournalEntry Where (len(ApprovedByID)<>4 or ApprovedByID is null) and Actioncode=1 Group By DVNo order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
        For X = 1 To PRec.RecordCount
            List1.AddItem PRec!DVNo
            PRec.MoveNext
        Next X
        List1.Enabled = True
    End If
    PRec.Close
    Set PRec = Nothing
    
End Sub

Private Sub cmb_trnYear_Click()
    Call LoadPrevTrans
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "1"
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
        End If
        cmbEntry.Visible = False
        Call GetSum
    Else
        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
    End If

End Sub





Private Sub cmdOK_Click()

End Sub

Private Sub cmbRC_Click()
    If Trim(cmbrc.Text) <> "" Then
        txtRC = Trim(cmbrc.Text)
        txtRC.Visible = True
        cmbrc.Visible = False
    End If
End Sub

Private Sub cmbRC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbRC_Click
    End If
End Sub

Private Sub Form_Load()
    
    Edited = False
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    
    ActiveUserID = Trim(ActiveUserID)
    
End Sub

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 50
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 2500
    MSFlexGrid1.ColWidth(2) = 5000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    'If LCase(Trim(lblMode)) = "Edit" Then
        MSFlexGrid1.ColWidth(5) = 1500
    'Else
    '    MSFlexGrid1.ColWidth(5) = 0
    'End If
    
    
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Sub List1_Click()
    Call LoadJEVDetails(List1.Text)
    cmbrc.Visible = False
    txtRC.Visible = True
End Sub

Private Sub LoadJEVDetails(ByVal DVNo As String)
Dim DRec As New ADODB.Recordset
Dim X As Integer
    
    CUFlag = False
    txtParticular.Locked = True
    xNAcode = ""
    Edited = True
    lblMode.Caption = "EDIT"
    DRec.Open ("Select * From [tblAMIS_JournalEntry] Where [DVNo]='" & DVNo & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        
        txtDVNo.Text = DRec![DVNo]
        'txtJEVNo.Text = DRec!JEVNo
        txtDate.Text = DRec![TransDate]
        If CInt(optCollection.Tag) = DRec![TransType] Then optCollection.Value = True
        If CInt(optCheck.Tag) = DRec![TransType] Then optCheck.Value = True
        If CInt(optCash.Tag) = DRec![TransType] Then optCash.Value = True
        If CInt(optOther.Tag) = DRec![TransType] Then optOther.Value = True
        
        If DRec!continuing = 1 Then
            XFlag = True
        Else
            XFlag = False
        End If
    
    End If
    DRec.Close
    Set DRec = Nothing
        
    DRec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & txtDVNo.Text & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        txtClaimant.Text = GetClaimant(DRec!ClaimantCode)
        txtClaimantCode = DRec!ClaimantCode
        txtRC.Text = GetOfficeName(DRec!RCenter, "OfficeMedium")
        txtParticular.Text = DRec!Particular
        txtFund.Text = DRec!FundType
        txtAmount.Text = DRec![GAmount]
        If DRec!NonAlobs = 1 Then
            xObR = GetNonAlobsName(DRec!ObrNo)
            xNAcode = DRec!ObrNo
        Else
            xObR = DRec!ObrNo
        End If
        txtAlobs.Text = xObR
    End If
    DRec.Close
    Set DRec = Nothing
        
    Call SetGrid
    'DRec.Close
    DRec.Open ("Select * From [tblAMIS_JournalEntry] Where [DVNo]='" & DVNo & "' And (ActionCode=1 or actioncode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        For X = 1 To DRec.RecordCount
            MSFlexGrid1.TextMatrix(X, 0) = DRec![FmisAccntCode]
            MSFlexGrid1.TextMatrix(X, 1) = GetAccountCodeByFMISAccountCode(DRec![FmisAccntCode])
            MSFlexGrid1.TextMatrix(X, 2) = GetAccountNameByFMISAccountCode(DRec![FmisAccntCode])
            If DRec![DebitCredit] = 0 Then
                MSFlexGrid1.TextMatrix(X, 4) = DRec!Amount
            Else
                MSFlexGrid1.TextMatrix(X, 3) = DRec!Amount
            End If
            If LCase(Trim(lblMode)) = "edit" Then MSFlexGrid1.TextMatrix(X, 5) = DRec!ActionCode  ' for coloraly purpose
            DRec.MoveNext
        Next X
        Call GetSum
    End If
    DRec.Close
    Set DRec = Nothing
    
    Call LoadAccountsByFund(Trim(txtFund.Text))

End Sub

Private Sub MSFlexGrid1_Click()

    Select Case MSFlexGrid1.Col
    Case 1 'AccntCode
        txt_entry.Visible = False
        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
        cmbEntry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            cmbEntry.Text = MSFlexGrid1.Text
        Else
            cmbEntry.ListIndex = -1
        End If
    Case 3 To 5 'Debit/Credit
        cmbEntry.Visible = False
        txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
        txt_entry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            txt_entry.Text = MSFlexGrid1.Text
            txt_entry.SelStart = 0
            txt_entry.SelLength = Len(txt_entry.Text)
        Else
            txt_entry.Text = ""
        End If
        txt_entry.SetFocus
    
    Case Else
        txt_entry.Visible = False
        cmbEntry.Visible = False
    End Select

End Sub

Private Sub optCash_Click()
    'txtJEVNo.Text = GetNewJEV(optCash.Tag)
End Sub

Private Sub optCheck_Click()
    'txtJEVNo.Text = GetNewJEV(optCheck.Tag)
End Sub

Private Sub optCollection_Click()
    'txtJEVNo.Text = GetNewJEV(optCollection.Tag)
End Sub

Private Sub optOther_Click()
    'txtJEVNo.Text = GetNewJEV(optOther.Tag)
End Sub

Private Function GetNewJEV(ByVal JournalCode As String) As String
Dim JREc As New ADODB.Recordset
Dim xCode As String

    GetNewJEV = ""
    xCode = GetFundCODE(txtFund.Text) & "-" & Format(Now, "yy-mm") & "-" & JournalCode
    JREc.Open ("Select * from tblAMIS_JournalEntry where JEVNo like '" & xCode & "%' Order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If JREc.RecordCount > 0 Then
        GetNewJEV = xCode & "-" & Format(CInt(Right(JREc!JEVNo, 3)) + 1, "000")
    Else
        GetNewJEV = xCode & "-001"
    End If
    JREc.Close
    Set JREc = Nothing
    
End Function

'-----RICHARD--------
Private Function getdetails(signal As Integer) As String
Dim rs As New ADODB.Recordset
Set rs = opndbaseFMIS.Execute("select top 1 rcenter,rcentercode,claimantcode,transactiondate,nonalobs,ooe from [tblAMIS_IncomingDVTrns] Where DVNo='" & Trim(txtDVNo.Text) & "'")
If Not rs.EOF Then
    If signal = 1 Then
        getdetails = Trim(rs(0))
    ElseIf signal = 2 Then
        getdetails = Trim(rs(1))
    ElseIf signal = 3 Then
        getdetails = Trim(rs(2))
    ElseIf signal = 4 Then
        getdetails = Trim(rs(3))
    ElseIf signal = 5 Then
        getdetails = Trim(rs(4))
    ElseIf signal = 6 Then
        getdetails = Trim(rs(5))

    End If
End If
End Function


'Private Function check_coloraly() As Boolean
'Dim nc_debit, nc_credit, c_debit, c_credit As Double, X As Integer
'
'     For X = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(X, 2) <> "TOTAL" Then
'            If MSFlexGrid1.TextMatrix(X, 0) <> "" Then
'                If MSFlexGrid1.TextMatrix(X, 3) <> "" Or MSFlexGrid1.TextMatrix(X, 4) <> "" Then
'                    If Trim(MSFlexGrid1.TextMatrix(X, 5)) = "5" Then
'                        If MSFlexGrid1.TextMatrix(X, 3) <> "" and Then
'
'                        End If
'                    End If
'                End If
'            End If
'        Else
'            Exit For
'        End If
'    Next X
'End Function


'--------------------




Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim X As Integer
Dim xType As Integer, coloraly_signal As Integer


    Select Case Button:
    Case "New":
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
                lblMode.Caption = "NEW"
                txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtRC.Text = ""
                txtParticular.Text = ""
                txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txtDate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                btnReturn.Enabled = False
                
                Call LoadTrnYear(cmb_trnYear)
                Call LoadTrnMonth(cmb_month)
                Call SetGrid
                
    Case "Save":
                If ChkEntry = True Then
                    If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                        
                        If not_coloraly_total_debit <> not_coloraly_total_credit Or coloraly_total_debit <> coloraly_total_credit Then
                            GoTo debit_credit_error
                        End If
                        
                        
                        
                        If optCollection.Value = True Then xType = CInt(optCollection.Tag)
                        If optCash.Value = True Then xType = CInt(optCash.Tag)
                        If optCheck.Value = True Then xType = CInt(optCheck.Tag)
                        If optOther.Value = True Then xType = CInt(optOther.Tag)
                        
                        
                        If Edited = True Then
                            opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where DVNo='" & List1.Text & "' And ActionCode=1"
                        End If
                        
                        If CUFlag = True Then
                            opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [Particular]='" & Trim(Replace(txtParticular.Text, "'", "''")) & "', [ClaimantCode]='" & txtClaimantCode.Text & "' Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
                        End If
                        
                        'DELETES THE COLORALY ENTRY IN THE INCOMINGdvTrns ENTRY and in the journal entry table
                        opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set ACTIONCODE=6 Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=5"
                        opndbaseFMIS.Execute "Update [tblAMIS_JournalEntry] set ACTIONCODE=6 Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=5"
                        
                        If xNAcode <> "" Then
                            xObR = xNAcode
                        End If
                        
                        For X = 1 To MSFlexGrid1.Rows - 1
                            If MSFlexGrid1.TextMatrix(X, 2) <> "TOTAL" Then
                                If MSFlexGrid1.TextMatrix(X, 0) <> "" Then
                                    If MSFlexGrid1.TextMatrix(X, 3) <> "" Or MSFlexGrid1.TextMatrix(X, 4) <> "" Then
                                        opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (TransType,DVNo,ObrNo,FmisAccntCode,Amount,DebitCredit,TransDate,UserID,Actioncode,DateTimeEntered,Continuing) values (" & xType & ",'" & Trim(Replace(txtDVNo.Text, "'", "''")) & "','" & xObR & "'," & CLng(MSFlexGrid1.TextMatrix(X, 0)) & "," & CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(X, 3)), MSFlexGrid1.TextMatrix(X, 3), 0)) + CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(X, 4)), MSFlexGrid1.TextMatrix(X, 4), 0)) & "," & IIf(Trim(MSFlexGrid1.TextMatrix(X, 3)) = "", 0, 1) & ",'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "','" & IIf(Trim(MSFlexGrid1.TextMatrix(X, 5)) = "1" Or Trim(MSFlexGrid1.TextMatrix(X, 5)) = "", 1, 5) & "' ,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ")"
                                        ' saves the record to the IncomingDVTrns table if coloraly entry
                                        If Trim(MSFlexGrid1.TextMatrix(X, 5)) = "5" And coloraly_signal = 0 Then
                                            
                                            opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,NonAlobs) values ('" & Trim(Replace(txtDVNo.Text, "'", "''")) & "','" & xObR & "','" & Trim(txtFund) & "','" & getdetails(1) & "','" & getdetails(2) & "','" & getdetails(6) & "','" & getdetails(3) & "','" & Replace(Trim(txtParticular), "'", "''") & "','" & Trim(txtAmount) & "', '" & getdetails(4) & "','" & ActiveUserID & "','5','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & getdetails(5) & "' ) "
                                            coloraly_signal = 1
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next X
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                    

                    End If
                Else
debit_credit_error:
                    MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
                
                End If
    Case "Delete":
                If Edited = True Then
                    If InStr(ChkIfAlreadyJEV(txtDVNo.Text), "Approved") <> 1 Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set UserID=UserID + '," & ActiveUserID & "',Actioncode=3,DateTimeEntered=DateTimeEntered +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & txtDVNo.Text & "' and Actioncode=1"
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                        End If
                    Else
                        MsgBox "This transaction is already approved!" & vbCrLf & vbCrLf & "Delete operation cancelled!", vbExclamation + vbOKOnly
                    End If
                End If
    Case "Close":
                If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                    Unload Me
                End If
    End Select
    
    
End Sub

Private Function coloraly() As Boolean
Dim X As Integer
    For X = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(X, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(X, 5) <> "" Then
                If MSFlexGrid1.TextMatrix(X, 5) = "5" Then
                    coloraly = True
                    Exit Function
                End If
            End If
        Else
            Exit For
        End If
    Next X
End Function


Private Function ChkEntry() As Boolean

    ChkEntry = False
    If Trim(txtDVNo.Text) <> "" And txtAlobs.Text <> "" And txtClaimant.Text <> "" And txtRC.Text <> "" And txtParticular.Text <> "" And txtFund.Text <> "" And txtAmount.Text <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
        If coloraly = True Then GoTo coloraly_jmp 'coloraly consideration - set chkentry to true even if not balance
            If Format(xDebit, "###,##0.00") = Format(txtAmount.Text, "###,##0.00") Then
coloraly_jmp:
                ChkEntry = True
            End If
        End If
    End If
    
End Function

Private Sub LoadExcessDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim X As Integer
Dim Y As Integer

    Call SetGrid
    OREc.Open ("Select * from [tblBMS_ExcessControl] where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For X = 1 To OREc.RecordCount
            For Y = 0 To cmbEntry.ListCount - 1
                If cmbEntry.List(Y) = "401" Then
                    cmbEntry.ListIndex = Y
                    Exit For
                Else
                    If Y = cmbEntry.ListCount - 1 Then
                        cmbEntry.ListIndex = -1
                    End If
                End If
            Next Y
            MSFlexGrid1.TextMatrix(X, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(X, 1) = "401"
            MSFlexGrid1.TextMatrix(X, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
            MSFlexGrid1.TextMatrix(X, 4) = OREc!Amount
            OREc.MoveNext
        Next X
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub


Private Sub LoadObRDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim X As Integer
    
    Call SetGrid
    OREc.Open ("Select * from tblBMS_SubsidiaryLedger where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For X = 1 To OREc.RecordCount
            MSFlexGrid1.TextMatrix(X, 0) = OREc!FMISAccountCode
            MSFlexGrid1.TextMatrix(X, 1) = GetAccountCodeByFMISAccountCode(OREc!FMISAccountCode)
            MSFlexGrid1.TextMatrix(X, 2) = GetAccountNameByFMISAccountCode(OREc!FMISAccountCode)
            MSFlexGrid1.TextMatrix(X, 4) = OREc!Amount
            OREc.MoveNext
        Next X
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub

Public Sub LoadAccountsByFund(ByVal fundmedium As String)
Dim ARec As New ADODB.Recordset
Dim X As Integer
Dim FundName As String

    cmbEntry.Clear
    cmbEntry.Visible = False
    FundName = GetFundName(fundmedium)
    ARec.Open ("Select distinct * from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & FundName & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For X = 1 To ARec.RecordCount
            cmbEntry.AddItem ARec![ChildAccountCode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FMISAccountCode]
            ARec.MoveNext
        Next X
    End If
    ARec.Close
    Set ARec = Nothing
    
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = txt_entry.Text
        If MSFlexGrid1.Col = 3 Then
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            End If
        
        ElseIf MSFlexGrid1.Col <> 5 Then
            
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
        End If
        txt_entry.Visible = False
        Call GetSum
    End If
End Sub

Private Sub GetSum()
Dim X As Integer
    not_coloraly_total_debit = 0
    not_coloraly_total_credit = 0
     coloraly_total_credit = 0
     coloraly_total_debit = 0
      
    xDebit = 0
    xCredit = 0
    For X = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(X, 0) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 3) = "", 0, MSFlexGrid1.TextMatrix(X, 3)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 4) = "", 0, MSFlexGrid1.TextMatrix(X, 4)))
                If Trim(MSFlexGrid1.TextMatrix(X, 5)) <> 5 Then
                    not_coloraly_total_debit = not_coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 3) = "", 0, MSFlexGrid1.TextMatrix(X, 3)))
                    not_coloraly_total_credit = not_coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 4) = "", 0, MSFlexGrid1.TextMatrix(X, 4)))
                Else
                    coloraly_total_debit = coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 3) = "", 0, MSFlexGrid1.TextMatrix(X, 3)))
                    coloraly_total_credit = coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(X, 4) = "", 0, MSFlexGrid1.TextMatrix(X, 4)))
                End If
        Else
            MSFlexGrid1.TextMatrix(X, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(X, 3) = xDebit
            MSFlexGrid1.TextMatrix(X, 4) = xCredit
            Exit For
        End If
    Next X
    
End Sub

Private Function ChkIfAlreadyJEV(ByVal DVNo As String) As String
Dim JREc As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    JREc.Open ("Select * from tblAMIS_JournalEntry where DVNo='" & DVNo & "' and Actioncode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If JREc.RecordCount > 0 Then
        If Not IsNull(JREc!ApprovedByID) Then
            ChkIfAlreadyJEV = "Approved" & "-" & JREc!JEVNo
        Else
            ChkIfAlreadyJEV = DVNo
        End If
    End If
    JREc.Close
    Set JREc = Nothing
    
End Function

Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String

    If KeyAscii = 13 Then
        btnReturn.Enabled = False
        CUFlag = False
        txtParticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkDVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                DVRec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & txtDVNo.Text & "' and (ActionCode=1 or ActionCode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If DVRec.RecordCount > 0 Then
                    If DVRec!PAout = 1 Then
                        If DVRec!ReturnFlag = 0 Then
                            btnReturn.Enabled = True
                            If DVRec!NonAlobs = 1 Then
                                xObR = GetNonAlobsName(DVRec!ObrNo)
                                xNAcode = DVRec!ObrNo
                            Else
                                xObR = DVRec!ObrNo
                            End If
                            
                            txtAlobs.Text = xObR
                            txtClaimant.Text = GetClaimant(DVRec!ClaimantCode)
                            txtClaimantCode.Text = DVRec!ClaimantCode
                            txtRC.Text = GetOfficeName(DVRec!RCenter, "OfficeMedium")
                            txtParticular.Text = DVRec!Particular
                            txtFund.Text = DVRec!FundType
                            txtAmount.Text = DVRec!GAmount
                            optCollection.Value = True
                            
                            Call optCollection_Click
                            Call LoadAccountsByFund(Trim(txtFund.Text))
                            
                            XFlag = False
                            If DVRec!continuing = 1 Then
                                    XFlag = True
                                Call LoadExcessDetails(DVRec!ObrNo)
                            Else
                                Call LoadObRDetails(DVRec!ObrNo)
                            End If
                            
                        Else
                            MsgBox "This transaction must pass to pre-audit first!", vbExclamation + vbOKOnly
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                        End If
                    Else
                        MsgBox "Please log out DV No. " & txtDVNo.Text & " on pre-audit first!", vbExclamation + vbOKOnly
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                    End If
                End If
                DVRec.Close
                Set DVRec = Nothing
            ElseIf InStr(1, xAlreadyJEV, "Approved") > 0 Then
                MsgBox "The DV No. " & txtDVNo.Text & " is already approved with JEV No. " & Mid(xAlreadyJEV, InStr(1, xAlreadyJEV, "-") + 1) & "!", vbExclamation + vbOKOnly
            Else
                List1.Text = xAlreadyJEV
                Call LoadJEVDetails(xAlreadyJEV)
            End If
        Else
            MsgBox "Invalid DV Number!", vbExclamation
            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
        End If
    End If
End Sub

Public Sub LoadOffice()
Dim OREc As New ADODB.Recordset
Dim X As Integer

cmbrc.Clear

OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For X = 1 To OREc.RecordCount
        cmbrc.AddItem OREc![OfficeMedium]
        cmbrc.ItemData(cmbrc.NewIndex) = OREc!FMISOfficeID
        OREc.MoveNext
    Next X
End If
OREc.Close
Set OREc = Nothing

End Sub
Private Sub txtRC_Click()
'   cmbRC.Visible = True
    'txtRC.Visible = False
End Sub
