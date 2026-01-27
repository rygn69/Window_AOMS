VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmcashCashReceipts_Option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Journal"
   ClientHeight    =   5325
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   4095
   Icon            =   "frmCashReceipts_Option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCashReceipts_Option.frx":09EA
   ScaleHeight     =   5325
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
      Begin VB.ComboBox cmb_FundType 
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
         ItemData        =   "frmCashReceipts_Option.frx":B099
         Left            =   240
         List            =   "frmCashReceipts_Option.frx":B0A6
         TabIndex        =   6
         Text            =   "cmb"
         Top             =   600
         Width           =   3420
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1530
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Group by Bank"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   1360
         Width           =   2580
         _ExtentX        =   4551
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
         Format          =   169607171
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
      Begin VB.CheckBox chkRecap 
         Caption         =   "With Recap"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Recap Criteria"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   3375
         Begin VB.OptionButton OptDetailed 
            Caption         =   "Detailed"
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
            TabIndex        =   15
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton OptConso 
            Caption         =   "Consolidated"
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
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fund type"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   25
      FullHeight      =   33
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   4680
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
      Image           =   "frmCashReceipts_Option.frx":B0DC
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   4680
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
      Image           =   "frmCashReceipts_Option.frx":BAD6
      cBack           =   16777215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Select Special Accouts and set the period that you want to print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipts Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmcashCashReceipts_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadSavedReport1(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal FTYPE As String)
Dim opnRCINo As New ADODB.Recordset
Dim frm As New frm_RPTviewer
Dim sql As String
Dim cc As Integer
On Error GoTo bad
'----filters data from the view to finalize and to remove the data redundancy----'
'filter_data "select * from vw_MP_cashDisbursement where accountname LIKE '%" & fund & "%' and (year(Transationdate)='" & TrnYear & "' and month(transactiondate)='" & trnMonth & "' and fundtype = '" & FTYPE & "') order by ,[111-1-11-WW] desc"
'--------------------------------------------------------------------------------'
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\refresh.avi"
Animation1.Play
With frm
If Check1.Value = 0 Then
.NoGroup = True
Else
.NoGroup = False
End If
'.query = "SELECT  * FROM [vw_MP_Final_CashReceipts] where (year(Jevdate)='" & TrnYear & "' and month(jevdate)='" & trnMonth & "' and fundtype = '" & cmb_FundType.Text & "' and transtype in (1,0)) order by substring(jevno,14,8) asc ,[Collection] desc,[deposits] desc,[acountcode] asc "
.Jquery = "Exec MPproc_RPTJournals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & cmb_FundType.Text & "',@Transtype = 1"
.Rquery = "Exec MPproc_RPTRecap_Journals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & cmb_FundType.Text & "',@Transtype = 1"
.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
.fund = Trim(cmb_FundType.Text)
.TrnsType = 1
Set .frm = Me
.Show
End With
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & Me.name & Me.Caption, err.description)
End Sub
Private Sub LoadSavedReport(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal FTYPE As String)
On Error GoTo bad
Dim opnRCINo As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
Dim frm As New frm_RPTviewer

With frm
If Check1.Value = 0 Then
.NoGroup = True
Else
.NoGroup = False
End If

 .Jquery = "Exec MPproc_RPTJournals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 1"

If chkRecap.Value = 1 Then
.WRecap = True
    If OptConso.Value = True Then
    .Rquery = "Exec MPproc_RPTRecap_Journals_Conso @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 1"
    Else
    .Rquery = "Exec MPproc_RPTRecap_Journals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 1"
    End If
Else
.WRecap = False
End If

If chkConsolidated.Value = 1 Then
   .Ifconso = True
Else
   .Ifconso = False
End If
.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
.fund = Trim(cmb_FundType.Text)
.TrnsType = 1
Set .frm = Me
.Show
End With
Exit Sub
bad:
    If err.Number = 364 Then
    Else
    Call LoadErr(err.Number, Me.name, err.description)
    End If
End Sub
Public Sub LoadFund(ByVal cmb As ComboBox)
Dim opnfund As New ADODB.Recordset
Dim cc As Integer

                 opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds  order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    cmb.Clear
    Do Until opnfund.EOF
        cmb.AddItem (opnfund!FundName)
        cmb.ItemData(cc) = opnfund!fundcode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub chkConsolidated_Click()
If chkConsolidated.Value = 1 Then
    Call LoadMotherFund(cmb_FundType)
Else
    Call Form_Load
End If
End Sub

Private Sub chkRecap_Click()
If chkRecap.Value = 1 Then
Frame2.Enabled = True
Else
Frame2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
   ' Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_FundType.Text)
Call PlayAVI(Me.Animation1, "Refresh.avi")
If chkConsolidated.Value = 1 Then
    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_FundType.ItemData(cmb_FundType.ListIndex))
Else
    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_FundType.ItemData(cmb_FundType.ListIndex))
End If
Call StopAvi(Me.Animation1)
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
Call chkRecap_Click
DTPicker2.Value = Now
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
