VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmGeneralJOurnal_option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Journal Report"
   ClientHeight    =   3645
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "By Special Accounts"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   3375
   End
   Begin VB.OptionButton OptCon 
      Caption         =   "Consolidated"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "View Report As:"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   1080
      TabIndex        =   12
      Top             =   6240
      Width           =   3780
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frmGeneraljournal_option.frx":0000
         Left            =   240
         List            =   "frmGeneraljournal_option.frx":000A
         TabIndex        =   13
         Top             =   300
         Width           =   3435
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Fund Type"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   4920
      TabIndex        =   10
      Top             =   3240
      Width           =   3735
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frmGeneraljournal_option.frx":0021
         Left            =   195
         List            =   "frmGeneraljournal_option.frx":002E
         TabIndex        =   11
         Top             =   300
         Width           =   3390
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Account Number:"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   8640
      TabIndex        =   8
      Top             =   4320
      Width           =   3780
      Begin VB.ComboBox Combo1 
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
         TabIndex        =   9
         Top             =   300
         Width           =   3435
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Special Accounts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3780
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
         ItemData        =   "frmGeneraljournal_option.frx":0064
         Left            =   195
         List            =   "frmGeneraljournal_option.frx":0071
         TabIndex        =   6
         Top             =   300
         Width           =   3435
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "For the Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3780
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   200
         TabIndex        =   3
         Top             =   240
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   98107395
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Bank:"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   4800
      TabIndex        =   0
      Top             =   4320
      Width           =   3780
      Begin VB.ComboBox cmb_Fund 
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
         ItemData        =   "frmGeneraljournal_option.frx":00A7
         Left            =   195
         List            =   "frmGeneraljournal_option.frx":00B4
         TabIndex        =   1
         Top             =   300
         Width           =   3435
      End
   End
   Begin MSComctlLib.ListView lvwdummy 
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CheckDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RCI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CheckNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "JEVNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "111"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "VAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "106-Cabalan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "106-Olaer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "106-Canda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "148-Operational"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "AccountCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "SundryDebit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "SundryCredit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "fndtype"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "accountname"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "General Journal Report"
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
      TabIndex        =   17
      Top             =   0
      Width           =   4575
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
      TabIndex        =   16
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   7725
      Left            =   0
      Picture         =   "frmGeneraljournal_option.frx":00C7
      Stretch         =   -1  'True
      Top             =   -3480
      Width           =   4920
   End
End
Attribute VB_Name = "frmGeneralJOurnal_option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Private Sub LoadAccountName(ByVal CHILD As String, ByVal fund As String)
Dim opnCHILD As New ADODB.Recordset
Dim sql As String
Dim cc As Integer

sql = "SELECT accountname FROM vwAMIS_CDJournal WHERE FNDTYPE = '" & fund & "' AND accountname LIKE '%" & CHILD & "%' group by accountname "

opnCHILD.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnCHILD.RecordCount <> 0 Then
    cmb_Fund.Clear
    Do Until opnCHILD.EOF
        cmb_Fund.AddItem (opnCHILD!Accountname)
        cc = cc + 1
        opnCHILD.MoveNext
    Loop
End If
opnCHILD.Close
Set opnCHILD = Nothing
End Sub


Private Sub LoadSavedReport(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal FTYPE As String)
Dim opnRCINo As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
On Error GoTo bad

'----filters data from the view to finalize and to remove the data redundancy----
'filter_data "select * from vwAMIS_CDJournal where accountname LIKE '%" & fund & "%' and (year(checkdate)='" & TrnYear & "' and month(checkdate)='" & trnMonth & "' and fndtype = '" & FTYPE & "') order by checkno,[111-1-11-WW] desc"
'--------------------------------------------------------------------------------'
Dim frm As New frm_RPTviewer

With frm
.Jquery = "Exec MPproc_RPTJournals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & cmb_FundType.Text & "',@Transtype = 4"
.Rquery = "Exec MPproc_RPTRecap_Journals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & cmb_FundType.Text & "',@Transtype = 4"
.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
.fund = Trim(cmb_FundType.Text)
.TrnsType = 4
Set .frm = Me
.Show
End With
'
'frmGeneralJournaViewer.query = "select * from vw_MP_GeneralJournal where (year(date_)='" & Me.DTPicker2.Year & "' and month(date_)='" & Me.DTPicker2.Month & "' and fundtype = '" & cmb_FundType.Text & "') order by cndate,dvno"
'frmGeneralJournaViewer.accnt = "select * from vw_MP_GeneralJournal_Recap where  year_='" & TrnYear & "' and month_='" & trnMonth & "' and fundtype = '" & FTYPE & "' order by [childaccountcode] asc "
''frmGeneralJournaViewer.query = "select * from vw_MP_GeneralJournal where fundtype = '" & cmb_FundType.Text & "' order by cndate,dvno"
''frmGeneralJournaViewer.accnt = "select * from vw_MP_GeneralJournal_Recap where fundtype = '" & FTYPE & "' order by [accountcode] asc "
''Debug.Print SQL
'frmGeneralJournaViewer.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
'frmGeneralJournaViewer.fund = Trim(cmb_FundType.Text)
'frmGeneralJournaViewer.TYP = Trim(Combo2.Text)
'frmGeneralJournaViewer.bankno = cmb_Fund.Text & " / " & Combo1.Text
'frmGeneralJournaViewer.Show

'opnRCINo.Open SQL, opndbaseFMIS, adOpenStatic, adLockOptimistic

'opnRCINo.Close
'Set opnRCINo = Nothing
Exit Sub
bad:
   
    
End Sub





Private Sub Combo2_click()
Call LoadFund
End Sub

Public Sub LoadFund()
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
                
opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 
If opnfund.RecordCount <> 0 Then
    cmb_FundType.Clear
    Do Until opnfund.EOF
        cmb_FundType.AddItem (opnfund!FundName)
        cmb_FundType.ItemData(cc) = opnfund!fundcode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_FundType.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub Command1_Click()

    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_FundType.Text)

End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
End Sub

