VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Option_trialbalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3810
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   3960
   Icon            =   "frm_Option_TrialBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   25
      FullHeight      =   33
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Special Accounts:"
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
      Height          =   900
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3540
      Begin VB.ComboBox cmb_FundType 
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
         ItemData        =   "frm_Option_TrialBalance.frx":076A
         Left            =   195
         List            =   "frm_Option_TrialBalance.frx":0777
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   3180
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
      Left            =   1440
      Picture         =   "frm_Option_TrialBalance.frx":07AD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "From"
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
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   3540
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   57344003
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Balance"
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
      TabIndex        =   7
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Select Special Accouts and set                  the period that you want to print"
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   7965
      Left            =   0
      Picture         =   "frm_Option_TrialBalance.frx":0F17
      Stretch         =   -1  'True
      Top             =   -3240
      Width           =   4920
   End
End
Attribute VB_Name = "frm_Option_trialbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Private Sub LoadSavedReport()
On Error GoTo bad
Dim SQL As String

Report9 = "Trial Balance"
frmSubsidiaryLedgerViewer.query = "SELECT  [Accntcode],[FundType],accounttitle,[sumDebit],[SumCredit],year_,month_ FROM [fmis].[dbo].[vw_MP_Final_TrialBalance] where fundtype ='" & cmb_FundType.Text & "' and year_ = '" & DTPicker2.Year & "' and month_ = '" & DTPicker2.Month & "'" & _
                                    "group by [Accntcode],[FundType],accounttitle,[sumDebit],[SumCredit],year_,month_ order by accntcode "
                                    

frmSubsidiaryLedgerViewer.mnth = Format(DTPicker2.Value, "MMMM yyyy")
frmSubsidiaryLedgerViewer.Show
Exit Sub
bad:
    If err.Number = 364 Then
    MsgBox "No Record Found..", vbInformation, "System Message"
    Else
    MsgBox err.Description
    End If
End Sub

Public Sub LoadFund()
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
                
opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 
If opnfund.RecordCount <> 0 Then
    cmb_FundType.Clear
    Do Until opnfund.EOF
        cmb_FundType.AddItem (opnfund!FundName)
        cmb_FundType.ItemData(cc) = opnfund!FundCode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_FundType.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub cmb_FundType_Change()
Call loadChildAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub



Private Sub Command1_Click()

  Call LoadSavedReport
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
End Sub



