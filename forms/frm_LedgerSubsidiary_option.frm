VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frm_LedgerGeneral_option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger"
   ClientHeight    =   4695
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   4860
   Icon            =   "frm_LedgerSubsidiary_option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   4200
      TabIndex        =   9
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Account Code"
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
      TabIndex        =   7
      Top             =   1920
      Width           =   4620
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
         ItemData        =   "frm_LedgerSubsidiary_option.frx":076A
         Left            =   240
         List            =   "frm_LedgerSubsidiary_option.frx":076C
         TabIndex        =   8
         Top             =   300
         Width           =   4260
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "To"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   2220
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   211288065
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
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
      Width           =   4620
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
         ItemData        =   "frm_LedgerSubsidiary_option.frx":076E
         Left            =   195
         List            =   "frm_LedgerSubsidiary_option.frx":077B
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   4260
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
      Left            =   2520
      Picture         =   "frm_LedgerSubsidiary_option.frx":07B1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
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
      Top             =   2880
      Width           =   2340
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   1980
         _ExtentX        =   3493
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
         Format          =   211288065
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "General Ledger"
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
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   4695
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
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   7965
      Left            =   0
      Picture         =   "frm_LedgerSubsidiary_option.frx":0F1B
      Stretch         =   -1  'True
      Top             =   -3240
      Width           =   4920
   End
End
Attribute VB_Name = "frm_LedgerGeneral_option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Private Sub LoadSavedReport()
Dim sql As String
On Error GoTo bad
Report9 = "General"
frmSubsidiaryLedgerViewer.query = "Select left(childAccountcode,3) as Accntcode,fundtype,accountname,accountcode,dated,date_,pariculars,jevno,debit,credit from vw_MP_Final_LedgerSubSidiary where  fundtype = '" & cmb_FundType.Text & "' and  left(childaccountcode,3) = '" & cmb_Accountcode & "' and (dated  between '" & Format(DTPicker2.Value, "MM/dd/yyyy") & "' and '" & Format(DTPicker1.Value, "MM/dd/yyyy") & "') group by left(childaccountcode,3),fundtype,accountname,dated,date_,pariculars,jevno,debit,credit,accountcode   order by accountcode"

frmSubsidiaryLedgerViewer.accnt = "SELECT min(date_) as MinDated,fundtype,sum([Debit]) as SumDebit " & _
        ",sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance " & _
        ",left(ChildAccountCode,3) as Accountcode " & _
        "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where  fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3) = '" & cmb_Accountcode & "' and dated  < '" & Format(DTPicker2.Value, "MM/dd/yyyy") & "' group by fundtype,left(childaccountcode,3)  order by left(childaccountcode,3)"
  
frmSubsidiaryLedgerViewer.maxdated = "SELECT max(date_) as MinDated,fundtype,sum([Debit]) as SumDebit " & _
        ",sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance " & _
        ",left([ChildAccountCode],3) as Accountcode " & _
        "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where  fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3) = '" & cmb_Accountcode & "' and dated  > '" & Format(DTPicker1.Value, "MM/dd/yyyy") & "' group by fundtype,left(childaccountcode,3)  order by left(childaccountcode,3)"
  
frmSubsidiaryLedgerViewer.GrndTotal = "SELECT sum([Debit]) As SumDebit,sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance" & _
        ",left([ChildAccountCode],3) as Accountcode,fundtype " & _
        "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3)  = '" & cmb_Accountcode & "' group by left(ChildAccountCode,3),fundtype"
        
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


Private Sub cmb_Accountcode_KeyPress(KeyAscii As Integer)
KeyAscii = AutoFind(cmb_Accountcode, KeyAscii, True)
End Sub

Private Sub cmb_FundType_Change()
Call loadChildAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub cmb_FundType_Click()
Call loadAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub Command1_Click()

  Call LoadSavedReport
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
End Sub

