VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmLedgerSubsidiary_Option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subsidiary Ledger"
   ClientHeight    =   9750
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   15555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "For the Period"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   3780
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   6
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
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Special Accounts:"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   120
      TabIndex        =   3
      Top             =   120
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
         ItemData        =   "frmLedgerSubsidiary_Option.frx":0000
         Left            =   195
         List            =   "frmLedgerSubsidiary_Option.frx":000D
         TabIndex        =   4
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
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "For the Period"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3780
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   200
         TabIndex        =   1
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
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
End
Attribute VB_Name = "frmLedgerSubsidiary_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub LoadSavedReport(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal FTYPE As String)
Dim opnRCINo As New ADODB.Recordset
Dim SQL As String
Dim cc As Integer
On Error GoTo bad

'----filters data from the view to finalize and to remove the data redundancy----'
'filter_data "select * from vwAMIS_CDJournal where accountname LIKE '%" & fund & "%' and (year(checkdate)='" & TrnYear & "' and month(checkdate)='" & trnMonth & "' and fndtype = '" & FTYPE & "') order by checkno,[111-1-11-WW] desc"
'--------------------------------------------------------------------------------'

frmcheckdisbursement.query = "select * from vw_MP_CDcheckdisbursement where (year(jevdate)='" & Me.DTPicker2.Year & "' and month(jevdate)='" & Me.DTPicker2.Month & "' and fundtype = '" & cmb_FundType.Text & "') order by [check no.],[111] desc"
frmcheckdisbursement.accnt = "select * from vw_MP_CDCheckdisbursement_Recap where  yr='" & TrnYear & "' and mnth='" & trnMonth & "' and fundtype = '" & FTYPE & "' order by [account code] asc "
frmcheckdisbursement.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
frmcheckdisbursement.fund = Trim(cmb_FundType.Text)
'frmcheckdisbursement.TYP = Trim(Combo2.Text)
'frmcheckdisbursement.bankno = cmb_Fund.Text & " / " & Combo1.Text
frmcheckdisbursement.Show



'opnRCINo.Open SQL, opndbaseFMIS, adOpenStatic, adLockOptimistic

'opnRCINo.Close
'Set opnRCINo = Nothing
Exit Sub
bad:
    MsgBox err.Description
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


Private Sub Command1_Click()

    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_FundType.Text)

End Sub

Private Sub Form_Load()
Call LoadFund
End Sub
Public Sub LoadFundType(ByVal cmb As ComboBox)
Dim opnfund As New ADODB.Recordset
Dim cc As Integer

opnfund.Open "Select * from tblRefBMS_Funds", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    cmb.Clear
    Do Until opnfund.EOF
        cmb.AddItem (opnfund!FundName)
        cmb.ItemData(cc) = opnfund!FundCode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub

