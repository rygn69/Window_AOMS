VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_OnlineUser 
   BorderStyle     =   0  'None
   Caption         =   "Object Explorer"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   DrawStyle       =   5  'Transparent
   Icon            =   "frm_OnlineUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Hide"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cBhover         =   8438015
      LockHover       =   1
      cGradient       =   8438015
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_OnlineUser.frx":1272
      cBack           =   -2147483633
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Autohide"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Timer tmrTv 
      Left            =   600
      Top             =   720
   End
   Begin MSComctlLib.ImageList imgNode 
      Left            =   240
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":4D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":59CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":6620
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":7272
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":7EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":8B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":9768
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":D272
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_OnlineUser.frx":D9EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Online Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   7800
   End
End
Attribute VB_Name = "frm_OnlineUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xSelNodeId As String
Private Sub Form_Load()
Call loadUsernode
End Sub
Public Sub loadUsernode()
Dim sqlQry  As String
Dim xNode As Node
Dim xRsTemp1 As New ADODB.Recordset
Dim xParent As String

TrvObject.Nodes.Clear


''Set xNode = TrvObject.Nodes.Add(, , "a", "Main Root", 1)
'xNode.Expanded = True
sqlQry = "SELECT a.[UserID],UserName FROM [fmis].[dbo].[tblAMIS_OnlineUser] as a inner join fmis.dbo.[tblAMIS_UserRegistry] as b on a.UserID = b.UserID order by Username" ' where nodeLevel=1"
xRsTemp1.Open sqlQry, opndbaseFMIS, adOpenDynamic, adLockOptimistic
With xRsTemp1
    If Not .BOF Or Not .EOF Then
        Do While Not .EOF
               
        .MoveNext
        Loop
    End If
    .Close
End With
End Sub
Private Sub lvButtons_H1_Click()
Call loadUsernode
End Sub

Private Sub List1_Click()

End Sub

Private Sub lvButtons_H2_Click()
Dim x As Long
On Error Resume Next
If Me.Width > 1 Then
For x = 1 To Me.Width
Me.Width = Me.Width - 1
DoEvents
Next x
End If
End Sub

Private Sub lvButtons_H3_Click()
If lvButtons_H3.Caption = "Log-out" Then
    If MsgBox("Are you sure do you want to log out?", vbInformation + vbYesNo, "System Message") = vbYes Then
    Call TransactionLogging("Log Out", "tblAMIS_log", Me.Caption, Winsock1.LocalIP)
    lvButtons_H3.Caption = "Log-In"
    ActiveUser = ""
    ActiveUserID = ""
    ActiveUserPass = ""
    TrvObject.Nodes.Clear
    frmUserPassword.Show 1
    End If
Else
frmUserPassword.Show 1
End If
End Sub

Private Sub tmrTv_Timer()
    '********************************************
    '* treeview will be refresh after 2 seconds *
    '********************************************
    xTimerInterval = xTimerInterval + 1
    If xTimerInterval = 2 Then
        ' Call local_fillTreeView
        txtNode.Text = ""
        tmrTv.Interval = 0
        xTimerInterval = 0
    End If
End Sub

Private Sub TrvObject_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Children <> 0 Then
        Node.Image = 3
    End If
End Sub


Private Sub TrvObject_DblClick()
If xSelNodeId = "a5" Then
    centerme frmIncomingTrn
    frmIncomingTrn.ZOrder (0)
    frmIncomingTrn.Show

ElseIf xSelNodeId = "a9" Then
    centerme frmJEVPreparation_New
    frmJEVPreparation_New.ZOrder (0)
    frmJEVPreparation_New.Show

ElseIf xSelNodeId = "a19" Then
    centerme frmFinalJev
    frmFinalJev.ZOrder (0)
    frmFinalJev.Show

ElseIf xSelNodeId = "a10" Then
    centerme frmJEVPreparationforColection_New
    frmJEVPreparationforColection_New.ZOrder (0)
    frmJEVPreparationforColection_New.Show

ElseIf xSelNodeId = "a11" Then
    centerme frmJEVPreparationfor_Liquidation
    frmJEVPreparationfor_Liquidation.ZOrder (0)
    frmJEVPreparationfor_Liquidation.Show

ElseIf xSelNodeId = "a12" Then
    centerme frm_FinalLogOut
    frm_FinalLogOut.ZOrder (0)
    frm_FinalLogOut.Show

ElseIf xSelNodeId = "a14" Then
    centerme frmCDCashReceiptsJevNumbering
    frmCDCashReceiptsJevNumbering.ZOrder (0)
    frmCDCashReceiptsJevNumbering.Show

ElseIf xSelNodeId = "a15" Then
    centerme frmJEVNumberingThruRCI
    frmJEVNumberingThruRCI.ZOrder (0)
    frmJEVNumberingThruRCI.Show

ElseIf xSelNodeId = "a16" Then
    centerme frmCDCashDisbursedReport
    frmCDCashDisbursedReport.ZOrder (0)
    frmCDCashDisbursedReport.Show

ElseIf xSelNodeId = "a18" Then
    centerme frmAccountantsAdvice
    frmAccountantsAdvice.ZOrder (0)
    frmAccountantsAdvice.Show

ElseIf xSelNodeId = "a20" Then
    centerme frmJEVDisapprove
    frmJEVDisapprove.ZOrder (0)
    frmJEVDisapprove.Show

ElseIf xSelNodeId = "a21" Then
'    CenterMe frm_transdetails
    frm_transdetails.ZOrder (0)
    frm_transdetails.Show

ElseIf xSelNodeId = "a22" Then
    centerme frmSystemUsers_new
    frmSystemUsers_new.ZOrder (0)
    frmSystemUsers_new.Show

ElseIf xSelNodeId = "a23" Then
    centerme frm_UtilityConnection
    frm_UtilityConnection.ZOrder (0)
    frm_UtilityConnection.Show

ElseIf xSelNodeId = "a24" Then
    centerme frm_Signatory
    frm_Signatory.ZOrder (0)
    frm_Signatory.Show

ElseIf xSelNodeId = "a25" Then
    centerme frm_AccountcodeSub
    frm_AccountcodeSub.ZOrder (0)
    frm_AccountcodeSub.Show

ElseIf xSelNodeId = "a26" Then
    centerme frm_relatedtableForCOA
    frm_relatedtableForCOA.ZOrder (0)
    frm_relatedtableForCOA.Show

ElseIf xSelNodeId = "a27" Then
    centerme frmOtherClass
    frmOtherClass.ZOrder (0)
    frmOtherClass.Show

ElseIf xSelNodeId = "a28" Then
    centerme frm_BeginBeginAccountcodeSub
    frm_BeginBeginAccountcodeSub.ZOrder (0)
    frm_BeginBeginAccountcodeSub.Show

ElseIf xSelNodeId = "a29" Then
    
    With frm_cashFlowClass
    .col = 0
    .Field1 = "Subcode1"
    .Field2 = "Subdesc1"
    .Condition = "Subcode1 IS NOT NULL"
     .Frame1.Visible = True
    .Show 1
    End With

ElseIf xSelNodeId = "a30" Then
    centerme frmDataUtility
    frmDataUtility.ZOrder (0)
    frmDataUtility.Show
    
ElseIf xSelNodeId = "a31" Then
    centerme FrmLogInOutHistory
    FrmLogInOutHistory.ZOrder (0)
    FrmLogInOutHistory.Show
    
ElseIf xSelNodeId = "a32" Then
    centerme frmDVSearch
    frmDVSearch.ZOrder (0)
    frmDVSearch.Show
    
ElseIf xSelNodeId = "a33" Then
    centerme frmAccomplishment
    frmAccomplishment.ZOrder (0)
    frmAccomplishment.Show
    
ElseIf xSelNodeId = "a35" Then
    centerme frmStatOfAppro
    frmStatOfAppro.ZOrder (0)
    frmStatOfAppro.Show
ElseIf xSelNodeId = "a38" Then
    
    centerme frmcashdisbursement_Option
    frmcashdisbursement_Option.ZOrder (0)
    frmcashdisbursement_Option.Show

ElseIf xSelNodeId = "a39" Then
    centerme frmcheckdisbursement_Option
    frmcheckdisbursement_Option.ZOrder (0)
    frmcheckdisbursement_Option.Show
    
ElseIf xSelNodeId = "a40" Then
    centerme frmcashCashReceipts_Option
    frmcashCashReceipts_Option.ZOrder (0)
    frmcashCashReceipts_Option.Show

ElseIf xSelNodeId = "a41" Then
    centerme frm_GeneralJournal_Option
    frm_GeneralJournal_Option.ZOrder (0)
    frm_GeneralJournal_Option.Show
    
ElseIf xSelNodeId = "a43" Then
    centerme frmLedgerSubsidiary
    frmLedgerSubsidiary.ZOrder (0)
    frmLedgerSubsidiary.Show
    
ElseIf xSelNodeId = "a44" Then
    centerme frmLedgerGeneral
    frmLedgerGeneral.ZOrder (0)
    frmLedgerGeneral.Show

ElseIf xSelNodeId = "a46" Then
    centerme frmTrialBalance
    frmTrialBalance.ZOrder (0)
    frmTrialBalance.Show
    
ElseIf xSelNodeId = "a47" Then
    centerme frm_SIE
    frm_SIE.ZOrder (0)
    frm_SIE.Show

ElseIf xSelNodeId = "a48" Then
    centerme frm_SGE
    frm_SGE.ZOrder (0)
    frm_SGE.Show
    
ElseIf xSelNodeId = "a49" Then
    centerme frm_balancesheet
    frm_balancesheet.ZOrder (0)
    frm_balancesheet.Show
    
ElseIf xSelNodeId = "a50" Then
    centerme frm_Schedule
    frm_Schedule.ZOrder (0)
    frm_Schedule.Show
    
ElseIf xSelNodeId = "a51" Then
    centerme frm_SCF
    frm_SCF.ZOrder (0)
    frm_SCF.Show
    
ElseIf xSelNodeId = "a52" Then
'    CenterMe frmVw_CheckIssued
    frmVw_CheckIssued.ZOrder (0)
    frmVw_CheckIssued.Show
ElseIf xSelNodeId = "a53" Then
    centerme frm_AccountsPayableEntry
    frm_AccountsPayableEntry.ZOrder (0)
    frm_AccountsPayableEntry.Show
ElseIf xSelNodeId = "a54" Then
    centerme frm_COAQueryGenerator
    frm_COAQueryGenerator.ZOrder (0)
    frm_COAQueryGenerator.Show
ElseIf xSelNodeId = "a55" Then
    centerme frm_SigAccom
    frm_SigAccom.ZOrder (0)
    frm_SigAccom.Show
ElseIf xSelNodeId = "a56" Then
    centerme frm_FindTransThroughQuery
    frm_FindTransThroughQuery.ZOrder (0)
    frm_FindTransThroughQuery.Show
ElseIf xSelNodeId = "a57" Then
    centerme frm_GeneralJournalJevNumbering
    frm_GeneralJournalJevNumbering.ZOrder (0)
    frm_GeneralJournalJevNumbering.Show
ElseIf xSelNodeId = "a58" Then
    centerme frmJEVPreparationforAjustment_new
    frmJEVPreparationforAjustment_new.ZOrder (0)
    frmJEVPreparationforAjustment_new.Show
ElseIf xSelNodeId = "a59" Then
    centerme frmJEVPreparationforMemoEntry
    frmJEVPreparationforMemoEntry.ZOrder (0)
    frmJEVPreparationforMemoEntry.Show
ElseIf xSelNodeId = "a60" Then
    centerme frm_StatOfCashAdvance
    frm_StatOfCashAdvance.ZOrder (0)
    frm_StatOfCashAdvance.Show
ElseIf xSelNodeId = "a61" Then
    centerme frm_JEVApproval
    frm_JEVApproval.ZOrder (0)
    frm_JEVApproval.Show
ElseIf xSelNodeId = "a62" Then
    With frm_RRR_Maintainance
    .col = 0
    .Field1 = "Subcode1"
    .Field2 = "Subdesc1"
    .Condition = "Subcode1 IS NOT NULL"
    .Frame1.Visible = True
    .Show 1
    End With
ElseIf xSelNodeId = "a63" Then
    With frm_bankRClass
    .col = 0
    .Field1 = "Subcode1"
    .Field2 = "Subdesc1"
    .Condition = "Subcode1 IS NOT NULL"
    .Frame1.Visible = True
    .Show 1
    End With
ElseIf xSelNodeId = "a64" Then
    centerme frm_BankReconciliation
    frm_BankReconciliation.ZOrder (0)
    frm_BankReconciliation.Show
ElseIf xSelNodeId = "a65" Then
    centerme frmJEVPreparationforColection_byJEVno
    frmJEVPreparationforColection_byJEVno.ZOrder (0)
    frmJEVPreparationforColection_byJEVno.Show
End If
End Sub

Private Sub TrvObject_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Children = 0 Then
        'Node.Image = 4
        Else
        Node.Image = 2
    End If
End Sub


Private Sub TrvObject_LostFocus()
Dim x As Long
On Error Resume Next
If Check1.Value = 1 Then
If Me.Width > 1 Then
For x = 1 To Me.Width
Me.Width = Me.Width - 1
DoEvents
Next x
End If
End If
End Sub

Private Sub TrvObject_NodeClick(ByVal Node As MSComctlLib.Node)
    xSelNodeId = Node.Key
End Sub

Private Sub Form_Resize()
With TrvObject
    .Height = Me.Height - 700
    Line1.Y2 = Me.Height
    If .Width < 300 Then
        If Me.Width = 0 Then
        .Width = Me.Width - 300
        
        End If
    End If
End With
End Sub
Private Function centerme(ByVal frm As Form)
Dim H, w, FW, FFW, FH, FFH, x, y As Long
frm.ScaleMode = 5
H = MDIFrm_MAIN.Height
FH = frm.Height
x = frm.ScaleHeight / 2
FFH = (H - FH) / x

w = MDIFrm_MAIN.Width
y = frm.ScaleWidth / 2
FW = frm.Width
FFW = (w - FW)

frm.Top = FFH / 2
frm.Left = FFW / 2
End Function
