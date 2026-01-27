VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.MDIForm MDIFrm_MAIN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Agusan del Sur Accounting Operations and Management System"
   ClientHeight    =   9525
   ClientLeft      =   1665
   ClientTop       =   600
   ClientWidth     =   21465
   Icon            =   "MDIFrm_MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm_MAIN.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   9435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   16642
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   3735
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   6588
         Caption         =   "O||B||J||E||C||T||||E||X||P||L||O||R||E||R"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   33023
         cGradient       =   33023
         Gradient        =   1
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   2895
         Left            =   0
         TabIndex        =   3
         Top             =   3720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   5106
         Caption         =   "O||N||L||I||N||E|| ||U||S||E||R"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   33023
         cGradient       =   33023
         Gradient        =   1
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   2775
         Left            =   0
         TabIndex        =   4
         Top             =   6600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   4895
         Caption         =   "L||O||C||K||||S||Y||S||T||E||M"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   33023
         cGradient       =   33023
         Gradient        =   1
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Timer Tdoevents 
      Interval        =   65000
      Left            =   3480
      Top             =   1080
   End
   Begin VB.Timer tmeConnChck 
      Interval        =   60000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer timerWatchCursor 
      Interval        =   1000
      Left            =   840
      Top             =   480
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   21465
      TabIndex        =   0
      Top             =   9435
      Width           =   21465
   End
   Begin VB.Menu mnuEnrollment 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuTransaction 
         Caption         =   "&Transaction"
         Visible         =   0   'False
         Begin VB.Menu trnMenu 
            Caption         =   "&Incoming Transaction (DV Numbering)"
            Index           =   0
         End
         Begin VB.Menu trnMenu 
            Caption         =   "JEV& Preparation for"
            Index           =   1
            Begin VB.Menu jvold 
               Caption         =   "JEV Preparation(Old)"
               Visible         =   0   'False
            End
            Begin VB.Menu cco 
               Caption         =   "Check, Cash Disbursement and General Journal Through DVNo"
            End
            Begin VB.Menu collect 
               Caption         =   "&Collection and Deposit through PTV number"
            End
            Begin VB.Menu GJ 
               Caption         =   "General Journal Through PTV number"
            End
            Begin VB.Menu lca 
               Caption         =   "Liquidation of Cash Advance"
            End
            Begin VB.Menu New 
               Caption         =   "OLDJEvprepation"
            End
         End
         Begin VB.Menu trnMenu 
            Caption         =   "JEV &Approval and Log Out"
            Index           =   2
         End
         Begin VB.Menu trnMenu 
            Caption         =   "Log Out DV (with Approved JEV)"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu trnMenu 
            Caption         =   "JEV &Numbering"
            Index           =   4
            Begin VB.Menu JEVNoMenu 
               Caption         =   "Collection and General Journal Through PTV No."
               Index           =   0
            End
            Begin VB.Menu JEVNoMenu 
               Caption         =   "Check Disbursement"
               Index           =   1
            End
            Begin VB.Menu JEVNoMenu 
               Caption         =   "Cash Disbursement"
               Index           =   2
            End
            Begin VB.Menu JEVNoMenu 
               Caption         =   "General Journal Through Credit Notice No."
               Index           =   3
            End
         End
         Begin VB.Menu trnMenu 
            Caption         =   "Prepare Accnt Advice"
            Index           =   5
         End
         Begin VB.Menu fjv 
            Caption         =   "Final Journal Entry Voucher"
         End
         Begin VB.Menu CNE 
            Caption         =   "Credit Notice Entry"
         End
         Begin VB.Menu DT 
            Caption         =   "Disapprove Transaction"
         End
         Begin VB.Menu son 
            Caption         =   "Search Transaction Details Through Obr No."
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaintenance 
         Caption         =   "&Maintenance"
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu MMenu 
            Caption         =   "Chart of Accounts"
            Index           =   0
         End
         Begin VB.Menu MMenu 
            Caption         =   "User Accounts"
            Index           =   1
         End
         Begin VB.Menu MC 
            Caption         =   "Manage Connection"
         End
         Begin VB.Menu s 
            Caption         =   "Signatory"
         End
         Begin VB.Menu acc 
            Caption         =   "Account Code Classification"
         End
         Begin VB.Menu trf 
            Caption         =   "Table Relator for Chart of Accounts"
         End
         Begin VB.Menu ct 
            Caption         =   "Conversion table"
         End
         Begin VB.Menu IBB 
            Caption         =   "Insert Beginning Balance"
         End
         Begin VB.Menu CFC 
            Caption         =   "Cash Flow Classification"
         End
         Begin VB.Menu ORC 
            Caption         =   "Other Responsibility Center"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUtilities 
         Caption         =   "Utilities"
         Visible         =   0   'False
         Begin VB.Menu UtiMenu 
            Caption         =   "DBase Utility"
            Index           =   0
         End
         Begin VB.Menu UtiMenu 
            Caption         =   "Log &History Viewer"
            Index           =   1
         End
         Begin VB.Menu UtiMenu 
            Caption         =   "Locate Transaction"
            Index           =   2
         End
         Begin VB.Menu UtiMenu 
            Caption         =   "Daily Accomplishment"
            Index           =   3
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Logmenu 
         Caption         =   "Log/Exit Application"
         Begin VB.Menu shutmenu 
            Caption         =   "Log &In"
            Index           =   0
         End
         Begin VB.Menu shutmenu 
            Caption         =   "Log &Out"
            Index           =   1
         End
         Begin VB.Menu shutmenu 
            Caption         =   "E&xit Application"
            Index           =   2
         End
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu view 
      Caption         =   "&Query"
      Visible         =   0   'False
      Begin VB.Menu sop 
         Caption         =   "Status of Appropriation"
      End
      Begin VB.Menu CI 
         Caption         =   "Check Issued"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Visible         =   0   'False
      Begin VB.Menu repmenu 
         Caption         =   "&Customer Accounts Report"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu repmenu 
         Caption         =   "&List of Customers"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu repmenu 
         Caption         =   "C&ustomer Information (Profile)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu repmenu 
         Caption         =   "Customer's Amortization Schedule"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu repmenu 
         Caption         =   "&Unsettled Accounts/Balances"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu jor 
         Caption         =   "Journals"
         Begin VB.Menu cadr 
            Caption         =   "Cash Disbursements Report"
         End
         Begin VB.Menu cdr 
            Caption         =   "Check Disbursement Report"
         End
         Begin VB.Menu samp 
            Caption         =   "sample"
            Visible         =   0   'False
         End
         Begin VB.Menu cRJ 
            Caption         =   "Cash Receipts Journal Report"
         End
         Begin VB.Menu gjr 
            Caption         =   "General Journal Report"
         End
      End
      Begin VB.Menu ledger 
         Caption         =   "Ledger"
         Begin VB.Menu SL 
            Caption         =   "Subsidiary Ledger"
         End
         Begin VB.Menu GL 
            Caption         =   "General Ledger"
         End
      End
      Begin VB.Menu fr 
         Caption         =   "Financial Reports"
         Begin VB.Menu TB 
            Caption         =   "Trial Balance"
         End
         Begin VB.Menu SIE 
            Caption         =   "Statement of Income and Expenses"
         End
         Begin VB.Menu SGE 
            Caption         =   "Statement of Government Equity"
         End
         Begin VB.Menu BS 
            Caption         =   "Balance Sheet"
         End
         Begin VB.Menu sched 
            Caption         =   "Schedules"
         End
         Begin VB.Menu scf 
            Caption         =   "Statement of Cash flow"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSystemDev 
         Caption         =   "&System Developers"
      End
      Begin VB.Menu mnu11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About MCAIS..."
      End
      Begin VB.Menu ssssss 
         Caption         =   "-"
      End
      Begin VB.Menu update 
         Caption         =   "Update"
      End
   End
   Begin VB.Menu pop 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu payroll 
         Caption         =   "Payroll"
      End
      Begin VB.Menu property 
         Caption         =   "Property"
      End
   End
End
Attribute VB_Name = "MDIFrm_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SW As Long
Dim x As Long
Dim Hyd As Boolean
Private Sub acc_Click()
frm_AccountcodeSub.Show
End Sub

Private Sub AD_Click()
frmJEVDisapprove.Show 1

End Sub

Private Sub BS_Click()
frm_balancesheet.Show
End Sub

Private Sub cadr_Click()
frmcashdisbursement_Option.Left = 6300
frmcashdisbursement_Option.Top = 3945
frmcashdisbursement_Option.Show
End Sub

Private Sub cco_Click()
frmJEVPreparation_New.Show 1
End Sub

Private Sub cdr_Click()
frmcheckdisbursement_Option.Left = 6300
frmcheckdisbursement_Option.Top = 3945
frmcheckdisbursement_Option.Show
End Sub

Private Sub CFC_Click()
With frm_cashFlowClass
.col = 0
.Field1 = "Subcode1"
.Field2 = "Subdesc1"
.Condition = "Subcode1 IS NOT NULL"
 .Frame1.Visible = True
.Show 1
End With
End Sub

Private Sub CI_Click()
frmVw_CheckIssued.Show
End Sub
Private Sub CNE_Click()
frmCreditNoticeEntry.Show vbModal
End Sub

Private Sub collect_Click()
frmJEVPreparationforColection_New.Show vbModal
End Sub

Private Sub crj_Click()
frmcashCashReceipts_Option.Left = 6300
frmcashCashReceipts_Option.Top = 3945
frmcashCashReceipts_Option.Show
End Sub

Private Sub ct_Click()
frmOtherClass.Show 1
End Sub

Private Sub DT_Click()
frmJEVDisapprove.Show 1
End Sub

Private Sub fjv_Click()
frmFinalJev.Show
End Sub

Private Sub GJ_Click()
frmJEVPreparationforGeneralJournal_New.Show 1
End Sub

Private Sub gjr_Click()
frmGeneralJOurnal_option.Show
End Sub

Private Sub GL_Click()
frmLedgerGeneral.Show
End Sub

Private Sub IBB_Click()
frm_BeginBeginAccountcodeSub.Show 1
End Sub

Private Sub JEVNoMenu_Click(Index As Integer)
Select Case Index
    Case 0
        frmCDCashReceiptsJevNumbering.Show 1
    Case 1
        frmJEVNumberingThruRCI.Show
        'frmJEVNumbering.Show
    Case 2
        frmCDCashDisbursedReport.Show vbModal
    Case 3
        frmGeneralJournalJevNumbering.Show vbModal
End Select
End Sub

Private Sub jvold_Click()
'frmJEVPreparation.Show 1
End Sub

Private Sub lca_Click()
frmJEVPreparationfor_Liquidation.Show 1
End Sub

Private Sub Logmenu_Click()
'Dim cc As Variant

'cc = InputBox("Enter UserID:", "User Log Registry")
'If Len(cc) = 4 Then
'    ActiveUserID = cc
'    MDIFrm_MAIN.Caption = "Agusan del Sur Accounting Operations and Management System " & "(" & ActiveUserID & ")"
'Else
'    Exit Sub
'End If
End Sub

Private Sub lvButtons_H1_Click()
On Error Resume Next
With frm_toolwindows
If frm_toolwindows.Width = 0 Then
.ZOrder (0)
.Top = 0
.Left = 0
Dim a As Long
For a = 1 To SW
.Height = Me.Height - 1000
.Width = .Width + 1
'DoEvents
Next a
Else
    If .Width > 1 Then
        .ZOrder (0)
    End If
End If
End With
End Sub

Private Sub lvButtons_H3_Click()
frmLock.Show 1
End Sub

Private Sub MC_Click()
frm_UtilityConnection.Show
End Sub

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    SW = frm_toolwindows.Width
    DefaultPost = Now
End Sub
Public Sub disableEnabletimer(ByVal bol As Boolean)
tmeConnChck.Enabled = bol
timerWatchCursor.Enabled = bol
Tdoevents.Enabled = bol
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim retval As Long  ' return value of the function

On Error Resume Next

If MsgBox("Are you sure want to Exit this Application?", vbQuestion + vbYesNo, "System Confirmation Query") <> vbYes Then
    Cancel = 1
Else
    SndLocation = readTXTDATA("Sound", "Extro", App.path & "\data\SystemDefault.ini")
    'Call DisplayChangeSetting(Res_Width, Res_Height, "Exit")
    retval = PlaySound(SndLocation, 0, SND_ALIAS Or SND_SYNC)
    Call OnlineDeleteLogging
End If
End Sub
Private Sub moo_Click()
frmIncomingTrn.Show 1
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
With frm_toolwindows
.Top = 0
.Left = 0
'.Width = 5000
.Height = Me.Height - 1000
End With
End Sub

Private Sub mnuSystemDev_Click()
frm_celemoc_import.Show
End Sub

Private Sub New_Click()
'frmJEVPreparation.Show
End Sub

Private Sub oon_Click()
frmIncomingTrn.Show 1
End Sub

Private Sub ORC_Click()
frm_OtherRC.Show
End Sub

Private Sub s_Click()
frm_Signatory.Show 1
End Sub

Private Sub scf_Click()
frm_SCF.Show
End Sub

Private Sub sched_Click()
frm_Schedule.Show
End Sub

Private Sub SGE_Click()
frm_SGE.Show
End Sub

Private Sub SIE_Click()
frm_SIE.Show
End Sub

Private Sub SL_Click()

frmLedgerSubsidiary.Show
End Sub

'Private Sub timerWatchCursor_Timer()
'
'    Static iC As Integer
'    Static op As POINTAPI
'
'    Dim p As POINTAPI
'
'    GetCursorPos p
'    If (p.x < (op.x + 5) And p.x > op.x - 5) And (p.Y < (op.Y + 5) And p.Y > op.Y - 5) Then
'        iC = iC + 1
'    Else
'        iC = 0
'    End If
'
'    op.x = p.x
'    op.Y = p.Y
'
'    If iC > AppSet_LockTimeOut Then
'        iC = 0
'        Call LockApp
'    End If
'End Sub
Public Function LockApp()
    frmLock.ShowForm
End Function
Private Sub MDIForm_Unload(Cancel As Integer)
opndbasePMIS.Close
opndbaseFMIS.Close
Set opndbasePMIS = Nothing
Set opndbaseFMIS = Nothing
End
End Sub

Private Sub MMenu_Click(Index As Integer)
If Len(Trim(ActiveUserID)) > 0 Then
    Select Case Index
        Case 0 'Chart of Accounts
            frmChartOfAccounts.Show vbModal
        Case 1 'Users
            frmSystemUsers_new.Show
    End Select
Else
    MsgBox "Please Log-In First, before you can use the system!", vbCritical, "System Warning"
End If
End Sub



Private Sub samp_Click()
Form2.Show
End Sub

Private Sub shutmenu_Click(Index As Integer)
Select Case Index
    Case 0 'Log In
        'If MsgBox("Are You Sure, Want to Log In?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
            
            'here write "Log In" to Log.ini----------------------
            'log in happens only during the Log In Event in the UserPassword Form
            '---------------------------------------------------
            
            ShutDownMode = "Log In System"
            frmUserPassword.Show
            Sleep 500
            Unload frmUserPassword
            Sleep 500
            frmUserPassword.Show
        'End If
    
    Case 1 'Log Out
        'If MsgBox("Are You Sure, You Want to Log Out?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
            
            'here write "Log Out" to Log.ini----------------------
            Call WriteLogOn
            Log = "Out"
            '---------------------------------------------------
            
            Call DisableAllMenus 'Setting all menu to disable
            
            'ReSetting the Status Bar Panels ----------------------------
            frmLogInPicture.img_userpic.Visible = False
            
            ActiveUser = ""
            ActiveUserID = ""
            'frmMother.StatusBar1.Panels(2) = ""
            'frmMother.StatusBar1.Panels(3) = ""
            'frmMother.StatusBar1.Panels(4) = ""
            
            Call VerifyLog
            Unload frmLogInPicture
        'End If
    
    Case 2 'Exit application
        Unload MDIFrm_MAIN
End Select
End Sub
Private Sub WriteLogOn()

'Writing Log Out to History ------------------------
Open LogLocation & "\ActivityLog" & Month(Date) & Year(Date) & ".ini" For Append As #1
Print #1, ActiveUserID & Chr(9) & UserLevel & Chr(9) & GetPCName & Chr(9) & GetOfficeIDbyUserID(ActiveUserID) & Chr(9) & Format(Date, "mmm dd,yyyy") & Chr(9) & Time & Chr(9) & "OUT"
Close #1

End Sub



Private Sub son_Click()
frm_transdetails.Show
End Sub

Private Sub sop_Click()
frmStatOfAppro.Show 1
End Sub

Private Sub TB_Click()
frmTrialBalance.Show
End Sub

Private Sub Tdoevents_Timer()
'Dim rec As New ADODB.Recordset
'Dim Vr As String
'
'If UpdateStat = 0 Then
'    Vr = App.Major & "." & App.Minor & "." & App.Revision
'    Set rec = opndbaseFMIS.Execute("SELECT top 1 [Version] FROM [fmis].[dbo].[tblAMIS_SystemUpdate]")
'    If rec.RecordCount > 0 And Trim(rec!Version) <> Vr Then
'        With frm_updatedesc
'        '.ZOrder (0)
'        .Show
'        End With
'    End If
'    rec.Close
'    Set rec = Nothing
'End If
End Sub

Private Sub tmeConnChck_Timer()
On Error GoTo conn
Dim OpnConn As New ADODB.Connection
'if OpnConn.State =
opndbaseFMIS.Execute ("Select trnno from tblAMIS_SystemUpdate where ISID = 1")
Exit Sub
conn:
    tmeConnChck.Enabled = False
    frmConnCheck.Show 1
    tmeConnChck.Enabled = True

End Sub

Private Sub TreeView1_GotFocus()
If Picture1.Width = 15 Then
For x = 1 To SW
  Picture1.Width = Picture1.Width + 1
  DoEvents
Next x
End If
End Sub

Private Sub TreeView1_LostFocus()
On Error Resume Next
If Picture1.Width > 15 Then
For x = 1 To SW
  Picture1.Width = Picture1.Width - x
  DoEvents
Next x
End If
End Sub

Private Sub trf_Click()
frm_relatedtableForCOA.Show 1
End Sub

Private Sub trnMenu_Click(Index As Integer)

If Len(Trim(ActiveUserID)) > 0 Then

    Select Case Index
        Case 0 'Incoming TRansaction REgistry
            frmIncomingTrn.Show vbModal
        Case 1 'JEV Preparation
            'frmJEVPreparation.Show vbModal
        Case 2 'JEV Approval
            frmJEVApproval.Show vbModal
        Case 3 'DV (with Approved JEV) Log out
            frmApprovedJEVLogOut.Show vbModal
        Case 5 'Accountant Advice Preparation
            frmAccountantsAdvice.Show vbModal
    End Select
Else
    MsgBox "Please Log-In First, Before you can use the System!", vbCritical, "System Warning"
End If
End Sub

Private Sub update_Click()
If MsgBox("Are you sure do you want to update the system? the system will close Automatically if do you want to proceed..", vbCritical + vbYesNo, "System Messgae") = vbYes Then
Shell App.path & "\Update.exe", vbNormalFocus
End
End If
End Sub

Private Sub UtiMenu_Click(Index As Integer)
Select Case Index
    Case 0 'Database Utility
            frmDataUtility.Show vbModal
    Case 1 'Log History Viewer
            FrmLogInOutHistory.Show vbModal
    Case 2 'Locate Transaction
            frmDVSearch.Show 1
    Case 3  'Daily Accomplishment
            frmAccomplishment.Show 1
End Select

End Sub

