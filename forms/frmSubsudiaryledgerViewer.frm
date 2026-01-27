VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSubsidiaryLedgerViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   7635
   ClientLeft      =   5760
   ClientTop       =   3180
   ClientWidth     =   7140
   Icon            =   "frmSubsudiaryledgerViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   7140
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7005
      lastProp        =   500
      _cx             =   12356
      _cy             =   12356
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
      Begin VB.Menu print 
         Caption         =   "Printer Setup"
      End
   End
End
Attribute VB_Name = "frmSubsidiaryLedgerViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public query, accnt, maxdated, GrndTotal, childaccountcode, FundType, accountcode, Accountname, person, Address, dated, fund, mnth, typ, bankno As String
Public Date_ As Date
Public Whatcon As Integer
Public Rsource As Object
Private Sub Form_Load()
'On Error GoTo bad
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim Mrec As New ADODB.Recordset
Dim Trec As New ADODB.Recordset
Select Case (Report9):

Case "Subsidiary"
Me.Caption = "Subsidiary Ledger"
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
Mrec.Open (maxdated), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
Trec.Open (GrndTotal), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_LedgerSubsidiary
        With Rpt_LedgerSubsidiary
            .DiscardSavedData
            .Database.SetDataSource rs6
            .Subreport1.OpenSubreport.DiscardSavedData
            .Subreport1.OpenSubreport.DiscardSavedData
            .Subreport3.OpenSubreport.DiscardSavedData
            .txtaccountcode.SetText childaccountcode
            .txtaccountname.SetText Accountname
            .txtgeneral.SetText accountcode
            .txtperson.SetText person
            .txtdated.SetText dated
            .txtfundtype.SetText FundType
            .Subreport1.OpenSubreport.Database.SetDataSource rec
            Set Rsource = Rpt_LedgerSubsidiary
            .Subreport3.OpenSubreport.Database.SetDataSource Trec
            Call TransactionLogging("Print Preview", "Subsidiary Ledger", Me.Caption, Winsock1.LocalIP)
        End With
'End If
Case "General"
Me.Caption = "General Ledger"
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
Mrec.Open (maxdated), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
Trec.Open (GrndTotal), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_LedgerGeneral
        With Rpt_LedgerGeneral
             .DiscardSavedData

            .Database.SetDataSource rs6
            .Subreport1.OpenSubreport.DiscardSavedData
            .Subreport2.OpenSubreport.DiscardSavedData
            .Subreport3.OpenSubreport.DiscardSavedData
            .txtaccountname.SetText Accountname
            .txtgeneral.SetText accountcode
            .txtdated.SetText dated
            .txtfundtype.SetText FundType
            .Subreport1.OpenSubreport.Database.SetDataSource rec
            .Subreport2.OpenSubreport.Database.SetDataSource Mrec
            .Subreport3.OpenSubreport.Database.SetDataSource Trec
            Set Rsource = Rpt_LedgerGeneral
            'Call TransactionLogging("Print Preview", "General Ledger", Me.Caption)
            Call TransactionLogging("Print Preview", "General Ledger", Me.Caption, Winsock1.LocalIP)
        End With
Case "Trial Balance"
Me.Caption = "Trial Balance"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_TrialBalance
        With Rpt_TrialBalance
             .DiscardSavedData
            .Text12.SetText "As of " & dated
            .txtprepared.SetText GetEmpName(ActiveUserID)
            .txtposition.SetText GetEmpPosition(ActiveUserID)
            .Database.SetDataSource rs6
            Set Rsource = Rpt_TrialBalance
            'Call TransactionLogging("Print Preview", "Trial Balance", Me.Caption)
            Call TransactionLogging("Print Preview", "Trial Balance", Me.Caption, Winsock1.LocalIP)
        End With
Case "BalanceSheet"
Me.Caption = "Balance Sheet"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = CrystalReportBS
        With CrystalReportBS
             .DiscardSavedData
            .txtdate.SetText (dated)
            .Database.SetDataSource rs6
            Set Rsource = CrystalReportBS
            .txtfund.SetText FundType
            .Subreport1.OpenSubreport.DiscardSavedData
            .Subreport1.OpenSubreport.Database.SetDataSource rec
            '.txtClerk.SetText getUserName(ActiveUserID, "FullName")
            'Call TransactionLogging("Print Preview", "Balance Sheet", Me.Caption)
            
            Call TransactionLogging("Print Preview", "Balance Sheet", Me.Caption, Winsock1.LocalIP)
            End With
Case "SIE"
Me.Caption = "Statement of Income and Expense"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_SIE
        With Rpt_SIE
             .DiscardSavedData
            .txtdate.SetText (dated)
            .Database.SetDataSource rs6
            .txtfund.SetText FundType
            MsgBox query
            Set Rsource = Rpt_SIE
            'Call TransactionLogging("Print Preview", "Balance Sheet", Me.Caption)
            Call TransactionLogging("Print Preview", "Statement of Income and Expense", Me.Caption, Winsock1.LocalIP)
        End With
Case "SGE"
Me.Caption = "Statement of Government Equity"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
If rs6.RecordCount > 0 Then

End If
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = CrystalReportSGE
        With CrystalReportSGE
             .DiscardSavedData
            .txtdate.SetText (dated)
            .Database.SetDataSource rs6
            .txtfund.SetText FundType
            Set Rsource = CrystalReportSGE
            '.txtClerk.SetText getUserName(ActiveUserID, "FullName")
            'Call TransactionLogging("Print Preview", "Balance Sheet", Me.Caption)
            Call TransactionLogging("Print Preview", "Statement of Government Equity", Me.Caption, Winsock1.LocalIP)
        End With
Case "Schedules"
Me.Caption = "Schedules"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = rpt_Schedule
        With rpt_Schedule
             .DiscardSavedData
             .txtfund.SetText FundType
            .txtdate.SetText (dated)
            .Database.SetDataSource rs6
             Set Rsource = rpt_Schedule
            '.txtClerk.SetText getUserName(ActiveUserID, "FullName")
            Call TransactionLogging("Print Preview", "Schedules", Me.Caption, Winsock1.LocalIP)
            End With
Case "SCF"
Me.Caption = "Statement of Cash Flow"
'MsgBox query
rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
rec.Open "Select fmis.dbo.MPfunc_GetBeginBalForCFlow ('" & Date_ & "','" & Whatcon & "','" & FundType & "')as BeginningBal", opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = CrystalReportSCF
        With CrystalReportSCF
             .DiscardSavedData
             .txtfund.SetText FundType
             .txtdate.SetText (dated)
             .Database.SetDataSource rs6
             .Database.Tables(2).SetDataSource rec
             Set Rsource = CrystalReportSCF
        End With
        Call TransactionLogging("Print Preview", "Statement of Cash Flow", Me.Caption, Winsock1.LocalIP)
End Select
    CRViewer91.PrintReport
    CRViewer91.ViewReport
    CRViewer91.Zoom 90 'Select  [fmis].[dbo].[MPfunc_GetBeginBalForCFlow] ('1/31/2011',1,'" & FundType & "')as BeginningBal
    Screen.MousePointer = vbDefault

'Set rs6 = Nothing
'rec.Close
Exit Sub
bad:
    If err.Number = -2147467259 Then
    Unload Me
    Else
    MsgBox "Error: " & err.Description & err.Number
    End If
End Sub
Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub

Private Sub print_Click()
Rsource.PrinterSetup (hwnd)
Rsource.PaperSize = crDefaultPaperSize
UseDefault = True
CRViewer91.RefreshEx (True)
End Sub
