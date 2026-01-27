VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frma_RPTfinancialViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   7410
   ClientLeft      =   5760
   ClientTop       =   3180
   ClientWidth     =   7125
   Icon            =   "frma_RPTfinancialViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   7125
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7005
      lastProp        =   600
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
      EnableLogonPrompts=   -1  'True
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
      Begin VB.Menu page 
         Caption         =   "Page Setup"
      End
   End
End
Attribute VB_Name = "frma_RPTfinancialViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public query, accnt, maxdated, GrndTotal, childaccountcode, FundType, accountcode, Accountname, person, Address, dated, fund, mnth, TYP, bankno As String
Public Date_ As Date
Public ComparativeDate As String
Public Whatcon, Conso As Integer
Public preclosing, AllFunds, Comparative As Boolean 'preclosing = true and postclosing = false
Public Rsource As Object
Public IEE, IEE100, IEE200, IEE300 As Currency
Private Sub Form_Load()
'On Error GoTo bad
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim Mrec As New ADODB.Recordset
Dim Trec As New ADODB.Recordset
Dim SRec As New ADODB.Recordset

Dim objCommand As ADODB.command
Set objCommand = New ADODB.command
objCommand.CommandTimeout = 0
objCommand.ActiveConnection = opndbaseFMIS

Dim objCommand2 As ADODB.command
Set objCommand2 = New ADODB.command
objCommand2.CommandTimeout = 0
objCommand2.ActiveConnection = opndbaseFMIS

opndbaseFMIS.Execute "Execute MPproc_NeedToExecute @type = 1"
Select Case (Report9):
Case "Subsidiary"
Me.Caption = "Subsidiary Ledger"
'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic





objCommand.CommandText = query '
'MsgBox query
Set rs6 = objCommand.Execute


'rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'Mrec.Open (maxdated), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'Trec.Open (GrndTotal), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
            'If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_Ledger_Subsidiary
        With Rpt_Ledger_Subsidiary
            .DiscardSavedData
            .Database.SetDataSource rs6
            .txtaccountcode.SetText childaccountcode
            .txtaccountname.SetText Accountname
            .txtgeneral.SetText accountcode
            .txtperson.SetText person
            .txtdated.SetText dated
            .txtfundtype.SetText FundType
             Call TransactionLogging("Print Preview", "Subsidiary Ledger", Me.Caption, Winsock1.LocalIP)
        End With
            'End If
            '.Subreport1.OpenSubreport.Database.SetDataSource rec
            'Set Rsource = Rpt_LedgerSubsidiary
            ' .Subreport3.OpenSubreport.Database.SetDataSource Trec
            '.Subreport1.OpenSubreport.DiscardSavedData
            '.Subreport1.OpenSubreport.DiscardSavedData
            '.Subreport3.OpenSubreport.DiscardSavedData
Case "General"
Me.Caption = "General Ledger"
'Set rs6 = opndbaseFMIS.Execute(query)

objCommand.CommandText = query
Set rs6 = objCommand.Execute
'rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'Mrec.Open (maxdated), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'Trec.Open (GrndTotal), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Rpt_LedgerGeneral
        With Rpt_LedgerGeneral
             .DiscardSavedData

            .Database.SetDataSource rs6
            .txtaccountname.SetText Accountname
            .txtgeneral.SetText accountcode
            .txtdated.SetText dated
            .txtfundtype.SetText FundType
            Set Rsource = Rpt_LedgerGeneral
            'Call TransactionLogging("Print Preview", "General Ledger", Me.Caption)
            Call TransactionLogging("Print Preview", "General Ledger", Me.Caption, Winsock1.LocalIP)
        End With
Case "Trial Balance"
Me.Caption = "Trial Balance"
'MsgBox query
'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
objCommand.CommandText = query
Set rs6 = objCommand.Execute

    Screen.MousePointer = vbHourglass
        If AllFunds = True Then 'all Funds
            CRViewer91.ReportSource = Rpt_TrialBalanceALLFunds
            With Rpt_TrialBalanceALLFunds
                .DiscardSavedData
                If preclosing = True Then
                    .Text11.SetText "Consolidated Pre-Closing Trial Balance"
                Else
                    .Text11.SetText "Consolidated Post-Closing Trial Balance"
                End If
                .Text12.SetText "As of " & dated
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText GetEmpPosition(ActiveUserID)
                .Database.SetDataSource rs6
                .txtFund.SetText FundType
                Set Rsource = Rpt_TrialBalanceALLFunds
            End With
        Else 'by Funds
            CRViewer91.ReportSource = Rpt_TrialBalance
            With Rpt_TrialBalance
                .DiscardSavedData
                If preclosing = True Then
                    .Text11.SetText "Pre-Closing Trial Balance"
                Else
                    .Text11.SetText "Post-Closing Trial Balance"
                End If
                .Text12.SetText "As of " & dated
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText GetEmpPosition(ActiveUserID)
                .Database.SetDataSource rs6
                .txtFund.SetText FundType
                Set Rsource = Rpt_TrialBalance
            End With
        End If
        Call TransactionLogging("Print Preview", "Trial Balance", Me.Caption, Winsock1.LocalIP)
Case "BalanceSheet"
Me.Caption = "Balance Sheet"
'MsgBox accnt

objCommand.CommandText = query
Set rs6 = objCommand.Execute

objCommand2.CommandText = accnt
Set rec = objCommand2.Execute

'
'Set rs6 = opndbaseFMIS.Execute(query)
'Set rec = opndbaseFMIS.Execute(accnt)
'MsgBox rs6.RecordCount
'MsgBox rec.RecordCount
    If Comparative = True Then
         Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrystalReportBS_Comparative
            With CrystalReportBS_Comparative
                .DiscardSavedData
                .txtDate.SetText (dated)
                .Database.SetDataSource rs6
                 Set Rsource = CrystalReportBS_Comparative
                .txtFund.SetText FundType
                .txtCurrentYear.SetText ComparativeDate
                .txtPreviousYear.SetText (ComparativeDate - 1)
                .txtComparativeYear.SetText "(With Comparative Figures from CY " & (ComparativeDate - 1) & ")"
                .Subreport1.OpenSubreport.DiscardSavedData
                .Subreport1.OpenSubreport.Database.SetDataSource rec
                Call TransactionLogging("Print Preview", "Balance Sheet Comparative Report", Me.Caption, Winsock1.LocalIP)
            End With
    Else
        If AllFunds = True Then 'all Funds
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrystalReportBS_AllFunds
            With CrystalReportBS_AllFunds
                .DiscardSavedData
                .txtDate.SetText (dated)
                .Database.SetDataSource rs6
                 Set Rsource = CrystalReportBS_AllFunds
                .txtFund.SetText FundType
                .Subreport1.OpenSubreport.DiscardSavedData
                .Subreport1.OpenSubreport.Database.SetDataSource rec
                Call TransactionLogging("Print Preview", "Balance Sheet All Funds", Me.Caption, Winsock1.LocalIP)
            End With
        Else
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrystalReportBS
            With CrystalReportBS
                .DiscardSavedData
                .txtDate.SetText (dated)
                .Database.SetDataSource rs6
                 Set Rsource = CrystalReportBS
                .txtFund.SetText FundType
                .Subreport1.OpenSubreport.DiscardSavedData
                .Subreport1.OpenSubreport.Database.SetDataSource rec
                Call TransactionLogging("Print Preview", "Balance Sheet By Funds", Me.Caption, Winsock1.LocalIP)
            End With
        End If
    End If
    rec.Close
    Set rec = Nothing
    rs6.Close
    Set rs6 = Nothing
Case "SIE"
Me.Caption = "Statement of Income and Expense"

objCommand.CommandText = query
Set rs6 = objCommand.Execute

objCommand2.CommandText = accnt
Set SRec = objCommand2.Execute
'
'Set rs6 = opndbaseFMIS.Execute(query)
'Set SRec = opndbaseFMIS.Execute(accnt)
    Screen.MousePointer = vbHourglass
    If Comparative = True Then
        CRViewer91.ReportSource = Rpt_SIE_Comparative
            With Rpt_SIE_Comparative
                .DiscardSavedData
                .Subreport1.OpenSubreport.DiscardSavedData
                .FormulaFields(16).Text = IEE
                .FormulaFields(19).Text = IEE100
                .txtDate.SetText (dated)
                .txtFund.SetText FundType
                .txtCurrentYear.SetText Format(ComparativeDate, "yyyy")
                .txtPreviousYear.SetText (Format(ComparativeDate, "yyyy") - 1)
                
                .txtComparativeYear.SetText "(With Comparative Figures from CY " & (Format(ComparativeDate, "yyyy") - 1) & ")"
                .Database.Tables(1).SetDataSource rs6
                .Subreport1.OpenSubreport.Database.Tables(1).SetDataSource SRec
               ' .txtfund.SetText FundType
                Set Rsource = Rpt_SIE_Comparative
                Call TransactionLogging("Print Preview", "Statement of Income and Expense(Comparative report)", Me.Caption, Winsock1.LocalIP)
            End With
        SRec.Close
        Set SRec = Nothing
    Else
        If AllFunds = True Then 'all Funds
            CRViewer91.ReportSource = Rpt_SIE_AllFunds
                With Rpt_SIE_AllFunds
                    .DiscardSavedData
                    .Subreport1.OpenSubreport.DiscardSavedData
                    .FormulaFields(16).Text = IEE
                    .FormulaFields(20).Text = IEE100
                    .FormulaFields(21).Text = IEE200
                    .txtDate.SetText (dated)
                    .Database.Tables(1).SetDataSource rs6
                    .Subreport1.OpenSubreport.Database.Tables(1).SetDataSource SRec
                   ' .txtfund.SetText FundType
                    Set Rsource = Rpt_SIE_AllFunds
                    Call TransactionLogging("Print Preview", "Statement of Income and Expense(All Funds Consolidated)", Me.Caption, Winsock1.LocalIP)
                End With
            SRec.Close
            Set SRec = Nothing
        Else
                CRViewer91.ReportSource = Rpt_SIE
                With Rpt_SIE
                    .DiscardSavedData
                    .Subreport1.OpenSubreport.DiscardSavedData
                    .FormulaFields(16).Text = IEE
                    .txtDate.SetText (dated)
                    .Database.Tables(1).SetDataSource rs6
                    .Subreport1.OpenSubreport.Database.Tables(1).SetDataSource SRec
                    .txtFund.SetText FundType
                    Set Rsource = Rpt_SIE
                    Call TransactionLogging("Print Preview", "Statement of Income and Expense(By funds)", Me.Caption, Winsock1.LocalIP)
                End With
            SRec.Close
            Set SRec = Nothing
    
        End If
    End If
Case "SGE"
Me.Caption = "Statement of Government Equity"
'MsgBox query

objCommand.CommandText = query
Set rs6 = objCommand.Execute



'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
If rs6.RecordCount > 0 Then
End If
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = CrystalReportSGE
        With CrystalReportSGE
             .DiscardSavedData
            .txtDate.SetText (dated)
            .Database.SetDataSource rs6
            .txtFund.SetText FundType
            Set Rsource = CrystalReportSGE
            '.txtClerk.SetText getUserName(ActiveUserID, "FullName")
            'Call TransactionLogging("Print Preview", "Balance Sheet", Me.Caption)
            Call TransactionLogging("Print Preview", "Statement of Government Equity", Me.Caption, Winsock1.LocalIP)
        End With
Case "Schedules"
Me.Caption = "Schedules"
'MsgBox query
objCommand.CommandText = query
Set rs6 = objCommand.Execute


'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = rpt_Schedule
        With rpt_Schedule
             .DiscardSavedData
             .txtFund.SetText FundType
            .txtDate.SetText (dated)
            .Database.SetDataSource rs6
             Set Rsource = rpt_Schedule
            '.txtClerk.SetText getUserName(ActiveUserID, "FullName")
            Call TransactionLogging("Print Preview", "Schedules", Me.Caption, Winsock1.LocalIP)
            End With
Case "SCF"
Me.Caption = "Statement of Cash Flow"
'MsgBox accnt

objCommand.CommandText = query
Set rs6 = objCommand.Execute

objCommand2.CommandText = accnt
Set rec = objCommand2.Execute

'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'rec.Open accnt, opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'MsgBox query
'MsgBox accnt
    If AllFunds = True Then
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = CrystalReportSCF_ALLFunds
            With CrystalReportSCF_ALLFunds
                 .DiscardSavedData
                 .txtFund.SetText FundType
                 .txtDate.SetText (dated)
                 .Database.SetDataSource rec
                 .Database.Tables(2).SetDataSource rs6
                 Set Rsource = CrystalReportSCF_ALLFunds
            End With
            rec.Close
            Set rec = Nothing
            Call TransactionLogging("Print Preview", "Statement of Cash Flow", Me.Caption, Winsock1.LocalIP)
    Else
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = CrystalReportSCF
            With CrystalReportSCF
                 .DiscardSavedData
                 .txtFund.SetText FundType
                 .txtDate.SetText (dated)
                 .Database.SetDataSource rec
                 .Database.Tables(2).SetDataSource rs6
                 Set Rsource = CrystalReportSCF
            End With
            rec.Close
            Set rec = Nothing
            Call TransactionLogging("Print Preview", "Statement of Cash Flow", Me.Caption, Winsock1.LocalIP)
            rs6.Close
            Set rs6 = Nothing
    End If
Case "Monitoring1", "Monitoring1"
Me.Caption = "Monitoring of Cash Advance"
'MsgBox query

objCommand.CommandText = query
Set rs6 = objCommand.Execute

'rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'MsgBox query
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = frm_QueryMonitoringofOOE
        With frm_QueryMonitoringofOOE
             .DiscardSavedData
             .txtFund.SetText FundType
             .txtDate.SetText (dated)
             .Database.SetDataSource rs6
             Set Rsource = frm_QueryMonitoringofOOE
        End With
        Call TransactionLogging("Print Preview", "Monitoring", Me.Caption, Winsock1.LocalIP)
Case "RRR"
Me.Caption = "REPORT OF REVENUE AND RECEIPTS"
'MsgBox query

objCommand.CommandText = accnt
Set rs6 = objCommand.Execute
'rs6.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = CrystalReportRRR
        With CrystalReportRRR
             .DiscardSavedData
             .txtFund.SetText FundType
             .txtdated.SetText (dated)
             '.txtname.SetText accountcode
             .Database.SetDataSource rs6
             Set Rsource = CrystalReportRRR
        End With
        Call TransactionLogging("Print Preview", "Report of Revenue and Receipts", Me.Caption, Winsock1.LocalIP)
End Select
    CRViewer91.PrintReport
    CRViewer91.ViewReport
    CRViewer91.Zoom 100 'Select  [fmis].[dbo].[MPfunc_GetBeginBalForCFlow] ('1/31/2011',1,'" & FundType & "')as BeginningBal
    Screen.MousePointer = vbDefault

'Set rs6 = Nothing
'rec.Close
Exit Sub
bad:
    If err.Number = -2147467259 Then
    Unload Me
    Else
    MsgBox "Error: " & err.description & err.Number
    End If
End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub


Private Sub page_Click()
Set frm_PageSetup.RPTreport = Rsource
centerme frm_PageSetup
frm_PageSetup.Show 1
CRViewer91.Refresh
End Sub

Private Sub print_Click()
Rsource.PrinterSetup (hwnd)
'Rsource.PaperSize = crDefaultPaperSize
'UseDefault = True
CRViewer91.Refresh
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
