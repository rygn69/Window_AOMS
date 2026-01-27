VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frm_RPTviewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   7335
   ClientLeft      =   3615
   ClientTop       =   1410
   ClientWidth     =   6240
   Icon            =   "frm_RPTviewer.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cbl 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6165
      lastProp        =   600
      _cx             =   10874
      _cy             =   12779
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
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
      Begin VB.Menu print 
         Caption         =   "Printer Setup"
      End
      Begin VB.Menu PS 
         Caption         =   "Page Setup"
      End
   End
End
Attribute VB_Name = "frm_RPTviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Jquery, Rquery As String, fund, mnth, TYP, bankno, RptName As String
Public frm As Form
Public WRecap As Boolean 'With Recap
Public TrnsType As Integer
Public NoGroup As Boolean

Public RPTcashReceitps As New Rpt_CashReceiptsConsolidated
Public RPTcashReceitpsNoGroup As New Rpt_CashReceiptsConsolidated_NOgroup
Public RPTCheckNoGroup As New Rpt_checkConsolidated_NoGroup
Public RPTCashNogroup As New Rpt_cashConsolidated_NoGroup
Public RPTCheck As New Rpt_checkConsolidated
Public RPTCash As New Rpt_cashConsolidated
Public RPTGeneralJournal As New Rpt_GeneralJOurnalReport
Public Rsource As Object
Public Ifconso As Boolean

Private Sub Form_Load()
'On Error Resume Next
Dim Jrec As New ADODB.Recordset
Dim Rrec As New ADODB.Recordset

Dim objCommand As ADODB.command
Set objCommand = New ADODB.command
objCommand.CommandTimeout = 9999999
objCommand.ActiveConnection = opndbaseFMIS

Dim objCommand2 As ADODB.command
Set objCommand2 = New ADODB.command
objCommand2.CommandTimeout = 9999999
objCommand2.ActiveConnection = opndbaseFMIS

'MsgBox Jquery
opndbaseFMIS.Execute "Execute MPproc_NeedToExecute @type = 1"


objCommand.CommandText = Jquery
Set Jrec = objCommand.Execute
If WRecap = True Then

objCommand2.CommandText = Rquery
    Set Rrec = objCommand2.Execute
End If

'If Not JREc.EOF Then
    Screen.MousePointer = vbHourglass
    
    Select Case TrnsType
    Case 1 ' Cash Receipts
        If NoGroup = True Then
            Set Rsource = RPTcashReceitpsNoGroup
        Else
            Set Rsource = RPTcashReceitps
        End If
        
        CRViewer91.ReportSource = Rsource
            With Rsource
                .DiscardSavedData
                .txtmonth.SetText Trim(mnth)
                .Text21.SetText Trim(fund)
                .Database.SetDataSource Jrec
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText (GetEmpPosition(ActiveUserID))
                If NoGroup = True Then
                    If SystemMaintainance(1) = True Then
                        .txtSystemGEN.Suppress = True
                    End If
                End If
                
                 If Ifconso = True Then
                .Text1.Suppress = False
                .SumofCollection2.Suppress = False
                .SumofDeposits2.Suppress = False
                .SumofCREDIT2.Suppress = False
                .SumofDEBIT2.Suppress = False
                Else
                .Text1.Suppress = True
                .SumofCollection2.Suppress = True
                .SumofDeposits2.Suppress = True
                .SumofCREDIT2.Suppress = True
                .SumofDEBIT2.Suppress = True
                End If
                If WRecap = True Then
                    .Subreport1.OpenSubreport.DiscardSavedData
                    .Subreport1_txtFUNDdate.SetText fund & "-" & mnth & ""
                    .Subreport1.OpenSubreport.Database.SetDataSource Rrec
                    .Database.SetDataSource Jrec
                    If NoGroup = True Then
                        .Subreport1_txtprepared.SetText GetEmpName(ActiveUserID)
                        .Subreport1_txtposition.SetText (GetEmpPosition(ActiveUserID))
                    End If
                Else
                    .Subreport1.Suppress = True
                End If

            End With
        Call TransactionLogging("Print Preview", "Cash Receipts Journal", Me.Caption, Winsock1.LocalIP)
    Case 2 'Check Disbursement
        If NoGroup = True Then
            Set Rsource = RPTCheckNoGroup
        Else
            Set Rsource = RPTCheck
        End If
            CRViewer91.ReportSource = Rsource
            With Rsource
                .DiscardSavedData
                .txtmonth.SetText Trim(mnth)
                .Text23.SetText Trim(fund)
                .Database.SetDataSource Jrec
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText (GetEmpPosition(ActiveUserID))
                If NoGroup = True Then
                    If SystemMaintainance(1) = True Then
                        .txtSystemGEN.Suppress = True
                    End If
                End If
                
                
                If Ifconso = True Then
                .Text1.Suppress = False
                .Sumof1113.Suppress = False
                .SumofSDebit3.Suppress = False
                .SumofSCredit3.Suppress = False
                Else
                .Text1.Suppress = True
                .Sumof1113.Suppress = True
                .SumofSDebit3.Suppress = True
                .SumofSCredit3.Suppress = True
                End If
                 If WRecap = True Then
                    .Subreport1.OpenSubreport.DiscardSavedData
                    .Subreport1_txtFUNDdate.SetText fund & "-" & mnth & ""
                    .Subreport1.OpenSubreport.Database.SetDataSource Rrec
                    If NoGroup = True Then
                        .Subreport1_txtprepared.SetText GetEmpName(ActiveUserID)
                        .Subreport1_txtposition.SetText (GetEmpPosition(ActiveUserID))
                    End If
                Else
                
                    .Subreport1.Suppress = True
                End If
            End With
        Call TransactionLogging("Print Preview", "Check Disbursement Journal", Me.Caption, Winsock1.LocalIP)
    Case 3 ' Cash Disbursement
        If NoGroup = True Then
            Set Rsource = RPTCashNogroup
        Else
            Set Rsource = RPTCash
        End If
        CRViewer91.ReportSource = Rsource
            With Rsource
                .DiscardSavedData
                .txtmonth.SetText Trim(mnth)
                .Text21.SetText Trim(fund)
                .Database.SetDataSource Jrec
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText (GetEmpPosition(ActiveUserID))
                If NoGroup = True Then
                    If SystemMaintainance(1) = True Then
                        .txtSystemGEN.Suppress = True
                    End If
                End If
                
                If Ifconso = True Then
                .Text1.Suppress = False
                .Sumof1062.Suppress = False
                .SumofDEBIT2.Suppress = False
                .SumofCREDIT2.Suppress = False
                Else
                .Text1.Suppress = True
                .Sumof1062.Suppress = True
                .SumofDEBIT2.Suppress = True
                .SumofCREDIT2.Suppress = True
                End If
                If WRecap = True Then
                    .Subreport1.OpenSubreport.DiscardSavedData
                    '.Subreport1_Text5.SetText fund & "-(" & mnth & ")"
                    .Subreport1.OpenSubreport.Database.SetDataSource Rrec
                    .Subreport1_txtFUNDdate.SetText fund & "-" & mnth & ""
                    If NoGroup = True Then
                        .Subreport1_txtprepared.SetText GetEmpName(ActiveUserID)
                        .Subreport1_txtposition.SetText (GetEmpPosition(ActiveUserID))
                    End If
                Else
                    .Subreport1.Suppress = True
                End If
            End With
          'MsgBox Jquery
        Call TransactionLogging("Print Preview", "Cash Dibursement Journal", Me.Caption, Winsock1.LocalIP)
    Case 4 'General Journal
    
    CRViewer91.ReportSource = RPTGeneralJournal
            With RPTGeneralJournal
                .DiscardSavedData
                .txtmonth.SetText Trim(mnth)
                .Text23.SetText Trim(fund)
                .Database.SetDataSource Jrec
                .txtprepared.SetText GetEmpName(ActiveUserID)
                .txtposition.SetText (GetEmpPosition(ActiveUserID))
                If NoGroup = True Then
                    If SystemMaintainance(1) = True Then
                        .txtSystemGEN.Suppress = True
                    End If
                End If
                 If Ifconso = True Then
                .Text1.Suppress = False
                .RTDEBIT1.Suppress = False
                .RTCREDIT1.Suppress = False
                Else
                .Text1.Suppress = True
                .RTDEBIT1.Suppress = True
                .RTCREDIT1.Suppress = True
                End If
                Set Rsource = RPTGeneralJournal
                If WRecap = True Then
                    .Subreport1.OpenSubreport.DiscardSavedData
                    .Subreport1_txtFUNDdate.SetText fund & "-" & mnth & ""
                    .Subreport1.OpenSubreport.Database.SetDataSource Rrec
                    If NoGroup = True Then
                        .Subreport1_txtprepared.SetText GetEmpName(ActiveUserID)
                        .Subreport1_txtposition.SetText (GetEmpPosition(ActiveUserID))
                    End If
                Else
                    .Subreport1.Suppress = True
                End If
            End With
    Call TransactionLogging("Print Preview", "General Journal", Me.Caption, Winsock1.LocalIP)
    End Select
    CRViewer91.ViewReport
    Screen.MousePointer = vbDefault
    Jrec.Close
Set Jrec = Nothing
Exit Sub
err:
End Sub
Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub

Private Sub print_Click()
Rsource.PrinterSetup (hwnd)
UseDefault = True
CRViewer91.RefreshEx (True)
End Sub

Private Sub ps_Click()
Set frm_PageSetup.RPTreport = Rsource
medll.centerme frm_PageSetup
frm_PageSetup.Show 1
CRViewer91.Refresh
End Sub
