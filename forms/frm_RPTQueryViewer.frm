VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frm_RPTQueryViewer 
   Caption         =   "Query Report Viewer"
   ClientHeight    =   7635
   ClientLeft      =   5760
   ClientTop       =   3180
   ClientWidth     =   6930
   Icon            =   "frm_RPTQueryViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   6930
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      lastProp        =   500
      _cx             =   10398
      _cy             =   11245
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
      DisplayBorder   =   -1  'True
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
      Begin VB.Menu page 
         Caption         =   "Page Setup"
      End
   End
End
Attribute VB_Name = "frm_RPTQueryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public query, accnt, dated, mnth As String
Public Date_ As Date
Public Whatcon As Integer
Public RPTSource As Object
Public RPTSourceSE As New RPT_SE

Private Sub Form_Load()
'On Error GoTo bad
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim Mrec As New ADODB.Recordset
Dim Trec As New ADODB.Recordset
opndbaseFMIS.Execute "Execute MPproc_NeedToExecute @type = 1"
Select Case (Report9):

Case "statofcashadvance"
    Me.Caption = "Statofcashadvance"
    Set rec = opndbaseFMIS.Execute(accnt)
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = rpt_StatOfCashAdvance
        With rpt_StatOfCashAdvance
            .DiscardSavedData
            .Database.SetDataSource rec
            .txtdate.SetText dated
            Call TransactionLogging("Print Preview", "Status of Cash Advance", Me.Caption, Winsock1.LocalIP)
        End With
    Set RPTSource = rpt_StatOfCashAdvance
Case "statofliquidation"
    Me.Caption = "statofliquidation"
    Set rec = opndbaseFMIS.Execute(accnt)
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = RPT_StatOfLiquidation
        With RPT_StatOfLiquidation
            .DiscardSavedData
            .Database.SetDataSource rec
            .txtdate.SetText dated
            Call TransactionLogging("Print Preview", "Status of liquidation", Me.Caption, Winsock1.LocalIP)
            Set RPTSource = RPT_StatOfLiquidation
        End With
Case "SE"
    Me.Caption = "Statement of Expenses"
    Set rec = opndbaseFMIS.Execute(accnt)
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = RPTSourceSE
        With RPTSourceSE
            .DiscardSavedData
            .Database.Tables(1).SetDataSource rec
            .Database.Tables(2).SetDataSource opndbaseFMIS.Execute("select * from tblAMIS_HeaderDynamicColumn")
            .txtdate.SetText dated
            Call TransactionLogging("Print Preview", "Statement of Expenses", Me.Caption, Winsock1.LocalIP)
        End With
        Set RPTSource = RPTSourceSE
Case "PTOAccnt111"
    Me.Caption = "Accounting Reconcilliation to PTO"
    Set rec = opndbaseFMIS.Execute(accnt)
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = RPT_111Recon
        With RPT_111Recon
            .DiscardSavedData
            .Database.SetDataSource rec
            Call TransactionLogging("Print Preview", "Accounting Reconcilliation to PTO", Me.Caption, Winsock1.LocalIP)
        End With
        Set RPTSource = RPT_111Recon
        
End Select
    CRViewer91.PrintReport
    CRViewer91.ViewReport
    CRViewer91.Zoom 90
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
Set frm_PageSetup.RPTreport = RPTSource
medll.centerme frm_PageSetup
frm_PageSetup.Show 1
CRViewer91.Refresh
End Sub

Private Sub print_Click()
RPTSource.PrinterSetup (hwnd)
RPTSource.PaperSize = crDefaultPaperSize
UseDefault = True
CRViewer91.RefreshEx (True)
End Sub
