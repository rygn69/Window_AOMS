VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frm_CashReceiptsConsolidated 
   Caption         =   "Report Viewer"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   Icon            =   "frm_CashReceiptsConsolidated.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   7275
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   500
      _cx             =   10231
      _cy             =   12347
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
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frm_CashReceiptsConsolidated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Rpt_CashReceiptsConsolidated
Public query, accnt As String, fund, mnth, typ, bankno As String

Private Sub Form_Load()
'On Error GoTo err
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
 
 rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Report
        With Report
             .DiscardSavedData
          .txtmonth.SetText Trim(mnth)
        .Text21.SetText Trim(fund)
'             .Text21.SetText Trim(TYP)
'             .txtbank.SetText bankno
            

            .Database.SetDataSource rs6
           .Subreport1.OpenSubreport.DiscardSavedData
           '.Subreport1_txtfunddate.
             .Subreport1.OpenSubreport.Database.SetDataSource rec
        End With
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Report
    CRViewer91.ViewReport
    Screen.MousePointer = vbDefault
    
Else
    MsgBox "No record found...", vbOKOnly, Me.Caption
    Unload Me
End If
    frmcashCashReceipts_Option.Animation1.Stop
    frmcashCashReceipts_Option.Animation1.Close
    frmcashCashReceipts_Option.Animation1.Visible = False
    rs6.Close
Set rs6 = Nothing
Exit Sub
err:
    
End Sub
Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
