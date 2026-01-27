VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmcheckdisbursement_gf 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   6885
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      lastProp        =   500
      _cx             =   11933
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
Attribute VB_Name = "frmcheckdisbursement_gf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptCD_GF
Public query, accnt As String, fund, mnth, TYP As String

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
             .txtfund.SetText Trim(fund)
             .Text21.SetText Trim(TYP)
            
            .Database.SetDataSource rs6
           .Subreport1.OpenSubreport.DiscardSavedData
             .Subreport1.OpenSubreport.Database.SetDataSource rec
        End With
    CRViewer91.PrintReport
    CRViewer91.ViewReport
    CRViewer91.Zoom 90
    Screen.MousePointer = vbDefault
Else
    MsgBox "No record found...", vbOKOnly, Me.Caption
    Unload Me
End If
rs6.Close
Set rs6 = Nothing
Exit Sub
err:
    MsgBox "Error: " & err.Description
End Sub


Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
