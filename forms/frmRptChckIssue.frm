VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmRptChckIssue 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
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
      EnableHelpButton=   -1  'True
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmRptChckIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptReportchchissued
Public query, fnd, bnk, DTE, RCI, sig As String
Private Sub Form_Load()
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
 
 rs6.Open (query), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic

If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Report
        With Report
             .DiscardSavedData
                .txtbankname.SetText Trim(bnk)
                .txtfundtype.SetText Trim(fnd)
                .txtmonth.SetText Trim(DTE)
                .txtrci.SetText Trim(RCI)
            .Database.SetDataSource rs6
           
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
