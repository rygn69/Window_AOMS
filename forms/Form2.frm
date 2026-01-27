VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   6075
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport3
Public query, accnt As String, fund, mnth, TYP, bankno As String
Private Sub Form_Load()
Dim rs6 As New ADODB.Recordset
Dim rec As New ADODB.Recordset
 
 rs6.Open "select * from vwAMIS_CDJournal where (year(checkdate)=2010 and month(checkdate)=9 and fndtype = 'General fund Proper') order by checkno,[111-1-11-WW] desc", opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
'rec.Open (accnt), opndbaseFMIS, adOpenDynamic, adLockBatchOptimistic
If Not rs6.EOF Then
    Screen.MousePointer = vbHourglass
    CRViewer91.ReportSource = Report
        With Report
             .DiscardSavedData
'             .txtmonth.SetText Trim(mnth)
'             .txtfund.SetText Trim(fund)
'             .Text21.SetText Trim(TYP)
'             .txtbank.SetText bankno
            
If TYP = "General Fund Proper" Then
.Text10.SetText "106-Others"
Else
.Text10.SetText "106-Canda"
End If
            .Database.SetDataSource rs6
'           .Subreport1.OpenSubreport.DiscardSavedData
'             .Subreport1.OpenSubreport.Database.SetDataSource rec
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
