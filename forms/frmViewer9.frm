VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewer9 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   8805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9645
      lastProp        =   500
      _cx             =   17013
      _cy             =   15531
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
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Menu ps 
      Caption         =   "Printer SetUp"
   End
End
Attribute VB_Name = "frmViewer9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sql As String

Private Sub Form_Load()

Me.Caption = App.CompanyName & " Report Dialogue"
Call LoadPageSetup

End Sub
Private Sub LoadPageSetup()
Dim rec As New ADODB.Recordset
Dim newval(0 To 3) As Variant
Dim usedef As Variant

'reading default --------
usedef = readTXTDATA("usepagesetup", "use", ReportLocation & "\pagesetup.ini")

If usedef = "Yes" Then
    newval(0) = readTXTDATA("Page Setup", "Top", ReportLocation & "\pagesetup.ini")
    newval(1) = readTXTDATA("Page Setup", "Left", ReportLocation & "\pagesetup.ini")
    newval(2) = readTXTDATA("Page Setup", "Bottom", ReportLocation & "\pagesetup.ini")
    newval(3) = readTXTDATA("Page Setup", "Right", ReportLocation & "\pagesetup.ini")
    
    Select Case Report9
        Case "Subsidiary"
            Rpt_SubsidiaryLedger.TopMargin = newval(0) * 1440
            Rpt_SubsidiaryLedger.LeftMargin = newval(1) * 1440
            Rpt_SubsidiaryLedger.BottomMargin = newval(2) * 1440
            Rpt_SubsidiaryLedger.RightMargin = newval(3) * 1440
            CRViewer91.ReportSource = frmViewer9
    End Select
   CRViewer91.ViewReport

Else
    Select Case Report9
        Case "Subsidiary"
'            Screen.MousePointer = vbHourglass
'             'CrystalReport1.Database.SetDataSource rec
'            CRViewer91.ReportSource = CrystalReport1
'            rec.Open (SQL), opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
'            CrystalReport1.DiscardSavedData
'           CrystalReport1.Database.SetDataSource rec
    End Select
     CRViewer91.PrintReport
    CRViewer91.ViewReport
    Screen.MousePointer = vbDefault
End If

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Select Case Report9
        Case "Subsidiary"
            Set Rpt_SubsidiaryLedger = Nothing
        Case "JEV"
            Set rptJEV = Nothing
        Case "Accomplishment"
            Set rptDailyAccomplishments = Nothing
        End Select
        
Set frmViewer9 = Nothing
End Sub

Private Sub page_Click()
frmPageSetup.Show vbModal
End Sub

Private Sub printer_Click()
Select Case Report9
        Case "Subsidiary"
            Rpt_SubsidiaryLedger.PrinterSetup Me.hwnd
        Case "JEV"
            rptJEV.PrinterSetup Me.hwnd
        Case "Accomplishment"
            rptDailyAccomplishments.PrinterSetup Me.hwnd
End Select
CRViewer91.Refresh
End Sub


