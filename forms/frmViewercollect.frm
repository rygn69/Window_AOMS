VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewercollect 
   Caption         =   "Print"
   ClientHeight    =   5055
   ClientLeft      =   4920
   ClientTop       =   7545
   ClientWidth     =   7065
   Icon            =   "frmViewercollect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7065
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      lastProp        =   500
      _cx             =   12091
      _cy             =   8705
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Menu pmenu 
      Caption         =   "Print menu"
      Begin VB.Menu printer 
         Caption         =   "P&rinter Setup"
      End
      Begin VB.Menu page 
         Caption         =   "&Page Setup"
      End
   End
End
Attribute VB_Name = "frmViewercollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Caption = App.CompanyName & " Report Dialogue"
Call LoadPageSetup
End Sub
Private Sub LoadPageSetup()
Dim newval(0 To 3) As Variant
Dim usedef As Variant
'reading default --------
usedef = readTXTDATA("usepagesetup", "use", ReportLocation & "\pagesetup.ini")
If usedef = "Yes" Then
    newval(0) = readTXTDATA("Page Setup", "Top", ReportLocation & "\pagesetup.ini")
    newval(1) = readTXTDATA("Page Setup", "Left", ReportLocation & "\pagesetup.ini")
    newval(2) = readTXTDATA("Page Setup", "Bottom", ReportLocation & "\pagesetup.ini")
    newval(3) = readTXTDATA("Page Setup", "Right", ReportLocation & "\pagesetup.ini")
    Select Case ReportName
        Case "AcctAdvice"
            rptAccntAdvice.TopMargin = newval(0) * 1440
            rptAccntAdvice.LeftMargin = newval(1) * 1440
            rptAccntAdvice.BottomMargin = newval(2) * 1440
            rptAccntAdvice.RightMargin = newval(3) * 1440
            CRViewer1.ReportSource = rptAccntAdvice
        Case "JEV"
            rptJEVCollection.TopMargin = newval(0) * 1440
            rptJEVCollection.LeftMargin = newval(1) * 1440
            rptJEVCollection.BottomMargin = newval(2) * 1440
            rptJEVCollection.RightMargin = newval(3) * 1440
            CRViewer1.ReportSource = rptJEVCollection
        Case "Accomplishment"
            rptDailyAccomplishments.TopMargin = newval(0) * 1440
            rptDailyAccomplishments.LeftMargin = newval(1) * 1440
            rptDailyAccomplishments.BottomMargin = newval(2) * 1440
            rptDailyAccomplishments.RightMargin = newval(3) * 1440
            CRViewer1.ReportSource = rptDailyAccomplishments
    End Select
    CRViewer1.ViewReport
Else
    Select Case ReportName
        Case "AcctAdvice"
            CRViewer1.ReportSource = rptAccntAdvice
        Case "JEV"
            CRViewer1.ReportSource = rptJEVCollection
        Case "Accomplishment"
            CRViewer1.ReportSource = rptDailyAccomplishments
    End Select
    CRViewer1.ViewReport
End If
End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
Select Case ReportName
        Case "AcctAdvice"
            Set rptAccntAdvice = Nothing
        Case "JEV"
            Set rptJEVCollection = Nothing
        Case "Accomplishment"
            Set rptDailyAccomplishments = Nothing
        End Select
Set frmViewer = Nothing
End Sub
Private Sub page_Click()
frmPageSetup.Show vbModal
End Sub
Private Sub printer_Click()
Select Case ReportName
    Case "AcctAdvice"
        rptAccntAdvice.PrinterSetup Me.hwnd
    Case "JEV"
        rptJEVCollection.PrinterSetup Me.hwnd
    Case "Accomplishment"
        rptDailyAccomplishments.PrinterSetup Me.hwnd
End Select
CRViewer1.Refresh
End Sub

