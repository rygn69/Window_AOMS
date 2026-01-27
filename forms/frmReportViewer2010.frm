VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.8#0"; "FLATBTN2.OCX"
Begin VB.Form frmReportViewer 
   Caption         =   "Preview"
   ClientHeight    =   4515
   ClientLeft      =   7035
   ClientTop       =   2820
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   7110
   WindowState     =   2  'Maximized
   Begin DevPowerFlatBttn.FlatBttn ftbClose 
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   609
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   3690
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6465
      lastProp        =   600
      _cx             =   11404
      _cy             =   6509
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
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
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Report As New CRAXDRT.Report

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    CRViewer.ReportSource = Report
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
    Set frmReportViewer = Nothing
End Sub

Private Sub ftbClose_Click()
    Unload Me
End Sub
