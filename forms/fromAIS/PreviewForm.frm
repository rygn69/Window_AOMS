VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PreviewForm 
   BorderStyle     =   0  'None
   ClientHeight    =   7110
   ClientLeft      =   2685
   ClientTop       =   1905
   ClientWidth     =   5880
   Icon            =   "PreviewForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDB 
      Left            =   6000
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   45
      Top             =   7500
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      MsgBox_InputBox =   0   'False
      Common_Dialog   =   0   'False
      OptionControl   =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5805
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu mnuPrntOpt 
      Caption         =   "&Printer Options"
      Begin VB.Menu mnuPrntSetup 
         Caption         =   "&Printer Setup"
      End
      Begin VB.Menu mnuPageMargin 
         Caption         =   "Page &Margin"
      End
   End
End
Attribute VB_Name = "PreviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ClickPrintButton As Boolean


'***************************************************************************
'*  Name         : CRViewer1_PrintButtonClicked
'*  Description  :
'*  Parameters   : UseDefault As Boolean
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Mar Paul M. Ajero
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)

    On Error GoTo errHandler
    ClickPrintButton = True
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".CRViewer1_PrintButtonClicked"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Load
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Mar Paul M. Ajero
'*  Date         : August 22, 2011
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Load()

'    On Error GoTo errHandler

    'WindowsXPC1.InitSubClassing
    Screen.MousePointer = vbHourglass
    'ImplodeFormToTray Me.hwnd, True
    'Call SelectPrinter(Printer.DeviceName)
    Select Case strReportName
        Case "POREG"
           ' Call SetPrinterAndPapersize(crptPORegistry)
            'CRViewer1.ReportSource = crptPORegistry
        Case "CRJRECAP"
            'Call SetPrinterAndPapersize(CrystalReportRecapCRJ)
            'CRViewer1.ReportSource = CrystalReportRecapCRJ
        Case "CKDJRECAP"
           ' Call SetPrinterAndPapersize(CrystalReportRecapCKDJ)
            'CRViewer1.ReportSource = CrystalReportRecapCKDJ
        Case "CDJRECAP"
            'CrystalReportRecapCDJ.Database.SetDataSource opndbaseFMIS.Execute(sql)
            'Call SetPrinterAndPapersize(CrystalReportRecapCDJ)
            'CRViewer1.ReportSource = CrystalReportRecapCDJ
            'Case "PRETB", "POSTTB"
            'CrystalReportTBConsolidated.Database.SetDataSource opndbaseFMIS.Execute(sql)
            'CRViewer1.ReportSource = CrystalReportTBConsolidated
        Case "CBS"
            'CrystalReportBalanceSheetConsolidated.Database.SetDataSource opndbaseFMIS.Execute(sql)
            'CRViewer1.ReportSource = CrystalReportBalanceSheetConsolidated
        Case "LRC"
           ' Call SetPrinterAndPapersize(CrystalReportResponCenter)
            'CrystalReportResponCenter.Database.SetDataSource opndbaseFMIS.Execute(SQL)
            'CRViewer1.ReportSource = CrystalReportResponCenter
        Case "SMRFS"
            'Call SetPrinterAndPapersize(CrystalReportSMRFS)
            'CRViewer1.ReportSource = CrystalReportSMRFS
        Case "SAAO"
            'CRViewer1.ReportSource = crpSAAO
        Case "CRJ"
            'Call SetPrinterAndPapersize(crpCRJ)
            'CRViewer1.ReportSource = crpCRJ
        Case "CDJ"
            'Call SetPrinterAndPapersize(crpCDJ)
            'CRViewer1.ReportSource = crpCDJ
        Case "CKDJ" 'error here..........................
            'Call SetPrinterAndPapersize(crpCKDJ)
            'CRViewer1.ReportSource = crpCKDJ
        Case "JEV"
'            Call SetPrinterAndPapersize(crpJEV)
'            crpJEV.Database.SetDataSource opndbaseFMIS.Execute(SQL)
'            crpJEV.Subreport1.OpenSubreport.Database.SetDataSource opndbaseFMIS.Execute(sql2)
'            CRViewer1.ReportSource = crpJEV
        Case "SOA"
            crptPPAallotment.DiscardSavedData
            CRViewer1.ReportSource = crptPPAallotment
            crptPPAallotment.DiscardSavedData
        Case "BS"
            CrystalReportBS.DiscardSavedData
            CRViewer1.ReportSource = CrystalReportBS
        Case "SIE"
            Rpt_SIE.DiscardSavedData
            CRViewer1.ReportSource = Rpt_SIE
'        Case Else
    End Select
    
    CRViewer1.Zoom 1
    CRViewer1.ViewReport
    
    Screen.MousePointer = vbDefault

    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Load"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Resize
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Resize()

    On Error GoTo errHandler
    
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Resize"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Form_Unload
'*  Description  :
'*  Parameters   : Cancel As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub mnuPageMargin_Click()
    frmPageSetup.Show vbModal
End Sub

Private Sub mnuPrntSetup_Click()
'
'    Select Case UCase$(strReportName)
'        Case "POREG"
'            Call SetPrinterOptions(crptPORegistry)
'        Case "CKDJRECAP"
'            Call SetPrinterOptions(CrystalReportRecapCKDJ)
'        Case "CDJRECAP"
'            Call SetPrinterOptions(CrystalReportRecapCDJ)
'        Case "CRJRECAP"
'            Call SetPrinterOptions(CrystalReportRecapCRJ)
'        Case "TB"
'            Call SetPrinterOptions(CrystalReportTB)
'        Case "LRC"
'            Call SetPrinterOptions(CrystalReportResponCenter)
'        Case "SMRFS"
'            Call SetPrinterOptions(CrystalReportSMRFS)
'        Case "CRJ"
'            Call SetPrinterOptions(crpCRJ)
'        Case "CDJ"
'            Call SetPrinterOptions(crpCDJ)
'        Case "CKDJ"
'            Call SetPrinterOptions(crpCKDJ)
'        Case "JEV", "JEVPRINT"
'            Call SetPrinterOptions(crpJEV)
'            Call SetPrinterOptions(crpJEVDetails)
'        Case "ADV"
'            Call SetPrinterOptions(CrystalReportADVICE)
'        Case "SL"
'            Call SetPrinterOptions(CrystalReportSL2)
'        Case "GL"
'            Call SetPrinterOptions(crpGL2)
'        Case "GJRECAP"
'            Call SetPrinterOptions(CrystalReportRecapGJ)
'        Case "GJ"
            Call SetPrinterOptions(Rpt_TrialBalance)
'        Case "BS"
'            Call SetPrinterOptions(CrystalReportBS)
'        Case "SIAE"
'            Call SetPrinterOptions(CrystalReportSIE)
'        Case "SCHEDULE"
'            Call SetPrinterOptions(CrystalReportSked)
'        Case "SOCF"
'            Call SetPrinterOptions(CrystalReportSCF)
'        Case "SOGE"
'            Call SetPrinterOptions(CrystalReportSGE)
'        Case "OTHERSCHEDULE"
'            Call SetPrinterOptions(CrystalReportOtherSked)
'        Case "LORPDO"
'            Call SetPrinterOptions(crptListOfReportNoPerDisbursing)
'        Case "RRR"
'            Call SetPrinterOptions(CrystalReportRRR)
'        Case "LIQREG"
'            Call SetPrinterOptions(crptLiquidationReg)
'        Case "LIQCA"
'            Call SetPrinterOptions(CrystalReportCA)
'        Case Else
'
'    End Select
CDB.ShowPrinter
CDB.PrinterDefault = True
 rptJEVNew.PaperSize = crDefaultPaperSize
CRViewer1.Refresh
End Sub

Private Sub SetPrinterOptions(ByVal crptName As Report)
    crptName.PrinterSetup Me.hwnd
    CRViewer1.Refresh
End Sub

Public Sub SetPrinterAndPapersize(ByVal crptName As Report)
'    crptName.SelectPrinter printer.DeviceName, printer.DriverName, printer.Port
'    crptName.PaperSize = printer.PaperSize
'    crptName.TopMargin = dblTop
'    crptName.BottomMargin = dblBottom
'    crptName.LeftMargin = dblLeft
'    crptName.RightMargin = dblRight
End Sub

