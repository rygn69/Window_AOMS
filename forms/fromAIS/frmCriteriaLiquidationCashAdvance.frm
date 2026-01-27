VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmCriteriaLiquidationCashAdvance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   4485
   ClientTop       =   2520
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCriteriaLiquidationCashAdvance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5175
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   360
      Left            =   4110
      TabIndex        =   14
      Top             =   3405
      Width           =   960
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&Preview"
      Height          =   360
      Left            =   3135
      TabIndex        =   13
      Top             =   3405
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fund Type"
      Height          =   2310
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   975
      Width           =   2490
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         Height          =   240
         Left            =   75
         TabIndex        =   18
         Top             =   450
         Width           =   1530
      End
      Begin VB.ComboBox cboFundType 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   705
         Width           =   2370
      End
      Begin VB.ComboBox cboEco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   405
         TabIndex        =   11
         Top             =   1155
         Width           =   2010
      End
      Begin VB.CheckBox chkEco 
         Enabled         =   0   'False
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   1200
         Width           =   210
      End
      Begin VB.CheckBox chkProperAnd20 
         Caption         =   "GF Proper and 20% Dev't."
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   1770
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   2310
      Index           =   2
      Left            =   2565
      TabIndex        =   1
      Top             =   975
      Width           =   2550
      Begin VB.OptionButton optDateRange 
         Caption         =   "Date Range"
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   135
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   2
         Top             =   1065
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   405
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   58130435
         CurrentDate     =   38838
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   5
         Top             =   765
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   58130435
         CurrentDate     =   38868
      End
      Begin VB.Label Label6 
         Caption         =   "From"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   35
      Left            =   -90
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1035
      Top             =   8565
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   525
      Left            =   60
      TabIndex        =   15
      Top             =   3330
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   926
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   39
      FullHeight      =   35
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRITERIA FORM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   180
      TabIndex        =   17
      Top             =   210
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sets criteria for Liquidation and Cash Advances."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   480
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -30
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmCriteriaLiquidationCashAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
'*  Name         : cboEco_KeyPress
'*  Description  :
'*  Parameters   : KeyAscii As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub cboEco_KeyPress(KeyAscii As Integer)

    On Error GoTo errHandler
    KeyAscii = MyDLL.AutoFind(cboEco, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".cboEco_KeyPress"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : cboFundType_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub cboFundType_Click()

    On Error GoTo errHandler
    Dim opntbl As New ADODB.Recordset
    

    If Len(Trim$(cboFundType.Text)) <> 0 Then
        chkEco.Enabled = True

        opntbl.Open "select distinct ResponsibilityCenter from tblAIS_JEV where TypeOfFund='" & Trim$(cboFundType.Text) & "' and actioncode=1 and isapproved=1 order by ResponsibilityCenter", fmisDB, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            cboEco.Clear
            Do Until opntbl.EOF
                If IsNull(opntbl!ResponsibilityCenter) = False Then

                    cboEco.AddItem Trim$(opntbl!ResponsibilityCenter)
                Else
                    MsgBox "INCORRECT ENTRY in the Responsibility Center.A NULL value is found!" & vbCrLf & "Please contact the System Administrator.", vbCritical, "System Information"
                End If
                opntbl.MoveNext
            Loop
        End If
        
        opntbl.Close
        Set opntbl = Nothing
    Else
        cboEco.Clear
        chkEco.Enabled = False
        cboEco.Enabled = False
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".cboFundType_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : cboFundType_KeyPress
'*  Description  :
'*  Parameters   : KeyAscii As Integer
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub cboFundType_KeyPress(KeyAscii As Integer)

    On Error GoTo errHandler
    KeyAscii = MyDLL.AutoFind(cboFundType, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".cboFundType_KeyPress"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Sub chkConsolidated_Click()

    On Error GoTo errHandler
    If chkConsolidated.Value = 1 Then
        Call MainFund(cboFundType)
        cboEco.Enabled = False
        chkEco.Enabled = False
        chkProperAnd20.Value = 0
    Else
        Call FundType(cboFundType)
        cboEco.Enabled = False
        chkEco.Enabled = False
        chkProperAnd20.Value = 0
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".chkConsolidated_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : chkConsolidated_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

'*  Name         : chkEco_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub chkEco_Click()

    On Error GoTo errHandler
    If chkEco.Value = 1 Then
        cboEco.Enabled = True
    Else
        cboEco.Enabled = False
        cboEco.ListIndex = -1
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".chkEco_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : chkProperAnd20_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub chkProperAnd20_Click()

    On Error GoTo errHandler
    If chkProperAnd20.Value = 1 Then
        cboFundType.Enabled = False
        cboEco.Enabled = False
        chkEco.Enabled = False
        chkConsolidated.Value = 0
    Else
        cboFundType.Enabled = True
        cboEco.Enabled = False
        chkEco.Enabled = True
        chkConsolidated.Value = 0
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".chkProperAnd20_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : Command1_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Command1_Click()

    On Error GoTo errHandler
    Unload Me
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Command1_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : DisplayALLJEV
'*  Description  :
'*  Parameters   : cboName As ComboBox
'*  Returns      : Nothing
'*  Called From  : optJEVRange_Click, optJEVRange_Click
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************


Private Sub DTPicker1_Change()
    DTPicker1.Value = Month(DTPicker1.Value) & "/" & "1" & "/" & Year(DTPicker1.Value)
End Sub

Private Sub DTPicker2_Change()
    DTPicker2.Value = Month(DTPicker2.Value) & "/" & GetEndDateoftheMonth(DTPicker2.Value) & "/" & Year(DTPicker2.Value)
End Sub

'***************************************************************************
'*  Name         : FlatBttn1_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn1_Click()
    'FIXIT: Declare 'strMant' with an early-bound data type                                    FixIT90210ae-R1672-R1B8ZE
    Dim strMant As String, strYer As String
    Dim strSQLStatement As String
    Dim strSQLStatement2 As String

    On Error GoTo errHandler
    strSQLStatement = ""
    strReportName = "LIQCA"
    Call SetAnimation(frmCriteriaLiquidationCashAdvance.Animation1)
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQLStatement = "SELECT * FROM qryAIS_CashAdvanceLiquidation WHERE upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                                & " AND MantYer IN (" & MantYerRange & ")"
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        strSQLStatement = "SELECT * FROM qryAIS_CashAdvanceLiquidation WHERE upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                                & " AND MantYer IN (" & MantYerRange & ") And ResponsibilityCenter='" & Trim$(cboEco.Text) & "'"
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If Mid(Trim(cboFundType.Text), 1, 7) = "General" Then
                strSQLStatement = "SELECT * FROM qryAIS_CashAdvanceLiquidation WHERE TypeOfFund in (select FundName from tblRefBMS_Funds where ((FundName like 'gen%') or (FundName like '20%') or (FundName like 'econ%')) " _
                                & " AND MantYer IN (" & MantYerRange & "))"
            Else
                strSQLStatement = "SELECT * FROM qryAIS_CashAdvanceLiquidation WHERE upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                                & " AND MantYer IN (" & MantYerRange & "))"
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then    'if Proper and 20
            strSQLStatement = "SELECT * FROM qryAIS_CashAdvanceLiquidation WHERE TypeOfFund in (select FundName from tblRefBMS_Funds where ((FundName like 'gen%') or (FundName like '20%')) " _
                                & " AND MantYer IN (" & MantYerRange & "))"
        End If
    If Len(Trim(strSQLStatement)) > 0 Then
        strSQLStatement2 = "select * from qryAIS_CashAdvancesLiquidationsOtherRef"
        CrystalReportCA.Database.SetDataSource fmisDB.Execute(strSQLStatement)
        CrystalReportCA.Subreport1.OpenSubreport.Database.SetDataSource fmisDB.Execute(strSQLStatement2)
        Call TransactionLogging("Print Preview", "Cash Advances and Liquidations", "frmCriteriaLiquidationCashAdvance")
        Call UnsetAnimation(frmCriteriaLiquidationCashAdvance.Animation1)
        PreviewForm.Show vbModal
    Else
        MsgBox "Please verify the criteria!", vbCritical, "System Information"
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn1_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub


'*********************Err******************************************************
'*  Name         : FlatBttn1_MouseMove
'*  Description  :
'*  Parameters   : Button As Integer, Shift As Integer,
'*               : x As Single, Y As Single
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub FlatBttn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".FlatBttn1_MouseMove"
        Set .Error = Err
     
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
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Load()

    On Error GoTo errHandler
    'WindowsXPC1.InitSubClassing
    'MyDLL.CenterMe Me
    'Call DisplayOfficeUnderEcoEnt
    DTPicker2.Value = Month(ServerDate()) & "/" & GetEndDateoftheMonth(ServerDate()) & "/" & Year(ServerDate())
    Call FundType(cboFundType)
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Load"
        Set .Error = Err
     
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

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo errHandler
    WindowsXPC1.EndWinXPCSubClassing
    Set frmSOIAE = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

'***************************************************************************
'*  Name         : MantYerRange
'*  Description  :
'*  Parameters   : None
'*  Returns      : String
'*  Called From  : FlatBttn1_Click, FlatBttn1_Click, FlatBttn1_Click, FlatBttn1_Click
'*  Author       : Errol Bagaipo
'*  Date         : 06 Dec 2006
'*  Note         :
'***************************************************************************

Private Function MantYerRange() As String
Dim strMantYerRange As String
Dim xxx As Integer
     
    strMantYerRange = ""
    For xxx = 0 To DateDiff("m", DTPicker1.Value, DTPicker2.Value)
        If xxx = 0 Then
            strMantYerRange = "'" & "" & Format(DateAdd("m", xxx, DTPicker1.Value), "M-YY") & "" & "'"
        Else
            strMantYerRange = strMantYerRange & ",'" & "" & Format(DateAdd("m", xxx, DTPicker1.Value), "M-YY") & "" & "'"
        End If
    Next
    MantYerRange = strMantYerRange
End Function

'***************************************************************************
'*  Name         : optDateRange_Click
'*  Description  :
'*  Parameters   : None
'*  Returns      : Nothing
'*  Called From  :
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub optDateRange_Click()

    On Error GoTo errHandler
    If optDateRange.Value = True Then
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".optDateRange_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub






