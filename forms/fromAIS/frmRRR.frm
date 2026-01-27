VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmRRR 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   5820
   ClientTop       =   1380
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRRR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5130
   Begin VB.Frame Frame1 
      Caption         =   "Fund Type"
      Height          =   2310
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   1005
      Width           =   2490
      Begin VB.CheckBox chkProperAnd20 
         Caption         =   "GF Proper and 20% Dev't."
         Height          =   240
         Left            =   75
         TabIndex        =   15
         Top             =   1770
         Width           =   2385
      End
      Begin VB.CheckBox chkEco 
         Enabled         =   0   'False
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   1200
         Width           =   210
      End
      Begin VB.ComboBox cboEco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   405
         TabIndex        =   13
         Top             =   1155
         Width           =   2010
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   420
         Width           =   1530
      End
      Begin VB.ComboBox cboFundType 
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   675
         Width           =   2370
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   2310
      Index           =   2
      Left            =   2535
      TabIndex        =   3
      Top             =   1005
      Width           =   2550
      Begin VB.OptionButton optDateRange 
         Caption         =   "Month Range"
         Height          =   240
         Left            =   60
         TabIndex        =   5
         Top             =   135
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   4
         Top             =   1065
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   6
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
         CustomFormat    =   "MMM yyyy"
         Format          =   20709379
         CurrentDate     =   38695
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   7
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
         CustomFormat    =   "MMM yyyy"
         Format          =   20709379
         CurrentDate     =   38330
      End
      Begin VB.Label Label6 
         Caption         =   "From Month"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To Month"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&Preview"
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   3405
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   360
      Left            =   4095
      TabIndex        =   1
      Top             =   3405
      Width           =   960
   End
   Begin VB.Frame Frame3 
      Height          =   35
      Left            =   -150
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1020
      Top             =   8565
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   525
      Left            =   45
      TabIndex        =   16
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
      Left            =   165
      TabIndex        =   18
      Top             =   210
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sets criteria for Revenue and Receipts Report prior to previewing."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   165
      TabIndex        =   17
      Top             =   480
      Width           =   4860
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -90
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmRRR"
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
    KeyAscii = mydll.AutoFind(cboEco, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboEco_KeyPress"
        Set .Error = err
     
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
    
    Call SetAnimation(frmRRR.Animation1)
    If Len(Trim$(cboFundType.Text)) <> 0 Then
        chkEco.Enabled = True

        opntbl.Open "select distinct ResponsibilityCenter from tblAIS_JEV where TypeOfFund='" & Trim$(cboFundType.Text) & "' and actioncode=1 and isapproved=1 order by ResponsibilityCenter", opndbaseFMIS, adOpenStatic, adLockOptimistic
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
    Call UnsetAnimation(frmRRR.Animation1)
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboFundType_Click"
        Set .Error = err
     
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
    KeyAscii = mydll.AutoFind(cboFundType, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboFundType_KeyPress"
        Set .Error = err
     
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
        err.Source = err.Source & "." & TypeName(Me) & ".chkConsolidated_Click"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

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
        err.Source = err.Source & "." & TypeName(Me) & ".chkEco_Click"
        Set .Error = err
     
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
        err.Source = err.Source & "." & TypeName(Me) & ".chkProperAnd20_Click"
        Set .Error = err
     
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
        err.Source = err.Source & "." & TypeName(Me) & ".Command1_Click"
        Set .Error = err
     
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

Private Sub DisplayALLJEV(ByVal cboName As ComboBox)
    '***************************************************************************
    '*  Name         : DisplayALLJEV
    '*  Description  :
    '*  Parameters   : cboName As ComboBox
    '*  Returns      : Nothing
    '*  Author       : Errol Bagaipo
    '*  Date         : 25 Oct 2006
    '***************************************************************************


    On Error GoTo errHandler
    Dim opntbl As New ADODB.Recordset
    
    cboName.Clear
    If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then

        If Len(Trim$(cboEco.Text)) = 0 Then
            opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where upper(typeoffund)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", opndbaseFMIS, adOpenStatic, adLockOptimistic
        Else
            opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where responsibilitycenter='" & Trim$(cboEco.Text) & "' and  upper(typeoffund)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", opndbaseFMIS, adOpenStatic, adLockOptimistic
        End If
    ElseIf chkConsolidated = 1 And chkProperAnd20.Value = 0 Then
        opntbl.Open "SELECT DISTINCT dbo.tblAIS_JEV.JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) FROM dbo.tblAIS_JEV INNER JOIN dbo.tblRefBMS_Funds ON dbo.tblAIS_JEV.TypeOfFund = dbo.tblRefBMS_Funds.FundName INNER JOIN dbo.tblREF_AIS_Fundtype ON dbo.tblRefBMS_Funds.MotherFund = dbo.tblREF_AIS_Fundtype.fundcode Where (dbo.tblAIS_JEV.actioncode = 1) And (dbo.tblAIS_JEV.IsApproved = 1) and upper(dbo.tblREF_AIS_Fundtype.MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsClosed = 0) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc,jevnumber", opndbaseFMIS, adOpenStatic, adLockOptimistic
    ElseIf chkConsolidated = 0 And chkProperAnd20.Value = 1 Then
        opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where upper(typeofFund) in ('" & "GENERAL FUND PROPER" & "','" & "20% DEVELOPMENT FUND" & "') and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", opndbaseFMIS, adOpenStatic, adLockOptimistic
    End If
    If opntbl.RecordCount > 0 Then
        Do Until opntbl.EOF
            cboName.AddItem opntbl!JEVNUmber
            opntbl.MoveNext
        Loop
    Else

        If Len(Trim$(cboFundType.Text)) = 0 Then
            MsgBox "Please specify the type of fund!", vbCritical, "System Information"
        End If
    End If
    
    opntbl.Close
    Set opntbl = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".DisplayALLJEV"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
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
    Dim opntbl As New ADODB.Recordset

    On Error GoTo errHandler
    strReportName = "RRR"
    Call SetAnimation(frmRRR.Animation1)
    opndbaseFMIS.Execute "update tblAIS_JEV set IsRRR=0"
    If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
        If optDateRange.Value = True Then    'if date range
            If chkEco.Value = 0 Then
                opntbl.Open "select AccountCode from tblRef_AIS_RRRMain", opndbaseFMIS, adOpenStatic, adLockOptimistic
                Do Until opntbl.EOF
                    opndbaseFMIS.Execute "update tblAIS_JEV set IsRRR=1 where AccountCode='" & Trim(opntbl!Accountcode) & "' and actioncode=1 and isapproved=1 and TypeOfFund='" & Trim(cboFundType.Text) & "' and cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "month(jevdate) ELSE " _
                                    & "cast(substring(jevnumber, 8, 2) AS Int) END) AS nvarchar) + '-' " _
                                    & "+ cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "substring(cast(Year(jevdate) " _
                                    & "AS nvarchar), 3, 2) ELSE substring(jevnumber, 5, 2) END) " _
                                    & "AS nvarchar) in (" & MantYerRange & ")"
                opntbl.MoveNext
                Loop
                
                opntbl.Close
                Set opntbl = Nothing
                
                'CrystalReportSked.Database.SetDataSource opndbaseFMIS.Execute("select AccountNameExt, AccountNameExt2, sum(BalanceAmount) + ISNULL ((SELECT " _
                        & "amount FROM qryAIS_ListOfAccountsPayableAndLiquidation WHERE qryAIS_ReportSchedulesFinal.JEVNumber = " _
                        & "dbo.qryAIS_ListOfAccountsPayableAndLiquidation.jevnumber AND ChildAccountCode = " _
                        & "qryAIS_ListOfAccountsPayableAndLiquidation.AccountCode), 0) " _
                        & "AS BalanceAmount, ReportNo, Particulars, MantYer, TypeOfFund, " _
                        & " MainCode, AccountName, AccountNameFull,ChildAccountCode from qryAIS_ReportSchedulesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") GROUP BY MainCode,AccountNameExt,AccountNameExt2,ReportNo, Particulars, MantYer, TypeOfFund, AccountName, AccountNameFull,ChildAccountCode, JEVNumber order by MainCode,AccountNameExt")
            Else
                If Len(Trim$(cboEco.Text)) > 0 Then
                    opntbl.Open "select AccountCode from tblRef_AIS_RRRMain", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    Do Until opntbl.EOF
                        opndbaseFMIS.Execute "update tblAIS_JEV set IsRRR=1 where AccountCode='" & Trim(opntbl!Accountcode) & "' and actioncode=1 and isapproved=1 and TypeOfFund='" & Trim(cboFundType.Text) & "' and ResponsibilityCenter='" & Trim(cboEco.Text) & "' and cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "month(jevdate) ELSE " _
                                    & "cast(substring(jevnumber, 8, 2) AS Int) END) AS nvarchar) + '-' " _
                                    & "+ cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "substring(cast(Year(jevdate) " _
                                    & "AS nvarchar), 3, 2) ELSE substring(jevnumber, 5, 2) END) " _
                                    & "AS nvarchar) in (" & MantYerRange & ")"
                    opntbl.MoveNext
                    Loop
                    
                    opntbl.Close
                    Set opntbl = Nothing
                    'CrystalReportSked.Database.SetDataSource opndbaseFMIS.Execute("select AccountNameExt, AccountNameExt2, sum(BalanceAmount) + ISNULL ((SELECT " _
                        & "amount FROM qryAIS_ListOfAccountsPayableAndLiquidation WHERE qryAIS_ReportSchedulesFinal.JEVNumber = " _
                        & "dbo.qryAIS_ListOfAccountsPayableAndLiquidation.jevnumber AND ChildAccountCode = " _
                        & "qryAIS_ListOfAccountsPayableAndLiquidation.AccountCode), 0) " _
                        & "AS BalanceAmount, ReportNo, Particulars, MantYer, TypeOfFund,ResponsibilityCenter," _
                            & " MainCode, AccountName, AccountNameFull,ChildAccountCode from qryAIS_ReportSchedulesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") GROUP BY MainCode,AccountNameExt,AccountNameExt2,ReportNo, Particulars, MantYer, TypeOfFund, AccountName, AccountNameFull,ResponsibilityCenter,ChildAccountCode, JEVNumber order by MainCode,AccountNameExt")
                Else
                    MsgBox "Please select type of fund!", vbCritical, "System Information"
                    Exit Sub
                End If
            End If
        End If
    ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
        If optDateRange.Value = True Then    'if date range
            opntbl.Open "select AccountCode from tblRef_AIS_RRRMain", opndbaseFMIS, adOpenStatic, adLockOptimistic
            Do Until opntbl.EOF
                opndbaseFMIS.Execute "update tblAIS_JEV set IsRRR=1 where AccountCode='" & Trim(opntbl!Accountcode) & "' and actioncode=1 and isapproved=1 and TypeOfFund in (SELECT FundName From qryAIS_MotherFundDetails WHERE (MotherFundType = '" & Trim(cboFundType.Text) & "')) and cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "month(jevdate) ELSE " _
                                    & "cast(substring(jevnumber, 8, 2) AS Int) END) AS nvarchar) + '-' " _
                                    & "+ cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "substring(cast(Year(jevdate) " _
                                    & "AS nvarchar), 3, 2) ELSE substring(jevnumber, 5, 2) END) " _
                                    & "AS nvarchar) in (" & MantYerRange & ")"
            opntbl.MoveNext
            Loop
            
            opntbl.Close
            Set opntbl = Nothing
            'CrystalReportSked.Database.SetDataSource opndbaseFMIS.Execute("select AccountNameExt, AccountNameExt2, sum(BalanceAmount) + ISNULL ((SELECT " _
                        & "amount FROM qryAIS_ListOfAccountsPayableAndLiquidation WHERE qryAIS_ReportSchedulesConsolidated.JEVNumber = " _
                        & "dbo.qryAIS_ListOfAccountsPayableAndLiquidation.jevnumber AND ChildAccountCode = " _
                        & "qryAIS_ListOfAccountsPayableAndLiquidation.AccountCode), 0) " _
                        & "AS BalanceAmount, ReportNo, Particulars, MantYer, MotherFundType," _
                    & " MainCode, AccountName, AccountNameFull,ChildAccountCode from qryAIS_ReportSchedulesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") GROUP BY MainCode,AccountNameExt,AccountNameExt2,ReportNo, Particulars, MantYer, MotherFundType, AccountName, AccountNameFull,ChildAccountCode, JEVNumber order by MainCode,AccountNameExt")
        End If
    ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            opntbl.Open "select AccountCode from tblRef_AIS_RRRMain", opndbaseFMIS, adOpenStatic, adLockOptimistic
            Do Until opntbl.EOF
                opndbaseFMIS.Execute "update tblAIS_JEV set IsRRR=1 where AccountCode='" & Trim(opntbl!Accountcode) & "' and actioncode=1 and isapproved=1 and TypeOfFund in (SELECT FundName From tblRefBMS_Funds WHERE (FundName LIKE N'gen%') OR (FundName LIKE N'20%%')) and cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "month(jevdate) ELSE " _
                                    & "cast(substring(jevnumber, 8, 2) AS Int) END) AS nvarchar) + '-' " _
                                    & "+ cast((CASE jevtranstype WHEN 0 THEN " _
                                    & "substring(cast(Year(jevdate) " _
                                    & "AS nvarchar), 3, 2) ELSE substring(jevnumber, 5, 2) END) " _
                                    & "AS nvarchar) in (" & MantYerRange & ")"
            opntbl.MoveNext
            Loop
            
            opntbl.Close
            Set opntbl = Nothing
        'CrystalReportSked.Database.SetDataSource opndbaseFMIS.Execute("select AccountNameExt, AccountNameExt2, sum(BalanceAmount) + ISNULL ((SELECT " _
                        & "amount FROM qryAIS_ListOfAccountsPayableAndLiquidation WHERE qryAIS_ReportSchedulesProperand20.JEVNumber = " _
                        & "dbo.qryAIS_ListOfAccountsPayableAndLiquidation.jevnumber AND ChildAccountCode = " _
                        & "qryAIS_ListOfAccountsPayableAndLiquidation.AccountCode), 0) " _
                        & "AS BalanceAmount, ReportNo, Particulars, MantYer, TypeOfFund," _
                & " MainCode, AccountName, AccountNameFull,ChildAccountCode from qryAIS_ReportSchedulesProperand20 where " _
                & " MantYer in (" & MantYerRange & ") GROUP BY MainCode,AccountNameExt,AccountNameExt2,ReportNo, Particulars, MantYer, TypeOfFund, AccountName, AccountNameFull,ChildAccountCode, JEVNumber order by MainCode,AccountNameExt")
    End If
'    Call TransactionLogging("Print Preview", "RRR", "frmRRR")
    Call UnsetAnimation(frmRRR.Animation1)
    PreviewForm.Show vbModal
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".FlatBttn1_Click"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub


'***************************************************************************
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

Private Sub FlatBttn1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    On Error GoTo errHandler
    'PlaySound App.Path & "\sounds\HIGHLITE.WAV"
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".FlatBttn1_MouseMove"
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
'*  Author       : Errol Bagaipo
'*  Date         : 25 Oct 2006
'*  Note         :
'*  History      :
'***************************************************************************

Private Sub Form_Load()

    On Error GoTo errHandler
    'WindowsXPC1.InitSubClassing
    'mydll.CenterMe Me
    'Call DisplayOfficeUnderEcoEnt
    DTPicker2.Value = Now
    Call FundType(cboFundType)
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
    Set frmRRR = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

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
        err.Source = err.Source & "." & TypeName(Me) & ".optDateRange_Click"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub




