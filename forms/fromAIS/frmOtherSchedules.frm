VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmOtherSchedules 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   2580
   ClientTop       =   3705
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOtherSchedules.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6045
   Begin VB.CheckBox chkBeginningBalance 
      Caption         =   "Beginning Balance Only"
      Height          =   240
      Left            =   90
      TabIndex        =   20
      Top             =   3390
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fund Type"
      Height          =   2310
      Index           =   1
      Left            =   30
      TabIndex        =   9
      Top             =   1005
      Width           =   2490
      Begin VB.CheckBox chkProperAnd20 
         Caption         =   "GF Proper and 20% Dev't."
         Height          =   240
         Left            =   75
         TabIndex        =   19
         Top             =   1905
         Width           =   2385
      End
      Begin VB.CheckBox chkEco 
         Enabled         =   0   'False
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   1200
         Width           =   210
      End
      Begin VB.ComboBox cboEco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   405
         TabIndex        =   12
         Top             =   1155
         Width           =   2010
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         Height          =   240
         Left            =   75
         TabIndex        =   11
         Top             =   420
         Width           =   1530
      End
      Begin VB.ComboBox cboFundType 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   675
         Width           =   2370
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   2310
      Index           =   2
      Left            =   2535
      TabIndex        =   2
      Top             =   1005
      Width           =   3465
      Begin VB.ComboBox cboCriteriaAccountCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         TabIndex        =   16
         Top             =   1830
         Width           =   2220
      End
      Begin VB.ComboBox cboTransType 
         Height          =   315
         Left            =   1455
         TabIndex        =   15
         Top             =   1350
         Width           =   1920
      End
      Begin VB.OptionButton optDateRange 
         Caption         =   "Month Range"
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   3
         Top             =   1065
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   5
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
         Format          =   57344003
         CurrentDate     =   38695
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   6
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
         Format          =   57344003
         CurrentDate     =   38330
      End
      Begin VB.Label Label2 
         Caption         =   "Account Code"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1860
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Transaction Type"
         Height          =   255
         Left            =   135
         TabIndex        =   17
         Top             =   1380
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "From Month"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To Month"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&Preview"
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   3375
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   360
      Left            =   5055
      TabIndex        =   0
      Top             =   3375
      Width           =   960
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   930
      Top             =   8565
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      FrameControl    =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "CRITERIA FORM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   570
      Left            =   165
      TabIndex        =   14
      Top             =   210
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BorderColor     =   &H80000006&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   6030
   End
End
Attribute VB_Name = "frmOtherSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCriteriaAccountCode_KeyPress(KeyAscii As Integer)

    On Error GoTo errHandler
    KeyAscii = AutoFind(cboCriteriaAccountCode, KeyAscii, False)
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".cboCriteriaAccountCode_KeyPress"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

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
    

    If Len(Trim$(cboFundType.Text)) <> 0 Then
        chkEco.Enabled = True
        Call DisplayAccountCode(cboCriteriaAccountCode, cboFundType.Text)
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

Private Sub cboTransType_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoFind(cboTransType, KeyAscii, False)
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
        Call DisplayAccountCode(cboCriteriaAccountCode, "General Fund Proper")
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

    On Error GoTo errHandler
    strReportName = "OTHERSCHEDULE"
    If chkBeginningBalance.Value = 0 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                            & " from qryAIS_ReportOtherSchedules where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and JEVTransType=" & CByte(cboTransType.ItemData(cboTransType.ListIndex)) & " and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                            & " AND MantYer in (" & MantYerRange & ") and JEVTransType=" & CByte(cboTransType.ItemData(cboTransType.ListIndex)) & " GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, AccountNameFull order by AccountCode")
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                                & " from qryAIS_ReportOtherSchedules where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and JEVTransType=" & CByte(cboTransType.ItemData(cboTransType.ListIndex)) & " and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") GROUP BY JEVDate, JEVNumber, AccountCode, Particulars,ResponsibilityCenter, MantYer, TypeOfFund, AccountNameFull order by AccountCode")
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                        & " from qryAIS_ReportOtherSchedulesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and JEVTransType=" & CByte(cboTransType.ItemData(cboTransType.ListIndex)) & " and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                        & " AND MantYer in (" & MantYerRange & ") GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, MantYer, MotherFundType, AccountNameFull order by AccountCode")
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                    & " from qryAIS_ReportOtherSchedulesProperand20 where AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                    & " and JEVTransType=" & CByte(cboTransType.ItemData(cboTransType.ListIndex)) & " and MantYer in (" & MantYerRange & ") GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, MantYer, TypeOfFund, AccountNameFull order by AccountCode")
        End If
    Else
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                            & " from qryAIS_ReportOtherSchedules where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                            & " AND MantYer in (" & MantYerRange & ") and upper(Particulars) like '%" & "BEGINNING BALANCE" & "%' GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, AccountNameFull order by AccountCode")
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                                & " from qryAIS_ReportOtherSchedules where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and upper(Particulars) like '%" & "BEGINNING BALANCE" & "%' GROUP BY JEVDate, JEVNumber, AccountCode, Particulars,ResponsibilityCenter, MantYer, TypeOfFund, AccountNameFull order by AccountCode")
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                        & " from qryAIS_ReportOtherSchedulesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                        & " AND MantYer in (" & MantYerRange & ") and upper(Particulars) like '%" & "BEGINNING BALANCE" & "%' GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, MantYer, MotherFundType, AccountNameFull order by AccountCode")
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            CrystalReportOtherSked.Database.SetDataSource opndbaseFMIS.Execute("Select AccountCode,JEVDate, JEVNumber, AccountNameFull,Particulars, SUM(DebitAmount) AS DebitAmount, SUM(CreditAmount) AS CreditAmount" _
                    & " from qryAIS_ReportOtherSchedulesProperand20 where AccountCode like '" & Trim(cboCriteriaAccountCode.Text) & "%'" _
                    & " and MantYer in (" & MantYerRange & ") and upper(Particulars) like '%" & "BEGINNING BALANCE" & "%' GROUP BY JEVDate, JEVNumber, AccountCode, Particulars, MantYer, TypeOfFund, AccountNameFull order by AccountCode")
        End If
    End If
   ' Call TransactionLogging("Print Preview", "Schedule", "frmSchedule")
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
    Call DisplayALLTranstype
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
    WindowsXPC1.EndWinXPCSubClassing
    Set frmOtherSchedules = Nothing
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
Private Sub DisplayAccountCode(ByVal cboAccountCodeName As ComboBox, ByVal FundName As String)
    '***************************************************************************
    '*  Name         : DisplayAccountCode
    '*  Description  :
    '*  Parameters   : cboAccountCodeName As ComboBox, FundName As String
    '*  Returns      : Nothing
    '*  Author       : Errol Bagaipo
    '*  Date         : 25 Oct 2006
    '***************************************************************************


    On Error GoTo errHandler
    Dim opntbl As New ADODB.Recordset
    Dim x As Integer
    
    cboAccountCodeName.Clear

    If chkConsolidated.Value = 0 Then
        opntbl.Open "select ChildAccountCode,FMISAccountCode from tblREF_AIS_ChartofAccounts group by childaccountcode,fundtype,FMISAccountCode having upper(fundtype)='" & UCase$(Trim$(FundName)) & "' order by ChildAccountCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            Do Until opntbl.EOF
                If IsNull(opntbl!ChildAccountCode) = False Then
                    cboAccountCodeName.AddItem opntbl!ChildAccountCode
                    cboAccountCodeName.ItemData(x) = opntbl!FmisAccountcode
                    x = x + 1
                End If
                opntbl.MoveNext
            Loop
        End If
    Else
        opntbl.Open "select DISTINCT ChildAccountCode from tblREF_AIS_ChartofAccounts group by childaccountcode,fundtype having upper(fundtype) in ('" & UCase$(Trim$("GENERAL FUND PROPER")) & "','" & UCase$(Trim$("20% DEVELOPMENT FUND")) & "','" & UCase$(Trim$("ECONOMIC ENTERPRISES")) & "') order by ChildAccountCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            Do Until opntbl.EOF
                If IsNull(opntbl!ChildAccountCode) = False Then
                    cboAccountCodeName.AddItem opntbl!ChildAccountCode
                End If
                opntbl.MoveNext
            Loop
        End If
    End If
    opntbl.Close
    Set opntbl = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".DisplayAccountCode"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

Private Sub DisplayALLTranstype()

Dim opntbl As New ADODB.Recordset
Dim xx As Byte
    
    cboTransType.Clear
    opntbl.Open "select * from tblREF_AIS_TransType where Transno<>0 order by Transtype asc", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opntbl.RecordCount > 0 Then
        Do Until opntbl.EOF
            cboTransType.AddItem opntbl!Transtype
            cboTransType.ItemData(xx) = opntbl!transNo
            xx = xx + 1
            opntbl.MoveNext
        Loop
    End If
    
    opntbl.Close
    Set opntbl = Nothing
End Sub



