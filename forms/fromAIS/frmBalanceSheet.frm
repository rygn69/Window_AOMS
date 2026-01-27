VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmBalanceSheet 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   6330
   ClientTop       =   3630
   ClientWidth     =   5130
   Icon            =   "frmBalanceSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdPreview2 
      Caption         =   "Preview"
      Height          =   360
      Left            =   3120
      TabIndex        =   20
      Top             =   3405
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2310
      Index           =   2
      Left            =   2550
      TabIndex        =   8
      Top             =   1005
      Width           =   2550
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   11
         Top             =   1065
         Width           =   2400
      End
      Begin VB.OptionButton optDateRange 
         Caption         =   "Date Range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   10
         Top             =   135
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.CheckBox chkClosing 
         Caption         =   "Include CLOSING Entries?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   9
         Top             =   1605
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   12
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
         Format          =   57606147
         CurrentDate     =   38838
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   13
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
         Format          =   57606147
         CurrentDate     =   38868
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   435
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4095
      TabIndex        =   7
      Top             =   3405
      Width           =   960
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fund Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   1005
      Width           =   2490
      Begin VB.ComboBox cboFundType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   675
         Width           =   2370
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   4
         Top             =   420
         Width           =   1530
      End
      Begin VB.ComboBox cboEco 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   405
         TabIndex        =   3
         Top             =   1155
         Width           =   2010
      End
      Begin VB.CheckBox chkEco 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Width           =   210
      End
      Begin VB.CheckBox chkProperAnd20 
         Caption         =   "GF Proper and 20% Dev't."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   1770
         Width           =   2385
      End
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
   Begin VB.Frame Frame7 
      Height          =   35
      Left            =   -90
      TabIndex        =   17
      Top             =   840
      Width           =   11220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Define criteria prior to preview a Balance Sheet Report."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   210
      TabIndex        =   19
      Top             =   480
      Width           =   4755
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE SHEEET"
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
      Left            =   210
      TabIndex        =   18
      Top             =   210
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -30
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "frmBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dblPriorAmount As Double
Public dblCurrentOp As Double
Public dblTransferPI As Double
Public dblGovernmentJan As Double
Public dblPriorDebit As Double
Public dblPriorCredit As Double
Public dblGEJanDebit As Double
Public dblGEJanCredit As Double
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
    
    Call SetAnimation(frmBalanceSheet.Animation1)
    If Len(Trim$(cboFundType.Text)) <> 0 Then
        chkEco.Enabled = True

        opntbl.Open "SELECT Description, ECOCode From tblref_ECOCode ORDER BY Description", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            cboEco.Clear
            Do Until opntbl.EOF
                If IsNull(opntbl!Description) = False Then

                    cboEco.AddItem Trim$(opntbl!Description)
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
    Call UnsetAnimation(frmBalanceSheet.Animation1)
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

Private Sub cmdPreview2_Click()
    Call UpdatetblREF_AIS_BalanceSheetFormat
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

Private Sub Command2_Click()
End Sub

Private Sub DTPicker1_Change()
    DTPicker1.Value = Month(DTPicker1.Value) & "/" & "1" & "/" & Year(DTPicker1.Value)
End Sub

Private Sub DTPicker2_Change()
    'DTPicker2.Value = Month(DTPicker2.Value) & "/" & GetEndDateoftheMonth(DTPicker2.Value) & "/" & Year(DTPicker2.Value)
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
    Dim opntbl As New ADODB.Recordset
    Dim dblIncome As Double
    Dim dblExpenses As Double
    Dim opntbl2 As New ADODB.Recordset
    Dim strMant As String, strYer As String
    Dim strSQL As String

    On Error GoTo errHandler
    strReportName = "BS"
    Call SetAnimation(frmBalanceSheet.Animation1)
    opndbaseFMIS.Execute "truncate table tblAIS_TempBS"
    If chkClosing.Value = 1 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                            & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                    'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute(strSQL)
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                                & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                        'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            'If optDateRange.Value = True Then    'if date range
                strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                        & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
                'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
            'End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                    & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetProperand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
            'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
        End If
    
        opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        
        
        
        '------------------------------------------------for Statement of GE
        '-------------------------------------------------------------------
        '----------------------------------------------------------------------
        opndbaseFMIS.Execute "truncate table tblAIS_tempSGE"
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    
                    'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        'dblGovernmentJan = opntbl!BalanceDiff
                        'dblGEJanCredit = opntbl!CreditBalance
                        'dblGEJanDebit = opntbl!DebitBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                    Else
                        'dblGovernmentJan = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblIncome = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblExpenses = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblPriorAmount = opntbl!BalanceDiff
                        'dblPriorDebit = opntbl!DebitBalance
                        'dblPriorCredit = opntbl!CreditBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        dblPriorAmount = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                                    
                    opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "SELECT SGEID, " _
                                    & "SUM(DebitBalance) AS BalanceDiff FROM     " _
                                    & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                    & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                    & "AS smallint) <> 1) and MainCode='501' AND (CAST(SGEID AS " _
                                    & "smallint) <> 2) AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, " _
                                        & "SUM(DebitBalance) AS DebitBalance,0 FROM     " _
                                        & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                    
                        opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                        End If
                        
                        opntbl2.Close
                        Set opntbl2 = Nothing
                    End If
                    
                    'dblCurrentOp = dblIncome - dblExpenses
                    
                    'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    'If opntbl.RecordCount > 0 Then
                    '    dblTransferPI = opntbl!BalanceDiff
                    'End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            'dblGovernmentJan = opntbl!BalanceDiff
                            'dblGEJanCredit = opntbl!CreditBalance
                            'dblGEJanDebit = opntbl!DebitBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                        Else
                            'dblGovernmentJan = 0
                            'dblGEJanCredit = 0
                            'dblGEJanDebit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Income" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblPriorAmount = opntbl!BalanceDiff
                            'dblPriorDebit = opntbl!DebitBalance
                            'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            dblPriorAmount = 0
                            'dblPriorDebit = 0
                            'dblPriorCredit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                                                
                        opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                        
                            opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                            If opntbl.RecordCount > 0 Then
                                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                            End If
                            
                            opntbl2.Close
                            Set opntbl2 = Nothing
                        End If
                        'dblCurrentOp = dblIncome - dblExpenses
                            
                        'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and MainCode='" & "260" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        'If opntbl.RecordCount > 0 Then
                        '    dblTransferPI = opntbl!BalanceDiff
                        'End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                    
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    'dblGovernmentJan = opntbl!BalanceDiff
                    'dblGEJanCredit = opntbl!CreditBalance
                    'dblGEJanDebit = opntbl!DebitBalance
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                Else
                    'dblGovernmentJan = 0
                    'dblGEJanCredit = 0
                    'dblGEJanDebit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblPriorAmount = opntbl!BalanceDiff
                    'dblPriorDebit = opntbl!DebitBalance
                    'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    dblPriorAmount = 0
                    'dblPriorDebit = 0
                    'dblPriorCredit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                                    
                opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType"
                
                    opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                & "FROM " _
                                & "dbo.qryAIS_StatementofGE WHERE " _
                                & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                    End If
                    
                    opntbl2.Close
                    Set opntbl2 = Nothing
                End If
                'dblCurrentOp = dblIncome - dblExpenses
                    
                'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                '        & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                'If opntbl.RecordCount > 0 Then
                '    dblTransferPI = opntbl!BalanceDiff
                'End If
                    
                opntbl.Close
                Set opntbl = Nothing
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where  " _
                    & " MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                'dblGovernmentJan = opntbl!BalanceDiff
                'dblGEJanCredit = opntbl!CreditBalance
                'dblGEJanDebit = opntbl!DebitBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
            Else
                'dblGovernmentJan = 0
                'dblGEJanCredit = 0
                'dblGEJanDebit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Income" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblIncome = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Expenses" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblExpenses = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Properand20 where  " _
                    & "  MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblPriorAmount = opntbl!BalanceDiff
                'dblPriorDebit = opntbl!DebitBalance
                'dblPriorCredit = opntbl!CreditBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                dblPriorAmount = 0
                'dblPriorDebit = 0
                'dblPriorCredit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
                                    
            opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
            End If
            opntbl.Close
            Set opntbl = Nothing
                    
            opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                            & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                            & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                            & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                            & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                            & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") GROUP BY SGEID", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") GROUP BY SGEID"
            
                opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                            & "FROM " _
                            & "dbo.qryAIS_StatementofGE WHERE " _
                            & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                            & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                End If
                opntbl2.Close
                Set opntbl2 = Nothing
            End If
            
            'dblCurrentOp = dblIncome - dblExpenses
            
            'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceProperand20 where MainCode='" & "260" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
            'If opntbl.RecordCount > 0 Then
            '    dblTransferPI = opntbl!BalanceDiff
            'End If
            
            opntbl.Close
            Set opntbl = Nothing
        End If
        
        opntbl.Open "SELECT SUM(LeftColumn) AS LeftColumn, SUM(RightColumn) AS RightColumn From tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            strSQL = "SELECT     FourthLevelGroup, ThirdLevelGroup, " _
                        & "SecondLevelGroup, FirstLevelGroup, AccountName," & "" & CCur(opntbl!RightColumn) & ", " _
                        & "AccountCode, OrderInBS FROM         " _
                        & "tblREF_AIS_ChartOfAccountsMother WHERE     (AccountCode " _
                        & "= '501')"
            opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        Else
        
        End If
        opntbl.Close
        Set opntbl = Nothing
        
    ElseIf chkClosing.Value = 0 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                            & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                    'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute(strSQL)
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                                & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                        'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            'If optDateRange.Value = True Then    'if date range
                strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                        & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
                'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
            'End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                    & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetProperand20 where " _
                    & " MantYer in (" & MantYerRange & ") and IsClosed=0 group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
            'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
        End If
    
        opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        
        
        
        '------------------------------------------------for Statement of GE
        '-------------------------------------------------------------------
        '----------------------------------------------------------------------
        opndbaseFMIS.Execute "truncate table tblAIS_tempSGE"
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    
                    'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        'dblGovernmentJan = opntbl!BalanceDiff
                        'dblGEJanCredit = opntbl!CreditBalance
                        'dblGEJanDebit = opntbl!DebitBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                    Else
                        'dblGovernmentJan = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblIncome = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblExpenses = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblPriorAmount = opntbl!BalanceDiff
                        'dblPriorDebit = opntbl!DebitBalance
                        'dblPriorCredit = opntbl!CreditBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        dblPriorAmount = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                                    
                    opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "SELECT SGEID, " _
                                    & "SUM(DebitBalance) AS BalanceDiff FROM     " _
                                    & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                    & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                    & "AS smallint) <> 1) and MainCode='501' AND (CAST(SGEID AS " _
                                    & "smallint) <> 2) AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, " _
                                        & "SUM(DebitBalance) AS DebitBalance,0 FROM     " _
                                        & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                    
                        opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                        End If
                        
                        opntbl2.Close
                        Set opntbl2 = Nothing
                    End If
                    
                    'dblCurrentOp = dblIncome - dblExpenses
                    
                    'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    'If opntbl.RecordCount > 0 Then
                    '    dblTransferPI = opntbl!BalanceDiff
                    'End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            'dblGovernmentJan = opntbl!BalanceDiff
                            'dblGEJanCredit = opntbl!CreditBalance
                            'dblGEJanDebit = opntbl!DebitBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                        Else
                            'dblGovernmentJan = 0
                            'dblGEJanCredit = 0
                            'dblGEJanDebit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Income" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblPriorAmount = opntbl!BalanceDiff
                            'dblPriorDebit = opntbl!DebitBalance
                            'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            dblPriorAmount = 0
                            'dblPriorDebit = 0
                            'dblPriorCredit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                                                
                        opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                        
                            opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                            If opntbl.RecordCount > 0 Then
                                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                            End If
                            
                            opntbl2.Close
                            Set opntbl2 = Nothing
                        End If
                        'dblCurrentOp = dblIncome - dblExpenses
                            
                        'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and MainCode='" & "260" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        'If opntbl.RecordCount > 0 Then
                        '    dblTransferPI = opntbl!BalanceDiff
                        'End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                    
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    'dblGovernmentJan = opntbl!BalanceDiff
                    'dblGEJanCredit = opntbl!CreditBalance
                    'dblGEJanDebit = opntbl!DebitBalance
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                Else
                    'dblGovernmentJan = 0
                    'dblGEJanCredit = 0
                    'dblGEJanDebit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblPriorAmount = opntbl!BalanceDiff
                    'dblPriorDebit = opntbl!DebitBalance
                    'dblPriorCredit = opntbl!CreditBalance
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    dblPriorAmount = 0
                    'dblPriorDebit = 0
                    'dblPriorCredit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                                    
                opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType"
                
                    opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                & "FROM " _
                                & "dbo.qryAIS_StatementofGE WHERE " _
                                & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                    End If
                    
                    opntbl2.Close
                    Set opntbl2 = Nothing
                End If
                'dblCurrentOp = dblIncome - dblExpenses
                    
                'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                '        & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                'If opntbl.RecordCount > 0 Then
                '    dblTransferPI = opntbl!BalanceDiff
                'End If
                    
                opntbl.Close
                Set opntbl = Nothing
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where  " _
                    & " MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                'dblGovernmentJan = opntbl!BalanceDiff
                'dblGEJanCredit = opntbl!CreditBalance
                'dblGEJanDebit = opntbl!DebitBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
            Else
                'dblGovernmentJan = 0
                'dblGEJanCredit = 0
                'dblGEJanDebit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Income" & "'" _
                    & " and MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblIncome = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Expenses" & "'" _
                    & " and MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblExpenses = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Properand20 where  " _
                    & "  MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblPriorAmount = opntbl!BalanceDiff
                'dblPriorDebit = opntbl!DebitBalance
                'dblPriorCredit = opntbl!CreditBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                dblPriorAmount = 0
                'dblPriorDebit = 0
                'dblPriorCredit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
                                    
            opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
            End If
            opntbl.Close
            Set opntbl = Nothing
                    
            opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                            & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                            & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                            & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                            & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                            & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY SGEID", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY SGEID"
            
                opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                            & "FROM " _
                            & "dbo.qryAIS_StatementofGE WHERE " _
                            & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                            & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                End If
                
                opntbl2.Close
                Set opntbl2 = Nothing
            End If
            
            'dblCurrentOp = dblIncome - dblExpenses
            
            'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceProperand20 where MainCode='" & "260" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
            'If opntbl.RecordCount > 0 Then
            '    dblTransferPI = opntbl!BalanceDiff
            'End If
            
            opntbl.Close
            Set opntbl = Nothing
        End If
        
        opntbl.Open "SELECT SUM(LeftColumn) AS LeftColumn, SUM(RightColumn) AS RightColumn From tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            strSQL = "SELECT     FourthLevelGroup, ThirdLevelGroup, " _
                        & "SecondLevelGroup, FirstLevelGroup, AccountName," & "" & CCur(opntbl!RightColumn) & ", " _
                        & "AccountCode, OrderInBS FROM         " _
                        & "tblREF_AIS_ChartOfAccountsMother WHERE     (AccountCode " _
                        & "= '501')"
            opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        Else
        
        End If
    End If
    opntbl.Close
    Set opntbl = Nothing
    CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute("select * from tblAIS_TempBS WHERE (BalanceAmount <> 0) order by OrderInBS")
    Call TransactionLogging("Print Preview", "Balance Sheet", "frmBalanceSheet")
    Call UnsetAnimation(frmBalanceSheet.Animation1)
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

Private Sub FlatBttn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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
    'MyDLL.CenterMe Me
    'Call DisplayOfficeUnderEcoEnt
    'DTPicker2.Value = Month(Now) & "/" & GetEndDateoftheMonth(Now()) & "/" & Year(ServerDate())
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
    WindowsXPC1.EndWinXPCSubClassing
    Set frmBalanceSheet = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        err.Source = err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = err
     
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
        err.Source = err.Source & "." & TypeName(Me) & ".optDateRange_Click"
        Set .Error = err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub
'Public Function GetTotalGovernmentEquity() As Double
'Dim opntbl As New ADODB.Recordset
'Dim dblIncome As Double
'Dim dblExpenses As Double
'Dim dblGovernmentJan As Double
'Dim dblPriorAmount As Double
'Dim dblCurrentOp As Double
'Dim dblTransferPI As Double
'
'    If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
'        If optDateRange.Value = True Then    'if date range
'            If chkEco.Value = 0 Then
'                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff,sum(DebitBalance) DebitBalance,sum(CreditBalance) CreditBalance from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
'                        & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName", opndbaseFMIS, adOpenStatic, adLockOptimistic
'                If opntbl.RecordCount > 0 Then
'                    dblGovernmentJan = opntbl!BalanceDiff
'                Else
'                    dblGovernmentJan = 0
'                End If
'
'                opntbl.Close
'                Set opntbl = Nothing
'
'                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff,sum(DebitBalance) DebitBalance,sum(CreditBalance) CreditBalance from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
'                        & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName", opndbaseFMIS, adOpenStatic, adLockOptimistic
'                If opntbl.RecordCount > 0 Then
'                    dblPriorAmount = opntbl!BalanceDiff
'                Else
'                    dblPriorAmount = 0
'                End If
'
'                opntbl.Close
'                Set opntbl = Nothing
'
'                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
'                        & " AND MantYer in (" & MantYerRange & ") group by FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
'                If opntbl.RecordCount > 0 Then
'                    dblIncome = opntbl!BalanceDiff
'                End If
'
'                opntbl.Close
'                Set opntbl = Nothing
'
'                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
'                        & " AND MantYer in (" & MantYerRange & ") group by FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
'                If opntbl.RecordCount > 0 Then
'                    dblExpenses = opntbl!BalanceDiff
'                End If
'
'                opntbl.Close
'                Set opntbl = Nothing
'
'                dblCurrentOp = dblIncome - dblExpenses
'
'                opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
'                        & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
'                If opntbl.RecordCount > 0 Then
'                    dblTransferPI = opntbl!BalanceDiff
'                End If
'
'                opntbl.Close
'                Set opntbl = Nothing
'
'                GetTotalGovernmentEquity = Format(dblGovernmentJan + ((dblCurrentOp + dblPriorAmount) - dblTransferPI), "#,##0.00")
'
'End Function
'
'Public Function GetTotalLiabilities() As Double
'Dim opntbl As New ADODB.Recordset
'
'    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportBalanceSheetFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Liabilities" & "'" _
'            & " AND MantYer in (" & MantYerRange & ") group by FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If opntbl.RecordCount > 0 Then
'        GetTotalLiabilities = opntbl!BalanceDiff
'    End If
'
'    opntbl.Close
'    Set opntbl = Nothing
'
'End Function
'
Private Sub UpdatetblREF_AIS_BalanceSheetFormat()
    Dim opntbl As New ADODB.Recordset
    Dim dblIncome As Double
    Dim dblExpenses As Double
    Dim opntbl2 As New ADODB.Recordset
    Dim strMant As String, strYer As String
    Dim strSQL As String
    
    strReportName = "BS"
    Call SetAnimation(frmBalanceSheet.Animation1)
    opndbaseFMIS.Execute "update tblREF_AIS_BalanceSheetFormat set Amount=0"
    If chkClosing.Value = 1 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    'strSql = "select SUM(BalanceDiff) AS BalanceDiff,MainCode from qryReportBalanceSheetFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode HAVING (sum(BalanceDiff) <> 0)"
                    strSQL = "select MainCode,sum(BalanceDiff) BalanceDiff from qryAIS_ReportBalanceSheet where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode HAVING (sum(BalanceDiff) <> 0)"
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        'strSql = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                                & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                        strSQL = "select MainCode,sum(BalanceDiff) BalanceDiff from qryAIS_ReportBalanceSheet where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode having (sum(BalanceDiff) <> 0)"
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
                'strSql = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                        & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
                strSQL = "select MainCode,sum(BalanceDiff) BalanceDiff from qryAIS_ReportBalanceSheet where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MainCode having (sum(BalanceDiff) <> 0)"
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            'strSql = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                    & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetProperand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
            strSQL = "select MainCode,sum(BalanceDiff) BalanceDiff from qryAIS_ReportBalanceSheet where (NewFundCode = N'101')" _
                    & " AND MantYer in (" & MantYerRange & ") group by MainCode HAVING (sum(BalanceDiff) <> 0)"
        End If
    
        opntbl.Open strSQL, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            Do Until opntbl.EOF
                opndbaseFMIS.Execute "Update tblREF_AIS_BalanceSheetFormat set amount=" & CCur(opntbl!BalanceDiff) & " where AccountCode='" & Trim(opntbl!MainCode) & "'"
            opntbl.MoveNext
            Loop
        End If
        
        opntbl.Close
        Set opntbl = Nothing
        
        'opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        
        
        '------------------------------------------------for Statement of GE
        '-------------------------------------------------------------------
        '----------------------------------------------------------------------
        opndbaseFMIS.Execute "truncate table tblAIS_tempSGE"
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                    End If
                    
                    
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblIncome = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblExpenses = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblPriorAmount = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        dblPriorAmount = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                    End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                                    
                    opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                    End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "SELECT SGEID, " _
                                    & "SUM(DebitBalance) AS BalanceDiff FROM     " _
                                    & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                    & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                    & "AS smallint) <> 1) and MainCode='501' AND (CAST(SGEID AS " _
                                    & "smallint) <> 2) AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, " _
                                        & "SUM(DebitBalance) AS DebitBalance,0 FROM     " _
                                        & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                    
                        opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                        End If
                        
                        opntbl2.Close
                        Set opntbl2 = Nothing
                    End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Income" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblPriorAmount = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            dblPriorAmount = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                        End If
                        
                        opntbl.Close
                        Set opntbl = Nothing
                                                
                        opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                        
                            opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                            If opntbl.RecordCount > 0 Then
                                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                            End If
                            
                            opntbl2.Close
                            Set opntbl2 = Nothing
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                    
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    'dblGovernmentJan = opntbl!BalanceDiff
                    'dblGEJanCredit = opntbl!CreditBalance
                    'dblGEJanDebit = opntbl!DebitBalance
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                Else
                    'dblGovernmentJan = 0
                    'dblGEJanCredit = 0
                    'dblGEJanDebit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                End If
                    
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                    
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                    
                opntbl.Close
                Set opntbl = Nothing
                
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblPriorAmount = opntbl!BalanceDiff
                    'dblPriorDebit = opntbl!DebitBalance
                    'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    dblPriorAmount = 0
                    'dblPriorDebit = 0
                    'dblPriorCredit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                End If
                                    
                opntbl.Close
                Set opntbl = Nothing
                                    
                opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                End If
                
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType"
                
                    opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                & "FROM " _
                                & "dbo.qryAIS_StatementofGE WHERE " _
                                & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                    End If
                    
                    opntbl2.Close
                    Set opntbl2 = Nothing
                End If
                'dblCurrentOp = dblIncome - dblExpenses
                    
                'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                '        & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                'If opntbl.RecordCount > 0 Then
                '    dblTransferPI = opntbl!BalanceDiff
                'End If
                    
                opntbl.Close
                Set opntbl = Nothing
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where  " _
                    & " MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                'dblGovernmentJan = opntbl!BalanceDiff
                'dblGEJanCredit = opntbl!CreditBalance
                'dblGEJanDebit = opntbl!DebitBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
            Else
                'dblGovernmentJan = 0
                'dblGEJanCredit = 0
                'dblGEJanDebit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
            End If
            
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Income" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblIncome = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Expenses" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblExpenses = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Properand20 where  " _
                    & "  MantYer in (" & MantYerRange & ") group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblPriorAmount = opntbl!BalanceDiff
                'dblPriorDebit = opntbl!DebitBalance
                'dblPriorCredit = opntbl!CreditBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                dblPriorAmount = 0
                'dblPriorDebit = 0
                'dblPriorCredit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
            End If
            
            opntbl.Close
            Set opntbl = Nothing
                                    
            opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
            End If
            
            opntbl.Close
            Set opntbl = Nothing
                    
            opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                            & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                            & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                            & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                            & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                            & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") GROUP BY SGEID", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") GROUP BY SGEID"
            
                opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                            & "FROM " _
                            & "dbo.qryAIS_StatementofGE WHERE " _
                            & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                            & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                End If
                
                opntbl2.Close
                Set opntbl2 = Nothing
            End If
            
            'dblCurrentOp = dblIncome - dblExpenses
            
            'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceProperand20 where MainCode='" & "260" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
            'If opntbl.RecordCount > 0 Then
            '    dblTransferPI = opntbl!BalanceDiff
            'End If
            
            opntbl.Close
            Set opntbl = Nothing
        End If
        
        opntbl.Open "SELECT SUM(LeftColumn) AS LeftColumn, SUM(RightColumn) AS RightColumn From tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            'strSQL = "SELECT     FourthLevelGroup, ThirdLevelGroup, " _
                        & "SecondLevelGroup, FirstLevelGroup, AccountName," & "" & CCur(opntbl!RightColumn) & ", " _
                        & "AccountCode, OrderInBS FROM         " _
                        & "tblREF_AIS_ChartOfAccountsMother WHERE     (AccountCode " _
                        & "= '501')"
            'opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
            opndbaseFMIS.Execute "update tblREF_AIS_BalanceSheetFormat set Amount=" & CCur(opntbl!RightColumn) & " where (AccountCode = '501')"
        Else
        
        End If
        opntbl.Close
        Set opntbl = Nothing
    ElseIf chkClosing.Value = 0 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                            & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                    'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute(strSQL)
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                                & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName,OrderInBS order by OrderInBS,MainCode"
                        'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            'If optDateRange.Value = True Then    'if date range
                strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                        & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
                'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
            'End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            strSQL = "select FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName," _
                    & " SUM(BalanceDiff) AS BalanceDiff,MainCode,OrderInBS from qryReportBalanceSheetProperand20 where " _
                    & " MantYer in (" & MantYerRange & ") and IsClosed=0 group by OrderInBS,MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, MainAccountName order by OrderInBS,MainCode"
            'CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute()
        End If
        opntbl.Open strSQL, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            Do Until opntbl.EOF
                opndbaseFMIS.Execute "Update tblREF_AIS_BalanceSheetFormat set amount=" & CCur(opntbl!BalanceDiff) & " where AccountCode='" & Trim(opntbl!MainCode) & "'"
            opntbl.MoveNext
            Loop
        End If
        
        opntbl.Close
        Set opntbl = Nothing
        
        'opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        
        
        
        '------------------------------------------------for Statement of GE
        '-------------------------------------------------------------------
        '----------------------------------------------------------------------
        opndbaseFMIS.Execute "truncate table tblAIS_tempSGE"
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    
                    'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        'dblGovernmentJan = opntbl!BalanceDiff
                        'dblGEJanCredit = opntbl!CreditBalance
                        'dblGEJanDebit = opntbl!DebitBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                    Else
                        'dblGovernmentJan = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblIncome = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblExpenses = opntbl!BalanceDiff
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                    Else
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                            & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        dblPriorAmount = opntbl!BalanceDiff
                        'dblPriorDebit = opntbl!DebitBalance
                        'dblPriorCredit = opntbl!CreditBalance
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                    Else
                        dblPriorAmount = 0
                        'dblGEJanCredit = 0
                        'dblGEJanDebit = 0
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                                    
                    opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                    End If
                    opntbl.Close
                    Set opntbl = Nothing
                    
                    opntbl.Open "SELECT SGEID, " _
                                    & "SUM(DebitBalance) AS BalanceDiff FROM     " _
                                    & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                    & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                    & "AS smallint) <> 1) and MainCode='501' AND (CAST(SGEID AS " _
                                    & "smallint) <> 2) AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, " _
                                        & "SUM(DebitBalance) AS DebitBalance,0 FROM     " _
                                        & "    dbo.qryReportTrialBalanceFinal WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                    
                        opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                        End If
                        
                        opntbl2.Close
                        Set opntbl2 = Nothing
                    End If
                    
                    'dblCurrentOp = dblIncome - dblExpenses
                    
                    'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceFinal where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                            & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    'If opntbl.RecordCount > 0 Then
                    '    dblTransferPI = opntbl!BalanceDiff
                    'End If
                    
                    opntbl.Close
                    Set opntbl = Nothing
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            'dblGovernmentJan = opntbl!BalanceDiff
                            'dblGEJanCredit = opntbl!CreditBalance
                            'dblGEJanDebit = opntbl!DebitBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                        Else
                            'dblGovernmentJan = 0
                            'dblGEJanCredit = 0
                            'dblGEJanDebit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Income" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                        Else
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                            
                        opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684 where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            dblPriorAmount = opntbl!BalanceDiff
                            'dblPriorDebit = opntbl!DebitBalance
                            'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                        Else
                            dblPriorAmount = 0
                            'dblPriorDebit = 0
                            'dblPriorCredit = 0
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                        End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                        
                        opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                        End If
                        opntbl.Close
                        Set opntbl = Nothing
                                                
                        opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If opntbl.RecordCount > 0 Then
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                        & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                        & "    dbo.qryReportTrialBalanceWithRC WHERE     " _
                                        & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                        & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                        & "smallint) <> 2) and MainCode='501' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, typeoffund"
                        
                            opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                            If opntbl.RecordCount > 0 Then
                                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                        & "FROM " _
                                        & "dbo.qryAIS_StatementofGE WHERE " _
                                        & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                        & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                            End If
                            
                            opntbl2.Close
                            Set opntbl2 = Nothing
                        End If
                        'dblCurrentOp = dblIncome - dblExpenses
                            
                        'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceWithRC where upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' and MainCode='" & "260" & "'" _
                                & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                        'If opntbl.RecordCount > 0 Then
                        '    dblTransferPI = opntbl!BalanceDiff
                        'End If
                            
                        opntbl.Close
                        Set opntbl = Nothing
                    
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName order by MainCode")
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    'dblGovernmentJan = opntbl!BalanceDiff
                    'dblGEJanCredit = opntbl!CreditBalance
                    'dblGEJanDebit = opntbl!DebitBalance
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
                Else
                    'dblGovernmentJan = 0
                    'dblGEJanCredit = 0
                    'dblGEJanDebit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
                End If
                    
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Income" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblIncome = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and FourthLevelGroup='" & "Expenses" & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblExpenses = opntbl!BalanceDiff
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
                Else
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                
                opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Consolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                        & " AND MantYer in (" & MantYerRange & ") and IsClosed=0 group by MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    dblPriorAmount = opntbl!BalanceDiff
                    'dblPriorDebit = opntbl!DebitBalance
                    'dblPriorCredit = opntbl!CreditBalance
                            opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
                Else
                    dblPriorAmount = 0
                    'dblPriorDebit = 0
                    'dblPriorCredit = 0
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
                End If
                opntbl.Close
                Set opntbl = Nothing
                                    
                opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
                End If
                opntbl.Close
                Set opntbl = Nothing
                    
                opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceConsolidated WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 and upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' GROUP BY SGEID, MotherFundType"
                
                    opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opntbl.RecordCount > 0 Then
                        opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                                & "FROM " _
                                & "dbo.qryAIS_StatementofGE WHERE " _
                                & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                                & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                    End If
                    
                    opntbl2.Close
                    Set opntbl2 = Nothing
                End If
                'dblCurrentOp = dblIncome - dblExpenses
                    
                'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceConsolidated where upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and MainCode='" & "260" & "'" _
                '        & " AND MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
                'If opntbl.RecordCount > 0 Then
                '    dblTransferPI = opntbl!BalanceDiff
                'End If
                    
                opntbl.Close
                Set opntbl = Nothing
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            'CrystalReportSGE.Database.SetDataSource opndbaseFMIS.Execute("select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where " _
                    & " MantYer in (" & MantYerRange & ") group by MainCode,FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, MainAccountName")
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity501Properand20 where  " _
                    & " MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                'dblGovernmentJan = opntbl!BalanceDiff
                'dblGEJanCredit = opntbl!CreditBalance
                'dblGEJanDebit = opntbl!DebitBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0," & CCur(opntbl!BalanceDiff) & ")"
            Else
                'dblGovernmentJan = 0
                'dblGEJanCredit = 0
                'dblGEJanDebit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (0,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Income" & "'" _
                    & " and MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblIncome = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementOfIncomeAndExpensesProperand20 where FourthLevelGroup='" & "Expenses" & "'" _
                    & " and MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund,FourthLevelGroup", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblExpenses = opntbl!BalanceDiff
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1," & CCur(-opntbl!BalanceDiff) & ",0)"
            Else
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (1,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
            
            opntbl.Open "select SUM(BalanceDiff) AS BalanceDiff from qryReportStatementofGovernmentEquity684Properand20 where  " _
                    & "  MantYer in (" & MantYerRange & ") and IsClosed=0 group by TypeOfFund", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                dblPriorAmount = opntbl!BalanceDiff
                'dblPriorDebit = opntbl!DebitBalance
                'dblPriorCredit = opntbl!CreditBalance
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2," & CCur(opntbl!BalanceDiff) & ",0)"
            Else
                dblPriorAmount = 0
                'dblPriorDebit = 0
                'dblPriorCredit = 0
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) values (2,0,0)"
            End If
            opntbl.Close
            Set opntbl = Nothing
                                    
            opntbl.Open "select max(trnno) LastID from tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=" & CCur(dblPriorAmount + dblIncome - dblExpenses) & " where trnno=" & CInt(opntbl!LastID) & ""
            End If
            opntbl.Close
            Set opntbl = Nothing
                    
            opntbl.Open "SELECT SGEID, SUM(CreditBalance) AS " _
                            & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                            & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                            & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                            & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                            & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY SGEID", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opntbl.RecordCount > 0 Then
                opndbaseFMIS.Execute "insert into tblAIS_tempSGE (SGEID,LeftColumn,RightColumn) SELECT SGEID, SUM(CreditBalance) AS " _
                                & "CreditBalance, SUM(DebitBalance) AS DebitBalance FROM     " _
                                & "    dbo.qryReportTrialBalanceProperand20 WHERE     " _
                                & "(CAST(SGEID AS smallint) <> 0) AND (CAST(SGEID " _
                                & "AS smallint) <> 1) AND (CAST(SGEID AS " _
                                & "smallint) <> 2) and MainCode='501' AND MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY SGEID"
            
                opntbl2.Open "select top 1 trnno from tblAIS_tempSGE order by SGEID desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If opntbl.RecordCount > 0 Then
                    opndbaseFMIS.Execute "update tblAIS_tempSGE set RightColumn=(SELECT SUM(LeftColumn) AS LeftCOLUMN " _
                            & "FROM " _
                            & "dbo.qryAIS_StatementofGE WHERE " _
                            & " (IsTransfer <> 0) AND (IsTransfer <> 1) AND " _
                            & "(IsTransfer <> 2)) where trnno=" & CInt(opntbl!Trnno) & ""
                End If
                opntbl2.Close
                Set opntbl2 = Nothing
            End If
            
            'dblCurrentOp = dblIncome - dblExpenses
            
            'opntbl.Open "select SUM(DebitBalance)-sum(creditBalance) AS BalanceDiff from qryReportTrialBalanceProperand20 where MainCode='" & "260" & "'" _
                    & " and MantYer in (" & MantYerRange & ") group by MainCode", opndbaseFMIS, adOpenStatic, adLockOptimistic
            'If opntbl.RecordCount > 0 Then
            '    dblTransferPI = opntbl!BalanceDiff
            'End If
            opntbl.Close
            Set opntbl = Nothing
        End If
        
        opntbl.Open "SELECT SUM(LeftColumn) AS LeftColumn, SUM(RightColumn) AS RightColumn From tblAIS_tempSGE", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntbl.RecordCount > 0 Then
            strSQL = "SELECT     FourthLevelGroup, ThirdLevelGroup, " _
                        & "SecondLevelGroup, FirstLevelGroup, AccountName," & "" & CCur(opntbl!RightColumn) & ", " _
                        & "AccountCode, OrderInBS FROM         " _
                        & "tblREF_AIS_ChartOfAccountsMother WHERE     (AccountCode " _
                        & "= '501')"
            opndbaseFMIS.Execute "insert into tblAIS_TempBS (FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,MainAccountName,BalanceAmount,MainCode,OrderInBS) " & strSQL
        Else
        
        End If
    End If
'    opntbl.Close
    Set opntbl = Nothing
    CrystalReportBS.Database.SetDataSource opndbaseFMIS.Execute("SELECT FourthLevelGroup, ThirdLevelGroup, SecondLevelGroup, FirstLevelGroup, AccountCode, AccountName, IsDebit, Amount From dbo.tblREF_AIS_BalanceSheetFormat WHERE (Amount <> 0)")
   ' Call TransactionLogging("Print Preview", "Balance Sheet", "frmBalanceSheet")
    Call UnsetAnimation(frmBalanceSheet.Animation1)
    PreviewForm.Show vbModal
    
End Sub
