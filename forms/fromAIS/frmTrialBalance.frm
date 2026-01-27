VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmTrialBalance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   4680
   ClientTop       =   2670
   ClientWidth     =   5115
   Icon            =   "frmTrialBalance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5115
   Begin VB.Frame Frame3 
      Height          =   35
      Left            =   -105
      TabIndex        =   17
      Top             =   840
      Width           =   7335
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
      TabIndex        =   12
      Top             =   1005
      Width           =   2490
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
         TabIndex        =   2
         Top             =   1770
         Width           =   2385
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
         TabIndex        =   14
         Top             =   1200
         Width           =   210
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
         TabIndex        =   1
         Top             =   1155
         Width           =   2010
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
         TabIndex        =   13
         Top             =   420
         Width           =   1530
      End
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
         TabIndex        =   0
         Top             =   675
         Width           =   2370
      End
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
      Left            =   2535
      TabIndex        =   7
      Top             =   1005
      Width           =   2550
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
         TabIndex        =   15
         Top             =   1605
         Width           =   2145
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
         TabIndex        =   9
         Top             =   135
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   150
         Left            =   105
         TabIndex        =   8
         Top             =   1065
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   3
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
         Format          =   56885251
         CurrentDate     =   38838
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
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
         Format          =   56885251
         CurrentDate     =   38868
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
         TabIndex        =   11
         Top             =   435
         Width           =   855
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
         TabIndex        =   10
         Top             =   795
         Width           =   720
      End
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
      TabIndex        =   5
      Top             =   3420
      Width           =   960
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
      TabIndex        =   6
      Top             =   3420
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sets criteria for Trial Balance prior to previewing."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   195
      TabIndex        =   19
      Top             =   480
      Width           =   3510
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
      Left            =   195
      TabIndex        =   18
      Top             =   210
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -45
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmTrialBalance"
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
        Err.Source = Err.Source & "." & TypeName(Me) & ".chkConsolidated_Click"
        Set .Error = Err
     
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


            opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where upper(typeoffund)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", fmisDB, adOpenStatic, adLockOptimistic
        Else



            opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where responsibilitycenter='" & Trim$(cboEco.Text) & "' and  upper(typeoffund)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", fmisDB, adOpenStatic, adLockOptimistic
        End If
    ElseIf chkConsolidated = 1 And chkProperAnd20.Value = 0 Then


        opntbl.Open "SELECT DISTINCT dbo.tblAIS_JEV.JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) FROM dbo.tblAIS_JEV INNER JOIN dbo.tblRefBMS_Funds ON dbo.tblAIS_JEV.TypeOfFund = dbo.tblRefBMS_Funds.FundName INNER JOIN dbo.tblREF_AIS_Fundtype ON dbo.tblRefBMS_Funds.MotherFund = dbo.tblREF_AIS_Fundtype.fundcode Where (dbo.tblAIS_JEV.actioncode = 1) And (dbo.tblAIS_JEV.IsApproved = 1) and upper(dbo.tblREF_AIS_Fundtype.MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' and (IsClosed = 0) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc,jevnumber", fmisDB, adOpenStatic, adLockOptimistic
    ElseIf chkConsolidated = 0 And chkProperAnd20.Value = 1 Then
        opntbl.Open "SELECT distinct JEVNumber,CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) From dbo.tblAIS_JEV Where upper(typeofFund) in ('" & "GENERAL FUND PROPER" & "','" & "20% DEVELOPMENT FUND" & "') and (IsApproved = 1) And (IsClosed = 0) And (actioncode = 1) And (JEVNumber Is Not Null) order by CAST(SUBSTRING(dbo.tblAIS_JEV.JEVNumber, 8, 2) AS int) asc, jevnumber", fmisDB, adOpenStatic, adLockOptimistic
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
        Err.Source = Err.Source & "." & TypeName(Me) & ".DisplayALLJEV"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub

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

    On Error GoTo errHandler
    strReportName = "TB"
    
    strSQLStatement = ""
    Call SetAnimation(frmTrialBalance.Animation1)
    If chkClosing.Value = 0 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQLStatement = "SELECT TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceFinal " _
                                    & "WHERE (CAST(SUBSTRING(MantYer, CHARINDEX('-', MantYer) + 1, 2) AS smallint) < " & CInt(Format(DTPicker2.Value, "yy")) & " ) and IsClosed=1 GROUP BY MainCode, MainAccountName, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) union all SELECT TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), " _
                                    & "SUM(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceFinal WHERE MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY MainCode, MainAccountName, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode"
                    'CrystalReportTB.Database.SetDataSource fmisDB.Execute("SELECT TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceFinal WHERE MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY MainCode, MainAccountName, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
                Else
                    If Len(Trim(cboEco.Text)) > 0 Then
                        strSQLStatement = "SELECT TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance), " _
                                    & "typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A  FROM qryReportTrialBalanceWithRC WHERE (CAST(SUBSTRING(MantYer, CHARINDEX('-', MantYer) + 1, 2) AS smallint) < " & CInt(Format(DTPicker2.Value, "yy")) & " ) AND IsClosed=1 AND " _
                                    & "ResponsibilityCenter='" & Trim$(cboEco.Text) & "' GROUP " _
                                    & "BY MainCode,MainAccountname, typeoffund HAVING " _
                                    & "upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "'" _
                                    & " AND (SUM(DebitBalance) - SUM(CreditBalance) <> " _
                                    & "0) union ALL SELECT TOP 100 " _
                                    & "PERCENT MainAccountname,MainCode, " _
                                    & "sum(DebitBalance), sum(CreditBalance), typeoffund, SUM(DebitBalance)- " _
                                    & "SUM(CreditBalance) AS A FROM qryReportTrialBalanceWithRC " _
                                    & "WHERE MantYer IN (" & MantYerRange & ") AND IsClosed=0 " _
                                    & "AND ResponsibilityCenter='" & Trim$(cboEco.Text) & "'" _
                                    & " GROUP BY MainCode,MainAccountname, typeoffund " _
                                    & "HAVING upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) " _
                                    & "- SUM(CreditBalance) <> 0) ORDER BY MainCode"
                        'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceWithRC where MantYer in (" & MantYerRange & ") and IsClosed=0 and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' group by MainCode,MainAccountname, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                strSQLStatement = "SELECT TOP 100 PERCENT MainAccountname,MainCode, " _
                            & " sum(DebitBalance), sum(CreditBalance)," _
                            & "MotherFundType, SUM(DebitBalance)- SUM(CreditBalance) AS " _
                            & "A FROM qryReportTrialBalanceConsolidated WHERE " _
                            & "(CAST(SUBSTRING(MantYer, CHARINDEX('-', MantYer) + 1, 2) AS smallint) < " & CInt(Format(DTPicker2.Value, "yy")) & " ) AND IsClosed=1 " _
                            & "GROUP BY MainCode,MotherFundType,MainAccountname " _
                            & "HAVING upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) " _
                            & "- SUM(CreditBalance) <> 0) union ALL SELECT TOP 100 " _
                            & "PERCENT MainAccountname,MainCode,  " _
                            & "sum(DebitBalance), sum(CreditBalance),MotherFundType, " _
                            & "SUM(DebitBalance)- SUM(CreditBalance) AS A FROM " _
                            & "qryReportTrialBalanceConsolidated WHERE MantYer IN (" & MantYerRange & ") AND IsClosed=0 GROUP BY MainCode," _
                            & "MotherFundType,MainAccountname HAVING upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND " _
                            & "(SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER " _
                            & "BY MainCode"
                'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance),MotherFundType, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceConsolidated where MantYer in (" & MantYerRange & ") and IsClosed=0 group by MainCode,MotherFundType,MainAccountname having upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            strSQLStatement = "SELECT TOP 100 PERCENT MainAccountName, MainCode," _
                            & " SUM(DebitBalance), SUM(CreditBalance), " _
                            & "SUM(DebitBalance)- SUM(CreditBalance) AS A FROM " _
                            & "qryReportTrialBalanceProperand20 WHERE (CAST(SUBSTRING(MantYer, CHARINDEX('-', MantYer) + 1, 2) AS smallint) < " & CInt(Format(DTPicker2.Value, "yy")) & " ) AND IsClosed=1 GROUP BY " _
                            & "MainCode, MainAccountName HAVING (SUM(DebitBalance) " _
                            & "- SUM(CreditBalance) <> 0) union ALL SELECT TOP 100 " _
                            & "PERCENT MainAccountName, MainCode, " _
                            & "SUM(DebitBalance), SUM(CreditBalance), " _
                            & "SUM(DebitBalance)- SUM(CreditBalance) AS A FROM qryReportTrialBalanceProperand20 " _
                            & "WHERE MantYer IN (" & MantYerRange & ") AND " _
                            & "IsClosed=0 GROUP BY MainCode, MainAccountName " _
                            & "HAVING (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER " _
                            & "BY MainCode"
            'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceProperand20 where MantYer in (" & MantYerRange & ") and IsClosed=0 GROUP BY MainCode, MainAccountName having (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
        End If
    ElseIf chkClosing.Value = 1 Then
        If chkConsolidated.Value = 0 And chkProperAnd20.Value = 0 Then    'if not consolidated
            If optDateRange.Value = True Then    'if date range
                If chkEco.Value = 0 Then
                    strSQLStatement = "SELECT TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceFinal WHERE MantYer in (" & MantYerRange & ") GROUP BY MainCode, MainAccountName, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode"
                    'CrystalReportTB.Database.SetDataSource fmisDB.Execute("SELECT TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceFinal WHERE MantYer in (" & MantYerRange & ") GROUP BY MainCode, MainAccountName, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
                Else
                    If Len(Trim$(cboEco.Text)) > 0 Then
                        strSQLStatement = "select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceWithRC where MantYer in (" & MantYerRange & ") and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' group by MainCode,MainAccountname, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode"
                        'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance), typeoffund, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceWithRC where MantYer in (" & MantYerRange & ") and ResponsibilityCenter='" & Trim$(cboEco.Text) & "' group by MainCode,MainAccountname, typeoffund having upper(TypeOfFund)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
                    Else
                        MsgBox "Please select type of fund!", vbCritical, "System Information"
                        Exit Sub
                    End If
                End If
            End If
        ElseIf chkConsolidated.Value = 1 And chkProperAnd20.Value = 0 Then    'if consolidated
            If optDateRange.Value = True Then    'if date range
                strSQLStatement = "select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance),MotherFundType, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceConsolidated where MantYer in (" & MantYerRange & ") group by MainCode,MotherFundType,MainAccountname having upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode"
                'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountname,MainCode,  sum(DebitBalance), sum(CreditBalance),MotherFundType, SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceConsolidated where MantYer in (" & MantYerRange & ") group by MainCode,MotherFundType,MainAccountname having upper(MotherFundType)='" & UCase$(Trim$(cboFundType.Text)) & "' AND (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
            End If
        ElseIf chkConsolidated.Value = 0 And chkProperAnd20.Value = 1 Then
            strSQLStatement = "select TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceProperand20 where MantYer in (" & MantYerRange & ") GROUP BY MainCode, MainAccountName having (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode"
            'CrystalReportTB.Database.SetDataSource fmisDB.Execute("select TOP 100 PERCENT MainAccountName, MainCode, SUM(DebitBalance), SUM(CreditBalance), SUM(DebitBalance)- SUM(CreditBalance) AS A from qryReportTrialBalanceProperand20 where MantYer in (" & MantYerRange & ") GROUP BY MainCode, MainAccountName having (SUM(DebitBalance) - SUM(CreditBalance) <> 0) ORDER BY MainCode")
        End If
    End If
    If Len(Trim(strSQLStatement)) > 0 Then
        CrystalReportTB.Database.SetDataSource fmisDB.Execute(strSQLStatement)
        Call TransactionLogging("Print Preview", "Trial Balance", "frmTrialBalance")
        Call UnsetAnimation(frmTrialBalance.Animation1)
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
    MyDLL.CenterMe Me
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
    Set frmTrialBalance = Nothing
    Exit Sub
 
errHandler:
 
    With frmVBError
        Err.Source = Err.Source & "." & TypeName(Me) & ".Form_Unload"
        Set .Error = Err
     
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
        Err.Source = Err.Source & "." & TypeName(Me) & ".optDateRange_Click"
        Set .Error = Err
     
        .Show vbModal
        Set frmVBError = Nothing
    End With
 
End Sub
