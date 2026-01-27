VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Begin VB.Form frmOfficeSelectSAAO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Offices"
   ClientHeight    =   4950
   ClientLeft      =   2610
   ClientTop       =   2400
   ClientWidth     =   7275
   Icon            =   "frmOfficeSelectSAAO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7275
   Begin VB.ComboBox List1Shadow 
      Height          =   315
      Left            =   3030
      TabIndex        =   5
      Top             =   1755
      Visible         =   0   'False
      Width           =   1215
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   2685
      Top             =   5070
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   135
      TabIndex        =   1
      Top             =   0
      Width           =   5685
      Begin VB.CheckBox byAlobsno 
         BackColor       =   &H00FF8080&
         Caption         =   "By AlobsNo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkClass 
         BackColor       =   &H00FF8080&
         Caption         =   "By Class"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3765
         Width           =   1095
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   3960
         Width           =   1845
      End
      Begin VB.CheckBox chkMonthly 
         BackColor       =   &H00FF8080&
         Caption         =   "Periodically"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   3360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkDetailed 
         BackColor       =   &H00FF8080&
         Caption         =   "Detailed"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   3360
         Width           =   900
      End
      Begin VB.CheckBox chkNegative 
         BackColor       =   &H00FF8080&
         Caption         =   "Negative Balances Only"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3360
         Width           =   2160
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   345
         TabIndex        =   2
         Top             =   630
         Width           =   4890
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   345
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Month"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1905
         TabIndex        =   11
         Top             =   3960
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Please Select an Office"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   345
         TabIndex        =   6
         Top             =   375
         Width           =   4890
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4860
      Left            =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   1290
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   750
         Left            =   240
         Picture         =   "frmOfficeSelectSAAO.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   840
      End
      Begin VB.CommandButton btnOk 
         Caption         =   "&Ok"
         Height          =   750
         Left            =   240
         Picture         =   "frmOfficeSelectSAAO.frx":4814
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmOfficeSelectSAAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub btnClose_Click()
'Unload Me
'End Sub
'
'Private Sub btnOk_Click()
'Dim XDate As Date
'Dim x As Integer
'Dim DRec As New ADODB.Recordset
'Dim DRec2 As New ADODB.Recordset
'Dim DRec3 As New ADODB.Recordset
'
'If List1.Text <> "" Then
'
'   ' ViewerCaption = "Status of Appropriations, Allotments and Obligation"
'
'
'    If Year(ServerDate) > GetTransYear Then
'        XDate = "January 1, " & Year(ServerDate)
'        If List1.Text = "All" Then
'            If chkDetailed.Value = 0 Then
'                If chkMonthly.Value = 0 Then
'                    CrptSAAO_All.txtAsOf.SetText "For the period of January 1, " & GetTransYear & " to " & Format(XDate - 1, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptMonthlySAAO_All.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & GetTransYear
'                    Else
'                        CrptMonthlySAAO_All.txtAsOf.SetText "For the month of January " & GetTransYear
'                    End If
'                End If
'            Else
'                If chkMonthly.Value = 0 Then
'                    CrptDetailedSAAO_All.txtAsOf.SetText "For the period of January 1, " & GetTransYear & " to " & Format(XDate - 1, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptDetailedMonthlySAAO_All.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & GetTransYear
'                    Else
'                        CrptDetailedMonthlySAAO_All.txtAsOf.SetText "For the month of January " & GetTransYear
'                    End If
'                End If
'            End If
'        Else
'            If chkDetailed.Value = 0 Then
'                If chkMonthly.Value = 0 Then
'                    CrptSAAO.txtAsOf.SetText "For the period of January 1, " & GetTransYear & " to " & Format(XDate - 1, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptMonthlySAAO.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & GetTransYear
'                    Else
'                        CrptMonthlySAAO.txtAsOf.SetText "For the month of January " & GetTransYear
'                    End If
'                End If
'            Else
'                If chkMonthly.Value = 0 Then
'                    CrptDetailedSAAO.txtAsOf.SetText "For the period of January 1, " & GetTransYear & " to " & Format(XDate - 1, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptDetailedMonthlySAAO.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & GetTransYear
'                    Else
'                        CrptDetailedMonthlySAAO.txtAsOf.SetText "For the month of January " & GetTransYear
'                    End If
'                End If
'            End If
'        End If
'    Else
'        If List1.Text = "All" Then
'            If chkDetailed.Value = 0 Then
'                If chkMonthly.Value = 0 Then
'                    CrptSAAO_All.txtAsOf.SetText "For the period of January 1, " & Year(ServerDate) & " to " & Format(ServerDate, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptMonthlySAAO_All.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & Year(ServerDate)
'                    Else
'                        CrptMonthlySAAO_All.txtAsOf.SetText "For the month of January " & Year(ServerDate)
'                    End If
'                End If
'            Else
'                If chkMonthly.Value = 0 Then
'                    CrptDetailedSAAO_All.txtAsOf.SetText "For the period of January 1, " & Year(ServerDate) & " to " & Format(ServerDate, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptDetailedMonthlySAAO_All.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & Year(ServerDate)
'                    Else
'                        CrptDetailedMonthlySAAO_All.txtAsOf.SetText "For the month of January " & Year(ServerDate)
'                    End If
'                End If
'            End If
'        Else
'            If chkDetailed.Value = 0 Then
'                If chkMonthly.Value = 0 Then
'                    CrptSAAO.txtAsOf.SetText "For the period of January 1, " & Year(ServerDate) & " to " & Format(ServerDate, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptMonthlySAAO.txtAsOf.SetText "For the period of January To " & cmbMonth.Text & " " & Year(ServerDate)
'                    Else
'                        CrptMonthlySAAO.txtAsOf.SetText "For the month of January " & Year(ServerDate)
'                    End If
'                End If
'            Else
'                If chkMonthly.Value = 0 Then
'                    CrptDetailedSAAO.txtAsOf.SetText "For the period of January 1, " & Year(ServerDate) & " to " & Format(ServerDate, "mmmm d, yyyy")
'                Else
'                    If cmbMonth.Text <> "January" Then
'                        CrptDetailedMonthlySAAO.txtAsOf.SetText "For the period of January to " & cmbMonth.Text & " " & Year(ServerDate)
'                    Else
'                        CrptDetailedMonthlySAAO.txtAsOf.SetText "For the month of January " & Year(ServerDate)
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    If List1.Text = "All" Then
'        If chkDetailed.Value = 0 Then
'            If chkMonthly.Value = 0 Then
'                ReportName = "SAAO_All"
'                CrptSAAO_All.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                If chkNegative.Value = 1 Then
'                    CrptSAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_SAAO_Report Where YearOf=" & GetTransYear & " And (ReleaseAccount - Obligation)<0")
'                Else
'                    CrptSAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_SAAO_Report Where YearOf=" & GetTransYear & "")
'                End If
'            Else
'                ReportName = "MonthlySAAO_All"
'                CrptMonthlySAAO_All.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                Call PeriodicSAAO(List1Shadow.List(List1.ListIndex), cmbMonth.ListIndex + 1)
'
'                '=====================UPDATE HERE===============================
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode='" & DRec2.RecordCount + 1 & "' where FMISOfficeCode='43'"
'                DRec2.Close
'                '===============================================================
'
'                'to be use if FundCode=101 of Contractuall/Statutory Obligations
'                'opndbaseFMIS.Execute "Update [tblBMS_SAAOPeriod_Dummy] set Fundcode=101,FundName='General Fund Proper' where FMISOfficeCode =43 and FMISProgramCode =55 and FMISAccountCode =2861 and yearof=" & GetTransYear & " ", opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'                If chkNegative.Value = 1 Then
'                    CrptMonthlySAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where (Allotment - Obligation)<0 And UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                Else
'                    CrptMonthlySAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "' Order by FundCode,FMISOfficeCode, MotherOfficeCode,ProgOrder, ExpenseClassCode, AcctOrder")
'                End If
'
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode=43 where FMISOfficeCode='" & DRec2.RecordCount + 1 & "'"
'                DRec2.Close
'
'                'to be use if FundCode=101 of Contractuall/Statutory Obligations
'                'opndbaseFMIS.Execute "Update [tblBMS_SAAOPeriod_Dummy] set Fundcode=118,FundName='20% Development Fund' where FMISOfficeCode =43 and FMISProgramCode =55 and FMISAccountCode =2861 and yearof=" & GetTransYear & " ", opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'            End If
'        Else
'            If chkMonthly.Value = 0 Then
'                ReportName = "DetailedSAAO_All"
'                CrptDetailedSAAO_All.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                If chkNegative.Value = 1 Then
'                    CrptDetailedSAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_DetailedSAAO_Report Where YearOf=" & GetTransYear & " And (ReleaseAccount - Obligation)<0")
'                Else
'                    CrptDetailedSAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_DetailedSAAO_Report Where YearOf=" & GetTransYear & "")
'                End If
'            Else
'                ReportName = "DetailedMonthlySAAO_All"
'                CrptDetailedMonthlySAAO_All.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'                Call PeriodicSAAO(List1Shadow.List(List1.ListIndex), cmbMonth.ListIndex + 1)
'
'                '=====================UPDATE HERE===============================
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode='" & DRec2.RecordCount + 1 & "' where FMISOfficeCode='43'"
'                DRec2.Close
'                '===============================================================
'
'                If chkNegative.Value = 1 Then
'                    CrptDetailedMonthlySAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where (Allotment - Obligation)<0 and UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                Else
'                    CrptDetailedMonthlySAAO_All.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                End If
'
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode=43 where FMISOfficeCode='" & DRec2.RecordCount + 1 & "'"
'                DRec2.Close
'
'            End If
'        End If
'        'Call LogActivity(Me.Caption, "From vwBMS_SAAO_Report", UserID, "Select * From vwBMS_SAAO_Report Where YearOf=" & GetTransYear & "")
'        'CrptSAAO_All.txtOfficeName.SetText List1.Text
'        'CrptSAAO_All.txtFundName.SetText "(" & GetFundType(GetFundCode(List1Shadow.List(List1.ListIndex))) & ")"
'    Else
'        If chkDetailed.Value = 0 Then
'            If chkMonthly.Value = 0 Then
'                ReportName = "SAAO"
'                CrptSAAO.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                If chkNegative.Value = 1 Then
'                    CrptSAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & " And (ReleaseAccount - Obligation)<0")
'                Else
'                    CrptSAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                End If
'
'                'Call LogActivity(Me.Caption, "From vwBMS_SAAO_Report", UserID, "Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                CrptSAAO.txtOfficeName.SetText List1.Text
'            Else
'                ReportName = "MonthlySAAO"
'                CrptMonthlySAAO.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                Call PeriodicSAAO(List1Shadow.List(List1.ListIndex), cmbMonth.ListIndex + 1)
'
'                '=====================UPDATE HERE===============================
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode='" & DRec2.RecordCount + 1 & "' where FMISOfficeCode='43'"
'                DRec2.Close
'                '===============================================================
'
'                If chkNegative.Value = 1 Then
'                    CrptMonthlySAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where (Allotment - Obligation)<0 And UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                Else
'                    CrptMonthlySAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                End If
'
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode=43 where FMISOfficeCode='" & DRec2.RecordCount + 1 & "'"
'                DRec2.Close
'
'                'Call LogActivity(Me.Caption, "From vwBMS_SAAO_Report", UserID, "Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                CrptMonthlySAAO.txtOfficeName.SetText List1.Text
'            End If
'        Else
'            If chkMonthly.Value = 0 Then
'                ReportName = "DetailedSAAO"
'                CrptDetailedSAAO.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                If chkNegative.Value = 1 Then
'                    CrptDetailedSAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_DetailedSAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & " And (ReleaseAccount - Obligation)<0")
'                Else
'                    CrptDetailedSAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From vwBMS_DetailedSAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                End If
'
'                'Call LogActivity(Me.Caption, "From vwBMS_SAAO_Report", UserID, "Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                CrptDetailedSAAO.txtOfficeName.SetText List1.Text
'            Else
'                ReportName = "DetailedMonthlySAAO"
'                CrptDetailedMonthlySAAO.txtUser.SetText "Printed by : " & StrConv(UserName, vbProperCase)
'
'                Call PeriodicSAAO(List1Shadow.List(List1.ListIndex), cmbMonth.ListIndex + 1)
'
'                '=====================UPDATE HERE===============================
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode='" & DRec2.RecordCount + 1 & "' where FMISOfficeCode='43'"
'                DRec2.Close
'                '===============================================================
'
'                If chkNegative.Value = 1 Then
'                    CrptDetailedMonthlySAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where (Allotment - Obligation)<0 and UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                Else
'                    CrptDetailedMonthlySAAO.Database.SetDataSource opndbaseFMIS.Execute("Select * From tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "' Order by FundCode,FMISOfficeCode,MotherOfficeCode, ProgOrder, ExpenseClassCode, AcctOrder")
'                End If
'
'                DRec2.Open ("Select * from tblREF_AIS_Offices"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy set FMISOfficeCode=43 where FMISOfficeCode='" & DRec2.RecordCount + 1 & "'"
'                DRec2.Close
'
'                'Call LogActivity(Me.Caption, "From vwBMS_SAAO_Report", UserID, "Select * From vwBMS_SAAO_Report Where FMISOfficeCode=" & List1Shadow.List(List1.ListIndex) & " And YearOf=" & GetTransYear & "")
'                CrptDetailedMonthlySAAO.txtOfficeName.SetText List1.Text
'            End If
'        End If
'        'CrptSAAO.txtFundName.SetText "(" & GetFundType(GetFundCode(List1Shadow.List(List1.ListIndex))) & ")"
'        'CrptSAAO.txtForTheYear.SetText "For the year " & GetTransYear
'        'CrptSAAO.txtPreparedBy.SetText StrConv(UserName, vbUpperCase)
'        'CrptSAAO.txtPreparedByDesignation.SetText UserDesignation
'    End If
'
'    'strPreparedby strPreparedbyPos strNotedby strNotedbyPos
'    CrptMonthlySAAO.txtPreparedBy.SetText strPreparedby
'    CrptMonthlySAAO.txtPreparedByDesignation.SetText strPreparedbyPos
'    CrptMonthlySAAO.Text18.SetText strNotedby
'    CrptMonthlySAAO.Text19.SetText strNotedbyPos
'
'
'    frmReportViewer.Caption = "Status of Appropriations, Allotments and Obligation"
'    frmReportViewer.Show 1
'Else
'    MsgBox "Please select an office.", vbExclamation + vbOKOnly, "BMS Security"
'End If
'
'End Sub
'
'Private Sub PeriodicSAAO(ByVal FMISOfficeCode As Integer, ByVal EndMonth As Integer)
'Dim DRec As New ADODB.Recordset
'Dim x As Integer
'
'opndbaseFMIS.Execute "Delete tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "'"
'
'If FMISOfficeCode = 0 Then
'    opndbaseFMIS.Execute "Insert Into tblBMS_SAAOPeriod_Dummy(FundCode,FundName,MotherOfficeCode,MotherOfficeName,FMISOfficeCode,OfficeName,ProgOrder,FMISProgramCode,ProgramName,ExpenseClassCode,ExpenseClassName,AcctOrder,FMISAccountCode,AccountName,Appropriation,YearOf,UserID) select *,'" & UserID & "' as UserID from vwBMS_SAAODummy_Prep where yearof=" & GetTransYear & ""
'Else
'    opndbaseFMIS.Execute "Insert Into tblBMS_SAAOPeriod_Dummy(FundCode,FundName,MotherOfficeCode,MotherOfficeName,FMISOfficeCode,OfficeName,ProgOrder,FMISProgramCode,ProgramName,ExpenseClassCode,ExpenseClassName,AcctOrder,FMISAccountCode,AccountName,Appropriation,YearOf,UserID) select *,'" & UserID & "' as UserID from vwBMS_SAAODummy_Prep where yearof=" & GetTransYear & " and FMISOfficeCode=" & FMISOfficeCode & ""
'End If
'
'ProgressBar1.Min = 0
'ProgressBar1.Value = 0
'
'DRec.Open ("Select * From tblBMS_SAAOPeriod_Dummy Where UserID='" & UserID & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'If DRec.RecordCount <> 0 Then
'
'    ProgressBar1.Visible = True
'    ProgressBar1.Max = DRec.RecordCount
'
'    For x = 1 To DRec.RecordCount
'        ProgressBar1.Value = x
'
'            opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy Set Reserve=" & GetReserveAmount(DRec!FMISProgramCode, DRec!FmisAccountcode) & ", Reallignment=" & GetReAllignedAmountTo(DRec!FMISProgramCode, DRec!FmisAccountcode) - GetReAlignDeduction(DRec!FMISProgramCode, DRec!FmisAccountcode) & ", Supplemental=" & GetTotalPeriodicSupplementalPerAccount(DRec!FMISProgramCode, DRec!FmisAccountcode, EndMonth) & ",ReversionPS=0,ReversionMOOE=0,ReversionCO=0, Allotment=" & GetTotalPeriodicRelease_Account(DRec!FMISProgramCode, DRec!FmisAccountcode, DRec!ExpenseClassCode, EndMonth) & ", Obligation=" & GetTotalPeriodicControl_Account(DRec!FMISProgramCode, DRec!FmisAccountcode, EndMonth) & ", EndMonth=" & EndMonth & " Where FMISProgramCode=" & DRec!FMISProgramCode & " And FMISAccountCode=" & DRec!FmisAccountcode & " And UserID='" & UserID & "'"
'            opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy Set ReversionPS=" & GetReversedTo(DRec!FMISProgramCode, DRec!FmisAccountcode, 1) - GetReversedFrom(DRec!FMISProgramCode, DRec!FmisAccountcode, 1) & " Where FMISProgramCode=" & DRec!FMISProgramCode & " And FMISAccountCode=" & DRec!FmisAccountcode & " And UserID='" & UserID & "'"
'            opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy Set ReversionMOOE=" & GetReversedTo(DRec!FMISProgramCode, DRec!FmisAccountcode, 2) - GetReversedFrom(DRec!FMISProgramCode, DRec!FmisAccountcode, 2) & " Where FMISProgramCode=" & DRec!FMISProgramCode & " And FMISAccountCode=" & DRec!FmisAccountcode & " And UserID='" & UserID & "'"
'            opndbaseFMIS.Execute "Update tblBMS_SAAOPeriod_Dummy Set ReversionCO=" & GetReversedTo(DRec!FMISProgramCode, DRec!FmisAccountcode, 3) - GetReversedFrom(DRec!FMISProgramCode, DRec!FmisAccountcode, 3) & " Where FMISProgramCode=" & DRec!FMISProgramCode & " And FMISAccountCode=" & DRec!FmisAccountcode & " And UserID='" & UserID & "'"
'
'        DRec.MoveNext
'    Next x
'End If
'DRec.Close
'Set DRec = Nothing
'
'ProgressBar1.Visible = False
'
'End Sub
'
'Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete Then
'    cmbMonth.Text = ""
'End If
'End Sub
'
'Private Sub cmbMonth_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'Else
'    KeyAscii = AutoFind(cmbMonth, KeyAscii, True)
'End If
'End Sub
'
'Private Sub Form_Load()
'Dim RecOffice As New ADODB.Recordset
'Dim x As Integer
'
'
'''Skinner1.ChangeSkinButton = False
'
''WindowsXPC1.InitSubClassing
'
'Me.Top = (Screen.Height / 2) - (Me.Height / 2)
'Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'
'
'
'If ActiveUserID <> "1237" And ActiveUserID <> "1735" Then
'    List1.AddItem "All"
'    List1Shadow.AddItem "0"
'    RecOffice.Open ("Select * From tblREF_AIS_Offices Order By OfficeName"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'Else
'    RecOffice.Open ("Select FMISOfficeID,upper(OfficeName) as OfficeName From tblREF_AIS_Offices where FMISOfficeID=43 or FMISOfficeID=1 Order By OfficeName"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'End If
'
'If RecOffice.RecordCount <> 0 Then
'    For x = 1 To RecOffice.RecordCount
'        List1.AddItem RecOffice!Officename
'        List1Shadow.AddItem RecOffice!fmisofficeid
'        RecOffice.MoveNext
'    Next x
'End If
'RecOffice.Close
'Set RecOffice = Nothing
'
'For x = 1 To 12
'    cmbMonth.AddItem MonthName(x)
'Next x
'
'chkMonthly.Enabled = False
'chkMonthly.Value = 1
'
''Call LogActivity(Me.Caption, "", ActiveUserID, "Open")
'
'End Sub
'
