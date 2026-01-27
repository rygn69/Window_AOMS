VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frm_ScheduleManagement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   10005
   ClientLeft      =   2280
   ClientTop       =   1170
   ClientWidth     =   15705
   Icon            =   "frm_ScheduleManagement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   15705
   Begin VB.Frame Frame1 
      Caption         =   "Bank Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   7215
      Begin VB.OptionButton opt_Statement 
         Caption         =   "Statement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "4"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Opt_Withdrawals 
         Caption         =   "Withdrawals"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "2"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2280
         Top             =   2715
      End
      Begin VB.OptionButton opt_Deposits 
         Caption         =   "Deposits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "1"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15435
      Begin VB.Frame Frame2 
         Caption         =   "Special Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3690
         Begin VB.ComboBox cmb_FundType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   195
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   3300
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   14160
         TabIndex        =   3
         Top             =   360
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1185
         Left            =   120
         Top             =   240
         Width           =   15225
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   4770
      Top             =   11265
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   8910
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   15716
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "itb32x32"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "slash"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Save"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Match"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Delete"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "s"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel"
               ImageIndex      =   7
            EndProperty
         EndProperty
         Begin MSComCtl2.Animation Animation1 
            Height          =   450
            Left            =   11400
            TabIndex        =   12
            Top             =   120
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   794
            _Version        =   393216
            FullWidth       =   32
            FullHeight      =   30
         End
         Begin MSComctlLib.ImageList itb32x32 
            Left            =   120
            Top             =   5520
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   17
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":0E42
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":27D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":4166
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":5AF8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":748A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":8E1C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":A7AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":C140
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":DAD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":F466
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":10142
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":10A22
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":116FE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":123DA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":130B6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":13D92
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_ScheduleManagement.frx":14A6E
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13755
      TabIndex        =   2
      Top             =   9975
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   9300
      Width           =   480
   End
End
Attribute VB_Name = "frm_ScheduleManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim tmpAccName As String
'Dim FMISNo As String
'
'Private Sub cmb_BankName_Change()
'Call Loadcmb(cmb_Accountno, "EXECUTE [fmis].[dbo].[MPproc_Get_AccountNoByBankname] @fundtype = '" & cmb_FundType.Text & "', @bankid = " & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "")
'End Sub
'
'Private Sub cmb_BankName_Click()
'Call Loadcmb(cmb_Accountno, "EXECUTE [fmis].[dbo].[MPproc_Get_AccountNoByBankname] @fundtype = '" & cmb_FundType.Text & "', @bankid = " & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "")
'End Sub
'
'Private Sub Command1_Click()
'LoadOPTBook
'LoadOPTBank
'End Sub
'
'Private Sub Form_Load()
''WindowsXPC1.InitSubClassing
'
'Call LoadFundType(cmb_FundType)
'Call Loadcmb(cmb_BankName, "SELECT [trnno] as field1,[BankName] as field2 FROM [fmis].[dbo].[tblCMS_CDBankLibrary]")
'Call Loadcmb(cmbBankClass, "EXECUTE [fmis].[dbo].[MPproc_BankSubClass] @what = 'bank'")
'Call Loadcmb(cmbBookClass, "EXECUTE [fmis].[dbo].[MPproc_BankSubClass] @what = 'Book'")
''Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
'End Sub
'
'
'
'Private Sub lst_Bank_Click()
'Dim x As Long
'Dim Ok As Boolean
'Ok = False
'If lst_Bank.ListItems.Count = 0 Then
'    Exit Sub
'End If
'With lst_Journal
'
'    For x = 1 To .ListItems.Count
'            .ListItems(x).Bold = False
'            .ListItems(x).ListSubItems(1).Bold = False
'            .ListItems(x).ListSubItems(2).Bold = False
'            .ListItems(x).ListSubItems(3).Bold = False
'
'            .ListItems(x).ForeColor = &H0&
'            .ListItems(x).ListSubItems(1).ForeColor = &H0&
'            .ListItems(x).ListSubItems(2).ForeColor = &H0&
'            .ListItems(x).ListSubItems(3).ForeColor = &H0&
'        If lst_Bank.SelectedItem.SubItems(2) = .ListItems(x).SubItems(2) Then
'            lst_Journal.HideSelection = True
'            .ListItems(x).Selected = True
'
'            .ListItems(x).Bold = True
'            .ListItems(x).ListSubItems(1).Bold = True
'            .ListItems(x).ListSubItems(2).Bold = True
'            .ListItems(x).ListSubItems(3).Bold = True
'            .ListItems(x).Top = 1
'            .ListItems(x).ForeColor = &HFF&
'            .ListItems(x).ListSubItems(1).ForeColor = &HFF&
'            .ListItems(x).ListSubItems(2).ForeColor = &HFF&
'            .ListItems(x).ListSubItems(3).ForeColor = &HFF&
'            Ok = True
'        End If
'    Next x
'.Refresh
'If Ok = False Then
'    MsgBox "No Match Found...!", vbInformation, "System Message"
'    lst_Journal.HideSelection = True
'End If
'End With
'If Trim(lst_Bank.SelectedItem.Text) <> "" Then
'    Call toolbarSTat("Open")
'End If
'End Sub
'
'
'
'
'
'Private Sub lst_Journal_KeyUp(KeyCode As Integer, Shift As Integer)
'KeyCode = vbUpArrow
'End Sub
'
'Private Sub opt_Check_Click()
'LoadOPTBook
'
'End Sub
'
'Private Sub Opt_CR_Click()
'LoadOPTBook
'End Sub
'
'Private Sub opt_Deposits_Click()
'LoadOPTBank
'End Sub
'Private Sub LoadOPTBank()
'If Opt_Withdrawals.Value = True Then
'    opt_Check.Value = True
'    Call LoadReconData(2, lst_Bank, 2)
'ElseIf opt_Deposits.Value = True Then
'    opt_CR.Value = True
'    Call LoadReconData(1, lst_Bank, 2)
'ElseIf opt_Statement.Value = True Then
'    Opt_GJ.Value = True
'    Call LoadReconData(4, lst_Bank, 2)
'End If
'End Sub
'Private Sub LoadOPTBook()
'If opt_Check.Value = True Then
'    Opt_Withdrawals.Value = True
'    Call LoadReconData(2, lst_Journal, 1)
'ElseIf opt_CR.Value = True Then
'    opt_Deposits.Value = True
'    Call LoadReconData(1, lst_Journal, 1)
'ElseIf Opt_GJ.Value = True Then
'    opt_Statement.Value = True
'    Call LoadReconData(4, lst_Journal, 1)
'End If
'End Sub
'Private Sub LoadReconData(ByVal TYP As Integer, ByVal lstview As ListView, ByVal What As String)
'Dim rec As New ADODB.Recordset
'On Error GoTo bad
'Dim x As Long
'Dim y
'lstview.ListItems.Clear
'Set rec = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_LoadJournaForBankRecon] @fundtype = '" & cmb_FundType.Text & "',@BankID = '" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "',@Accountno = '" & cmb_Accountno.Text & "',@month = '" & DTPicker3.Month & "',@year = '" & DTPicker3.Year & "',@transtype = " & TYP & ",@what = " & What & "")
'If rec.RecordCount > 0 Then
'    For x = 1 To rec.RecordCount
'        'DoEvents
'        Set y = lstview.ListItems.Add(, , Format(rec!Date_, "mm/dd/yyyy"))
'        y.SubItems(1) = Trim(rec!Particular)
'        y.SubItems(2) = Trim(rec!checkno)
'        y.SubItems(3) = Format(rec!AMOUNT, "#,##0.00")
'        y.SubItems(4) = rec!id
'        rec.MoveNext
'    Next x
'End If
'rec.Close
'Exit Sub
'bad:
'If What = 1 Then
'    If err.Number = 3704 Then
'    MsgBox "Please Iditify the Transaction type...", vbInformation, "System Message"
'    End If
'End If
'End Sub
''Private Sub LoadMatch(ByVal TYP As Integer, ByVal lstview As ListView, ByVal What As String)
''Dim rec As New ADODB.Recordset
''On Error GoTo bad
''Dim x As Long
''Dim y
''lstview.ListItems.Clear
''Set rec = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_LoadJournaForBankRecon] @fundtype = '" & cmb_FundType.Text & "',@BankID = '" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "',@Accountno = '" & cmb_Accountno.Text & "',@month = '" & DTPicker3.Month & "',@year = '" & DTPicker3.Year & "',@transtype = " & TYP & ",@what = " & What & "")
''If rec.RecordCount > 0 Then
''    For x = 1 To rec.RecordCount
''        'DoEvents
''                    Set z = ListView3.ListItems.Add(, , lst_Journal.ListItems(x).Text)
''                        z.SubItems(1) = lst_Journal.ListItems(x).SubItems(1)
''                        z.SubItems(2) = lst_Journal.ListItems(x).SubItems(2)
''                        z.SubItems(3) = lst_Journal.ListItems(x).SubItems(3)
''                        z.SubItems(4) = lst_Journal.ListItems(x).SubItems(4)
''
''                        z.SubItems(6) = lst_Bank.ListItems(y).Text
''                        z.SubItems(7) = lst_Bank.ListItems(y).SubItems(1)
''                        z.SubItems(8) = lst_Bank.ListItems(y).SubItems(2)
''                        z.SubItems(9) = lst_Bank.ListItems(y).SubItems(3)
''                        z.SubItems(10) = lst_Bank.ListItems(y).SubItems(4)
''        rec.MoveNext
''    Next x
''End If
''rec.Close
''Exit Sub
''bad:
''If What = 1 Then
''    If err.Number = 3704 Then
''    MsgBox "Please Iditify the Transaction type...", vbInformation, "System Message"
''    End If
''End If
''End Sub
'Private Sub Opt_GJ_Click()
'LoadOPTBook
'End Sub
'
'Private Sub opt_Statement_Click()
'LoadOPTBank
'End Sub
'
'Private Sub Opt_Withdrawals_Click()
'LoadOPTBank
'End Sub
'
'Private Sub Option4_Click()
'
'End Sub
'Private Sub toolbarSTat(ByVal Stat As String)
'
'If Stat = "New" Then
'    Toolbar1.Buttons(5).Caption = "&Save" 'save
'    Toolbar1.Buttons(5).Enabled = True
'    Toolbar1.Buttons(11).Enabled = True ' Cancel
'
'    Toolbar1.Buttons(3).Enabled = False ' Edit
'    Toolbar1.Buttons(7).Enabled = False 'match
'    Toolbar1.Buttons(9).Enabled = False 'delete
'
'    fme_Details(1).Visible = True
'    dt_checkdate.Value = Now
'    txtAmount.Text = ""
'    txtcheckno.Text = ""
'    txtdescription.Text = ""
'    txtID.Text = ""
'ElseIf Stat = "Edit" Then
'    Toolbar1.Buttons(5).Caption = "&Update"
'    Toolbar1.Buttons(5).Enabled = True
'    Toolbar1.Buttons(11).Enabled = True
'
'    Toolbar1.Buttons(3).Enabled = True
'    Toolbar1.Buttons(7).Enabled = False
'    Toolbar1.Buttons(9).Enabled = True
'    fme_Details(1).Visible = True
'ElseIf Stat = "Open" Then
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(11).Enabled = True
'
'    Toolbar1.Buttons(3).Enabled = True
'    Toolbar1.Buttons(7).Enabled = True
'    Toolbar1.Buttons(9).Enabled = True
'ElseIf Stat = "Cancel" Then
'    Toolbar1.Buttons(5).Caption = "&Save" 'save
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(11).Enabled = True ' Cancel
'
'    Toolbar1.Buttons(3).Enabled = False ' Edit
'    Toolbar1.Buttons(7).Enabled = False 'match
'    Toolbar1.Buttons(9).Enabled = False 'delete
'    fme_Details(1).Visible = False
'    dt_checkdate.Value = Now
'    txtAmount.Text = ""
'    txtcheckno.Text = ""
'    txtdescription.Text = ""
'    txtID.Text = ""
'End If
'End Sub
'
'Private Sub Option1_Click()
'
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim t As Integer
'If opt_Check.Value = True Then: t = opt_Check.Tag
'If opt_CR.Value = True Then: t = opt_CR.Tag
'If Opt_GJ.Value = True Then: t = Opt_GJ.Tag
'Select Case Button:
'    Case "&New":
'                Call toolbarSTat("New")
'     Case "&Edit":
'                With lst_Bank
'                    dt_checkdate.Value = .SelectedItem.Text
'                    txtAmount.Text = .SelectedItem.SubItems(3)
'                    txtcheckno.Text = .SelectedItem.SubItems(2)
'                    txtdescription.Text = .SelectedItem.SubItems(1)
'                    txtID.Text = .SelectedItem.SubItems(4)
'                End With
'                Call toolbarSTat("Edit")
'    Case "&Save":
'                If CheckEntry = True Then
'                    If MsgBox("Are you sure you want to Save this Entry?", vbQuestion + vbYesNo) = vbYes Then
'                        opndbaseFMIS.Execute ("INSERT INTO fmis.dbo.tblAMIS_BankReconCiliation(Fundtype,BankID,BankAccountno,Checkdate,[Description],Checkno,Amount,month_,Year_,datetimeentered,UserID,transtype,actioncode) " & _
'                        " VALUES  ('" & cmb_FundType.Text & "','" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "','" & cmb_Accountno.Text & "','" & dt_checkdate.Value & "','" & txtdescription.Text & "','" & txtcheckno.Text & "','" & txtAmount.Text & "','" & DTPicker3.Month & "','" & DTPicker3.Year & "','" & Now & "','" & ActiveUserID & "'," & t & ",1)")
'                        Call LoadReconData(2, lst_Bank, 2)
'                    End If
'                End If
'    Case "&Update":
'                If MsgBox("Are you sure you want to Update this Entry?", vbQuestion + vbYesNo) = vbYes Then
'                    opndbaseFMIS.Execute ("Update [tblAMIS_BankReconCiliation] set actioncode = 2,userid = '" & Trim(ActiveUserID) & "',datetimeentered = '" & Now & "' where trnno = " & txtID.Text & "")
'                        opndbaseFMIS.Execute ("INSERT INTO fmis.dbo.tblAMIS_BankReconCiliation(Fundtype,BankID,BankAccountno,Checkdate,[Description],Checkno,Amount,month_,Year_,datetimeentered,UserID,transtype,actioncode) " & _
'                        " VALUES  ('" & cmb_FundType.Text & "','" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "','" & cmb_Accountno.Text & "','" & dt_checkdate.Value & "','" & txtdescription.Text & "','" & txtcheckno.Text & "','" & txtAmount.Text & "','" & DTPicker3.Month & "','" & DTPicker3.Year & "','" & Now & "','" & ActiveUserID & "'," & t & ",1)")
'                        Call LoadReconData(2, lst_Bank, 2)
'                        Call toolbarSTat("Cancel")
'                End If
'    Case "&Match":
'                If MsgBox("Are you sure you want to System Matching?", vbQuestion + vbYesNo) = vbYes Then
'                    Call MatchTrans
'                End If
'    Case "&Delete":
'                If MsgBox("Are you sure you want to delete this Entry?", vbQuestion + vbYesNo) = vbYes Then
'                   opndbaseFMIS.Execute ("Update [tblAMIS_BankReconCiliation] set actioncode = 3,userid = '" & Trim(ActiveUserID) & "',datetimeentered = '" & Now & "' where trnno = " & lst_Bank.SelectedItem.SubItems(4) & "")
'                   Call LoadReconData(2, lst_Bank, 2)
'                End If
'    Case "&Cancel":
'                If MsgBox("Are you sure you want to Cancel the Entry?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
'                   Call toolbarSTat("Cancel")
'                End If
'    End Select
'End Sub
'Private Sub MatchTrans()
'Dim x, y As Long
'Dim z
'For y = 1 To lst_Bank.ListItems.Count
'        With lst_Journal
'            For x = 1 To .ListItems.Count
'                If lst_Bank.ListItems(y).SubItems(2) = .ListItems(x).SubItems(2) And lst_Bank.ListItems(y).SubItems(3) = .ListItems(x).SubItems(3) Then
'                    .HideSelection = False
'                    .ListItems(x).Selected = True
'                    opndbaseFMIS.Execute "Update [tblAMIS_FinalJEV] set bankmatch = 1 where trnno = " & .ListItems(x).SubItems(4) & " and actioncode = 1"
'                    opndbaseFMIS.Execute "Update [tblAMIS_BankReconCiliation] set havematch = 1 where trnno = " & lst_Bank.ListItems(y).SubItems(4) & " and actioncode = 1"
'                    Set z = ListView3.ListItems.Add(, , .ListItems(x).Text)
'                        z.SubItems(1) = .ListItems(x).SubItems(1)
'                        z.SubItems(2) = .ListItems(x).SubItems(2)
'                        z.SubItems(3) = .ListItems(x).SubItems(3)
'                        z.SubItems(4) = .ListItems(x).SubItems(4)
'
'                        z.SubItems(6) = lst_Bank.ListItems(y).Text
'                        z.SubItems(7) = lst_Bank.ListItems(y).SubItems(1)
'                        z.SubItems(8) = lst_Bank.ListItems(y).SubItems(2)
'                        z.SubItems(9) = lst_Bank.ListItems(y).SubItems(3)
'                        z.SubItems(10) = lst_Bank.ListItems(y).SubItems(4)
'                        DoEvents
'                    Exit For
'                End If
'            Next x
'        End With
'Next y
'End Sub
'Private Function CheckEntry() As Boolean
'Dim rec As New ADODB.Recordset
'CheckEntry = False
'If txtAmount.Text <> "" And txtcheckno.Text <> "" Then
'    CheckEntry = True
'Else
'    MsgBox "Please Specify the checkno and Amount..!", vbInformation, "System Message"
'    Exit Function
'End If
'Set rec = opndbaseFMIS.Execute("Select [Checkno] from [tblAMIS_BankReconCiliation] where checkno = '" & txtcheckno.Text & "' and actioncode = 1")
'    If rec.RecordCount > 0 Then
'        CheckEntry = False
'        MsgBox "Checkno Already Exist on the Database...!", vbInformation, "System Message"
'        Exit Function
'    Else
'        CheckEntry = True
'    End If
'rec.Close
'End Function
