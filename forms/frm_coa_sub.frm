VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_coa_sub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frm_coa_sub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   7605
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Delete Account"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_coa_sub.frx":076A
      cBack           =   -2147483633
   End
   Begin VB.CheckBox optAutoclose 
      Caption         =   "Auto Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      Top             =   0
      Width           =   1455
   End
   Begin VB.OptionButton optName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   15
      Top             =   7750
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optCode 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   14
      Top             =   7750
      Width           =   855
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   12
      Top             =   8040
      Width           =   3135
   End
   Begin VB.TextBox txtDebit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtCredit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   8040
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   9345
      TabIndex        =   3
      Top             =   1200
      Width           =   9375
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6960
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11245
         _Version        =   393216
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6960
      Top             =   -120
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   8295
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accountcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accountname"
         Object.Width           =   10583
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3600
      Top             =   1440
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
      TextControl     =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_coa_sub.frx":4274
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Close all"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_coa_sub.frx":45C6
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   8040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_coa_sub.frx":4918
      cBack           =   16777215
   End
   Begin VB.Label Label4 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7710
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Debit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   8070
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Credit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   8070
      Width           =   855
   End
   Begin VB.Label lblaccountcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_coa_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Address, accountcode, Field1, Field2, Condition, dvno, Fullcode, Accntname As String
Public Subcode1, Subdesc1, Subcode2, Subdesc2, Subcode3, Subdesc3, Subcode4, Subdesc4, Subcode5, Subdesc5, Subcode6, Subdesc6, Subcode7, Subdesc7 As String
Public col As Integer
Public fundcode As Integer
Public Gamount As Currency
Public IfEdited, Isclose As Boolean
Public frm As Form
Public IsDbclick As Boolean
Public isPOSTED As Boolean
Private Sub Form_Load()
On Error GoTo bad
    Me.Caption = Trim(Fullcode)
    txtaddress.Text = Address
    Field1 = "Subcode" & col
    Field2 = "Subdesc" & col
    If Autoclose = True Then
    optAutoclose.Value = 1
    Else
    optAutoclose.Value = 0
    End If
    Call Timer1_Timer
    Call GetAccountNamesInGrid(Condition)
Exit Sub
Call LoadErr(err.Number, Me.name, err.description)
'txtcode.Text = GetmaxID
Exit Sub
bad:
Call LoadErr(err.Number, Me.name & "-Form_load", err.description)
End Sub
Public Function GetAccountNamesInGrid(ByVal query As String)
Dim rec As New ADODB.Recordset
On Error GoTo bad
Dim x
Dim z As Integer
'Condition = Replace(Condition, "'", "")
'MsgBox "Select " & field & " as field," & Field2 & " as field2 from " & TableName & " where " & Condition & " and  actioncode = 0 group by " & field & "," & Field2 & " order by " & Order & ""
DoEvents
'MsgBox query
Set rec = opndbaseFMIS.Execute(query)
    If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
        Call SetGrid
    End If
    GetSum
rec.Close
Set rec = Nothing
Exit Function
bad:
If err.Number = -2147217900 Then
MsgBox "Please Identify the Subsidiary", vbInformation, "System Message"
opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where dvno = '" & dvno & "' and Accountcode= '" & Trim(Me.Caption) & "'"
Call GetAccountNamesInGrid(Condition)
Else
MsgBox err.description
End If
End Function
Private Sub GetSum()
On Error Resume Next
Dim Debit As Currency
Dim Credit As Currency
Dim x As Long
For x = 1 To MSHFlexGrid1.Rows - 1
    If Trim(MSHFlexGrid1.TextMatrix(x, 3)) <> "" Then
    Debit = CDbl(Debit) + CDbl(MSHFlexGrid1.TextMatrix(x, 3))
    End If
    If Trim(MSHFlexGrid1.TextMatrix(x, 4)) <> "" Then
    Credit = CDbl(Credit) + CDbl(MSHFlexGrid1.TextMatrix(x, 4))
    End If
    DoEvents
Next x
txtDebit.Text = Format(Debit, "#,##0.00")
txtCredit.Text = Format(Credit, "#,##0.00")
End Sub
Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader
Case "Accountcode"
Call GetAccountNames(LstAccountcode, "tblReff_CodeClassification", Field1, Field2, Condition, Field1)
Case "Accountname"
Call GetAccountNames(LstAccountcode, "tblReff_CodeClassification", Field1, Field2, Condition, Field2)
End Select
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H2_Click()
On Error Resume Next
frm.Isclose = True
Unload Me
End Sub

Private Sub lvButtons_H3_Click()
With frm_AccountSpecialEntry
    Set .frm = Me
    .Searchname = txtsearch.Text
    .field = Accntname
    .Text1.Text = Me.Caption
    .txtdetails.Text = txtaddress.Text
    .Show 1
    Call LoadAccountbyname
End With
End Sub

Private Sub lvButtons_H4_Click()
If MsgBox("Are you sure Do you want to DELETE this Account?", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
opndbaseFMIS.Execute "Update tblAMIS_tmpJournal set debit = '" & Debit & "',credit = '" & Credit & "' where ltrim(rtrim(accountcode)) + '-' Like '" & Trim(Me.Caption) & "-%" & "' and dvno = '" & dvno & "' "
Unload Me
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim Newform As New frm_SubAccountcode
On Error GoTo bad
    Select Case MSHFlexGrid1.col
    Case 3 To 4 'Debit/Credit
        If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'," & col & ")") > 1 Then
        txt_entry.Visible = False
        IsDbclick = True
        
                If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) = "" Then
                Exit Sub
                End If
                
                If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'," & col & ")") = 0 Then
                        txt_entry.Visible = False
                        MsgBox "End of the Accounts.", vbInformation, "System Message"
                      Exit Sub
                End If
                If MSHFlexGrid1.Rows - 1 <> 0 Then
                    If col + 1 > 7 Then
                        MsgBox "Sory this is the end of the SubName", vbInformation, "System Message"
                        Exit Sub
                    End If
                    Newform.col = val(col) + 1
                    Newform.Address = Address & "~" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                    With Newform
                        Select Case (Newform.col):
                        Case 0
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .dvno = dvno
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(Subcode2) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1)
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 2
                            .Accntname = Subdesc2
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .dvno = dvno
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2)
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 3
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .dvno = dvno
                            .Subcode2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                            .Subdesc2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 4
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .dvno = dvno
                            .Subcode3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                            .Subdesc3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                           .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                           .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                           .fundcode = fundcode
                           .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 5
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .Subcode3 = Subcode3
                            .Subdesc3 = Subdesc3
                            .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .dvno = dvno
                            .Subcode4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                            .Subdesc4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 6
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .Subcode3 = Subcode3
                            .Subdesc3 = Subdesc3
                            .Subcode4 = Subcode4
                            .Subdesc4 = Subdesc4
                            .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .dvno = dvno
                            .Subcode5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                            .Subdesc5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        Case 7
                            .Subcode1 = Subcode1
                            .Subdesc1 = Subdesc1
                            .Subcode2 = Subcode2
                            .Subdesc2 = Subdesc2
                            .Subcode3 = Subcode3
                            .Subdesc3 = Subdesc3
                            .Subcode4 = Subcode4
                            .Subdesc4 = Subdesc4
                            .Subcode5 = Subcode5
                            .Subdesc5 = Subdesc5
                            .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .dvno = dvno
                            .Subcode6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                            .Subdesc6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                            .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(Subcode5) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                            .fundcode = fundcode
                            .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
                        End Select
                    End With
                    Set Newform.frm = Me
                    Newform.Show 1
                If Autoclose = True Then
                    Unload Me
                Else
                    Call Form_Load
                End If
    End If
        Else
            txt_entry.Move MSHFlexGrid1.CellLeft, MSHFlexGrid1.CellTop, MSHFlexGrid1.CellWidth, MSHFlexGrid1.CellHeight
            txt_entry.Visible = True
            If Len(Trim(MSHFlexGrid1.Text)) <> 0 Then
                txt_entry.Text = MSHFlexGrid1.Text
                txt_entry.SelStart = 0
                txt_entry.SelLength = Len(txt_entry.Text)
            Else
                If MSHFlexGrid1.col = 3 Then
                    txt_entry.Text = Gamount - CCur(txtDebit.Text)
                ElseIf MSHFlexGrid1.col = 4 Then
                    txt_entry.Text = Gamount - CCur(txtCredit.Text)
                End If
                txt_entry.SelStart = 0
                txt_entry.SelLength = Len(txt_entry.Text)
            End If
            txt_entry.SetFocus
        End If
    Case Else
        txt_entry.Visible = False
    End Select
Exit Sub
bad:
Call LoadErr(err.Number, Me.name, err.description)
End Sub
Public Sub getAmount()
Dim x As Integer
For x = 1 To MSHFlexGrid1.Rows - 1
    
Next x
End Sub
Private Sub MSHFlexGrid1_DblClick()
Dim Newform As New frm_SubAccountcode


If MSHFlexGrid1.col = 2 Then
    If MSHFlexGrid1.col > 2 And IsDbclick = False Then
        Exit Sub
    End If
    
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) = "" Then
    Exit Sub
    End If
    
    If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'," & col & ")") = 0 Then
            txt_entry.Visible = False
            MsgBox "End of the Accounts.", vbInformation, "System Message"
          Exit Sub
    End If
    If MSHFlexGrid1.Rows - 1 <> 0 Then
        If col + 1 > 7 Then
            MsgBox "Sory this is the end of the SubName", vbInformation, "System Message"
            Exit Sub
        End If
        Newform.col = val(col) + 1
        Newform.Address = Address & "~" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
        With Newform
            Select Case (Newform.col):
            Case 0
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .dvno = dvno
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(Subcode2) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1)
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 2
                .Accntname = Subdesc2
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .dvno = dvno
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2)
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 3
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .dvno = dvno
                .Subcode2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                .Subdesc2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 4
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .dvno = dvno
                .Subcode3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                .Subdesc3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
               .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
               .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
               .fundcode = fundcode
               .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 5
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .Subcode3 = Subcode3
                .Subdesc3 = Subdesc3
                .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .dvno = dvno
                .Subcode4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                .Subdesc4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 6
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .Subcode3 = Subcode3
                .Subdesc3 = Subdesc3
                .Subcode4 = Subcode4
                .Subdesc4 = Subdesc4
                .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .dvno = dvno
                .Subcode5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                .Subdesc5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            Case 7
                .Subcode1 = Subcode1
                .Subdesc1 = Subdesc1
                .Subcode2 = Subcode2
                .Subdesc2 = Subdesc2
                .Subcode3 = Subcode3
                .Subdesc3 = Subdesc3
                .Subcode4 = Subcode4
                .Subdesc4 = Subdesc4
                .Subcode5 = Subcode5
                .Subdesc5 = Subdesc5
                .Accntname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .dvno = dvno
                .Subcode6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                .Subdesc6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(dvno) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ""
                .Fullcode = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(Subcode5) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                .fundcode = fundcode
                .Gamount = Gamount - CCur(IIf(CCur(txtDebit.Text) = 0, CCur(txtCredit.Text), CCur(txtDebit.Text)))
            End Select
        End With
        Set Newform.frm = Me
        Newform.Show 1
    If Autoclose = True Then
        Unload Me
    Else
        Call Form_Load
    End If
    End If
End If
End Sub
Private Function IfCodeIsUses(ByVal id As Double) As Boolean
Dim rec As New ADODB.Recordset
IfCodeIsUses = False
rec.Open "Select top 1 id from tblAMIS_FinalJEV where id = " & id & ""
    If rec.RecordCount > 0 Then
        IfCodeIsUses = True
    End If
rec.Close
End Function
Private Function IfUsecode(ByVal accountcode As String)
Dim rec As New ADODB.Recordset

End Function
Public Function IfExistname(ByVal name As String) As Boolean
Dim x As Integer
IfExistname = False
    For x = 1 To MSHFlexGrid1.Rows - 1
        If UCase(name) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) Then
            IfExistname = True
            Exit For
        End If
    Next x
End Function
Public Function IfExistcode(ByVal Code As String) As Boolean
Dim x As Integer
IfExistcode = False
    For x = 1 To MSHFlexGrid1.Rows - 1
        If Trim(Code) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) Then
            IfExistcode = True
            Exit For
        End If
    Next x
End Function
Public Function GetmaxID()
Dim rec As New ADODB.Recordset
Dim sql As String

Select Case col:
Case 1
Case 2
    sql = "Select   max(subcode2) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 3
    sql = "Select max(subcode3) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 4
sql = "Select max(subcode4) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 5
    sql = "Select max(subcode5) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 6
    sql = "Select max(subcode6) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 7
    sql = "Select max(subcode7) as maxid from tblReff_CodeClassification where " & Getcondition & ""
End Select

Set rec = opndbaseFMIS.Execute(sql)
    If rec.RecordCount > 0 Then
        GetmaxID = IIf(IsNull(rec!maxid), 0, rec!maxid) + 1
    Else
        GetmaxID = 1
    End If
rec.Close
Set rec = Nothing
End Function
Private Function Getcondition() As String
        Select Case (col):
        Case 2
            Getcondition = "Subcode1 = " & Subcode1
        Case 3
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "subcode2 = " & Subcode2
        Case 4
           Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "subcode3 = " & Subcode3
        Case 5
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "subcode4 = " & Subcode4
        Case 6
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "subcode5 = " & Subcode5
        Case 7
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "Subcode5 = " & Subcode5 & " and " & "subcode6 = " & Subcode6
        End Select
End Function




Private Sub txtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtcode.Text = GetmaxID
End If
End Sub
Public Sub SetGrid()
        MSHFlexGrid1.Cols = 5
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Debit"
        MSHFlexGrid1.TextMatrix(0, 4) = "Credit"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 5200
        MSHFlexGrid1.ColWidth(3) = 1500
        MSHFlexGrid1.ColWidth(4) = 1500
        MSHFlexGrid1.ColWidth(5) = 0
For x = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(x, 3) <> "" And val(MSHFlexGrid1.TextMatrix(x, 3)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 3) = Format(MSHFlexGrid1.TextMatrix(x, 3), "#,###.00")
    
    Else
    MSHFlexGrid1.TextMatrix(x, 3) = ""
    End If
    
    If MSHFlexGrid1.TextMatrix(x, 4) <> "" And val(MSHFlexGrid1.TextMatrix(x, 4)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "#,###.00")
    Else
    MSHFlexGrid1.TextMatrix(x, 4) = ""
End If
Next x
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
Call lvButtons_H1_Click
End Sub

Private Sub MShFlexGrid1_Scroll()
txt_entry.Visible = False
End Sub

Private Sub optAutoclose_Click()
If optAutoclose.Value = 1 Then
Autoclose = True
Else
Autoclose = False
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call LoadAccountbyname
End If
End Sub
Public Sub LoadAccountbyname()
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim sql As String
    If optCode.Value = True Then
        If IsNumeric(txtsearch.Text) = True Then
        sql = "EXECUTE  [fmis].[dbo].[MPproc_LoadJEVfromtmpIndi] @Dvno = '" & dvno & "',@Accountcode = '" & Trim(Me.Caption) & "',@lvl = " & col & ",@Type = 1,@code = '" & Replace(txtsearch.Text, "'", "''") & "',@name = ''"
        Else
            If Trim(txtsearch.Text) <> "" Then
                MsgBox "None Numeric entry"
                Exit Sub
            Else
                Call Form_Load
            End If
        End If
    Else
            sql = "EXECUTE  [fmis].[dbo].[MPproc_LoadJEVfromtmpIndi] @Dvno = '" & dvno & "',@Accountcode = '" & Trim(Me.Caption) & "',@lvl = " & col & ",@Type = 2,@code = 0,@name = '" & Replace(txtsearch.Text, "'", "''") & "'"
        
    End If
        If Trim(txtsearch.Text) <> "" Then
       ' MsgBox sql
            Set rec = opndbaseFMIS.Execute(sql)
                If rec.RecordCount > 0 Then
                    Set MSHFlexGrid1.DataSource = rec
                    Call SetGrid
                End If
                'GetSum
            rec.Close
            Set rec = Nothing
        Else
            Call Form_Load
        End If
End Sub
Private Sub Timer1_Timer()
DoEvents
End Sub
Private Sub txt_entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If isPOSTED = True Then
                MsgBox "Unable to Edit the Entry, the Transaction is Already generate the report", vbInformation, "System Information"
                txt_entry.Visible = False
                Exit Sub
            End If
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                If InStr(1, txt_entry.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
            If Right(txt_entry.Text, 1) = "+" Then
                 MsgBox "Invalid format, Please Check Your Entry", vbCritical, "System Message"
                 Exit Sub
            End If
            txt_entry.Text = sumAmount(txt_entry.Text)
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = Format((txt_entry.Text), "#,##0.00")
                
                txt_entry.Visible = False
                
                Call SaveAmount(IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)), IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)))
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Function sumAmount(ByVal amnt As String) As String
On Error GoTo sum
Dim x As Integer
Dim y As String
Dim str() As String
    If Left(amnt, 1) = "+" Then
    Else
    amnt = "+" & amnt
    End If
 
    str = Split(Trim(amnt), "+", -1, vbTextCompare)
    y = 0

 For x = 1 To 1000
    y = CCur(y) + CCur(str(x))
 Next x
 Exit Function
sum:
If err.Number = 9 Then
 sumAmount = y
ElseIf err.Number = 13 Then

Else
MsgBox "Incorrect Format", vbInformation, "System Message"
End If
End Function
Public Function SaveAmount(ByVal Debit As Currency, ByVal Credit As Currency)
Dim rec As New ADODB.Recordset
If Debit = 0 And Credit = 0 Then
opndbaseFMIS.Execute "Delete from tblAMIS_tmpJournal where accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and dvno = '" & dvno & "' "
Else
    Set rec = opndbaseFMIS.Execute("Select Accountcode from tblAMIS_tmpjournal where accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and dvno = '" & dvno & "'")
        If rec.RecordCount > 0 Then
            opndbaseFMIS.Execute "Update tblAMIS_tmpJournal set debit = '" & Debit & "',credit = '" & Credit & "' where accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and dvno = '" & dvno & "' "
        Else
            opndbaseFMIS.Execute "Insert into tblAMIS_tmpJournal (accountcode,debit,Credit,dvno) values ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "','" & Debit & "','" & Credit & "','" & Trim(dvno) & "')"
        End If
        GetSum
    rec.Close
End If
End Function
