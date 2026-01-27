VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_BeginAccountcodeSub1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subsidiary Ledger"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "frm_BeginAccountcodeSub1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H7 
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   50
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Import"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
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
      TabIndex        =   12
      Top             =   7920
      Width           =   1935
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
      TabIndex        =   11
      Top             =   7920
      Width           =   1935
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
      TabIndex        =   9
      Top             =   8010
      Width           =   3255
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
      TabIndex        =   8
      Top             =   7725
      Width           =   855
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
      TabIndex        =   7
      Top             =   7725
      Value           =   -1  'True
      Width           =   975
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
         Left            =   4320
         TabIndex        =   5
         Top             =   2520
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
      Interval        =   300
      Left            =   7440
      Top             =   0
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
      Top             =   540
      Width           =   8295
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10398
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
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   615
      Left            =   8520
      TabIndex        =   15
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "&Back"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_BeginAccountcodeSub1.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
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
      TabIndex        =   14
      Top             =   7950
      Width           =   855
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
      TabIndex        =   13
      Top             =   7950
      Width           =   855
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
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
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
Attribute VB_Name = "frm_BeginAccountcodeSub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Address, accountcode, Field1, Field2, Condition As String
Public Subcode1, Subdesc1, Subcode2, Subdesc2, Subcode3, Subdesc3, Subcode4, Subdesc4, Subcode5, Subdesc5, Subcode6, Subdesc6, Subcode7, Subdesc7 As String
Public col As Integer
Public fundcode As Integer
Public YEAR_ As Long
Public IfEdited As Boolean
Private Sub Form_Load()
txtaddress.Text = Address
Field1 = "Subcode" & col
Field2 = "Subdesc" & col
Call GetAccountNamesInGrid(Condition)
'txtcode.Text = GetmaxID
End Sub
Public Function GetAccountNamesInGrid(ByVal query As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
On Error GoTo bad
'Condition = Replace(Condition, "'", "")
'MsgBox "Select " & field & " as field," & Field2 & " as field2 from " & TableName & " where " & Condition & " and  actioncode = 0 group by " & field & "," & Field2 & " order by " & Order & ""
'rec.Open "Select " & field & " as field," & Field2 & " as field2,Sum(Sdebit),Sum(Scredit),max(lvl) from " & TableName & " where " & Condition & "  group by " & field & "," & Field2 & "  order by " & Order & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
Set rec = opndbaseFMIS.Execute(query)
    If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
        Call SetGrid
    End If
rec.Close
Call GetSum
Set rec = Nothing
Exit Function
bad:
If err.Number = -2147217900 Then
MsgBox "Please Identify the Subsidiary", vbInformation, "System Message"
opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where fundcode = '" & fundcode & "' and  Accountcode= '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'"
Call GetAccountNamesInGrid(Condition)
Else
MsgBox err.Description
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
Public Function GetAccountNamesInGridIndi(ByVal MSHFlex As MSHFlexGrid, ByVal TableName As String, ByVal field As String, ByVal Field2 As String, ByVal Condition As String, ByVal Order As String, nme As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim Subcode As String
Dim Subdesc As String
Subcode = "Subcode" & Val(col)
Subdesc = "Subdesc" & Val(col)
'Condition = Replace(Condition, "'", "")
'MsgBox "Select " & field & " as field," & Field2 & " as field2 from " & TableName & " where " & Condition & " and  actioncode = 0 group by " & field & "," & Field2 & " order by " & Order & ""
rec.Open "Select " & field & " as field," & Field2 & " as field2,Sum(Sdebit),Sum(Scredit),max(lvl) from " & TableName & " where " & Condition & " and " & Subdesc & " like '%" & nme & "%'  group by " & field & "," & Field2 & "," & Subdesc & " order by " & Subdesc & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
        Call SetGrid
    End If
rec.Close
Set rec = Nothing
End Function
Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader
Case "Accountcode"
Call GetAccountNames(LstAccountcode, "tblReff_CodeClassification", Field1, Field2, Condition, Field1)
Case "Accountname"
Call GetAccountNames(LstAccountcode, "tblReff_CodeClassification", Field1, Field2, Condition, Field2)
End Select
End Sub

Private Sub lvButtons_H7_Click()
If fundcode <> 0 And YEAR_ <> 0 Then
frm_AccntsPayImport.fundcode = fundcode
frm_AccntsPayImport.Accntcode = Me.Caption
frm_AccntsPayImport.YEAR_ = YEAR_
medll.CenterMe frm_AccntsPayImport
frm_AccntsPayImport.Show 1
Call Form_Load
End If
End Sub

Private Sub MSHFlexGrid1_Click()
On Error GoTo bad
    Select Case MSHFlexGrid1.col
    Case 3 To 4 'Debit/Credit
        If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'," & col & ")") > 1 Then
        txt_entry.Visible = False
        Call MSHFlexGrid1_DblClick
        Else
            txt_entry.Move MSHFlexGrid1.CellLeft, MSHFlexGrid1.CellTop, MSHFlexGrid1.CellWidth, MSHFlexGrid1.CellHeight
            txt_entry.Visible = True
            If Len(Trim(MSHFlexGrid1.Text)) <> 0 Then
                txt_entry.Text = MSHFlexGrid1.Text
                txt_entry.SelStart = 0
                txt_entry.SelLength = Len(txt_entry.Text)
            Else
                txt_entry.Text = ""
            End If
            txt_entry.SetFocus
        End If
    Case Else
        txt_entry.Visible = False
    End Select
Exit Sub
bad:
MsgBox err.Description
End Sub
Public Sub getAmount()
Dim x As Integer
For x = 1 To MSHFlexGrid1.Rows - 1
    
Next x
End Sub
Private Sub MSHFlexGrid1_DblClick()
Dim Newform As New frm_BeginAccountcodeSub1
If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) = "" Then
Exit Sub
End If
If MSHFlexGrid1.Rows - 1 <> 0 Then
    If col + 1 > 7 Then
        MsgBox "Sory this is the end of the SubName", vbInformation, "System Message"
        Exit Sub
    End If
    Newform.col = Val(col) + 1
    Newform.Address = Address & "~" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
    With Newform
        Select Case (Newform.col):
        Case 0
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = Subcode2
            .Subdesc2 = Subdesc2
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ",@year = '" & YEAR_ & "'"
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "subcode2 is not null And (fundcode = " & fundcode & " or fundcode is null)"
            .Caption = Trim(Subcode1)
            .fundcode = fundcode
            .YEAR_ = YEAR_
        Case 2
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = Subcode2
            .Subdesc2 = Subdesc2
'            .Condition = "Subcode1 = " & Subcode1 & " and " & "subcode2 is not null And (fundcode = " & fundcode & " or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
            .Caption = Trim(Subcode1) & "-" & Trim(Subcode2)
            .fundcode = fundcode
            .YEAR_ = YEAR_
        Case 3
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
            .Subdesc2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "subcode2 = " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & " and " & "subcode3 is not null And (fundcode = " & fundcode & " or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
            .Caption = Trim(Subcode1) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            .fundcode = fundcode
            .YEAR_ = YEAR_
        Case 4
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = Subcode2
            .Subdesc2 = Subdesc2
            .Subcode3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
            .Subdesc3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "subcode3 = " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & " and " & "subcode4 is not null And (fundcode = " & fundcode & "or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
           .Caption = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
           .fundcode = fundcode
           .YEAR_ = YEAR_
        Case 5
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = Subcode2
            .Subdesc2 = Subdesc2
            .Subcode3 = Subcode3
            .Subdesc3 = Subdesc3
            .Subcode4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
            .Subdesc4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "subcode4 = " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & " and " & "subcode5 is not null And (fundcode = " & fundcode & "or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
            .Caption = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            .fundcode = fundcode
            .YEAR_ = YEAR_
        Case 6
            .Subcode1 = Subcode1
            .Subdesc1 = Subdesc1
            .Subcode2 = Subcode2
            .Subdesc2 = Subdesc2
            .Subcode3 = Subcode3
            .Subdesc3 = Subdesc3
            .Subcode4 = Subcode4
            .Subdesc4 = Subdesc4
            .Subcode5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
            .Subdesc5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "subcode5 = " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & " and " & "subcode6 is not null And (fundcode = " & fundcode & "or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
            .Caption = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            .fundcode = fundcode
            .YEAR_ = YEAR_
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
            .Subcode6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
            .Subdesc6 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
            '.Condition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "Subcode5 = " & Subcode5 & " and " & "subcode6 = " & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & " and " & "subcode7 is not null And (fundcode = " & fundcode & "or fundcode is null)"
            .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(fundcode) & "',@Accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =" & Newform.col & ", @year = '" & YEAR_ & "'"
            .Caption = Trim(Subcode1) & "-" & Trim(Subcode2) & "-" & Trim(Subcode3) & "-" & Trim(Subcode4) & "-" & Trim(Subcode5) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            .fundcode = fundcode
            .YEAR_ = YEAR_
        End Select
    End With
    Newform.Show 1
Call Form_Load
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

rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
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

Private Sub lvButtons_H4_Click()
Unload Me
End Sub


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
    If MSHFlexGrid1.TextMatrix(x, 3) <> "" And Val(MSHFlexGrid1.TextMatrix(x, 3)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 3) = Format(MSHFlexGrid1.TextMatrix(x, 3), "#,###.00")
    
    Else
    MSHFlexGrid1.TextMatrix(x, 3) = ""
    End If
    
    If MSHFlexGrid1.TextMatrix(x, 4) <> "" And Val(MSHFlexGrid1.TextMatrix(x, 4)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "#,###.00")
    Else
    MSHFlexGrid1.TextMatrix(x, 4) = ""
End If
Next x
End Sub

Private Sub Timer1_Timer()
DoEvents
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
 'On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = Format((txt_entry.Text), "#,##0.00")
                If MSHFlexGrid1.col = 3 Then
                    If Trim(txt_entry.Text) <> "" Then
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = ""
                    Else
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = ""
                    End If
                
                ElseIf MSHFlexGrid1.col <> 5 Then
                    
                    If Trim(txt_entry.Text) <> "" Then
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = ""
                    Else
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = ""
                    End If
                End If
                txt_entry.Visible = False
                If MSHFlexGrid1.col = 5 Then
                    If txt_entry.Text = "1" Or txt_entry.Text = "5" Then
                    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = txt_entry.Text
                    Else
                    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = "1"
                    End If
                End If
                    Call SaveAmount(IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)), IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)))
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.Description)
End Sub
Public Function SaveAmount(ByVal Debit As Currency, ByVal Credit As Currency)
Dim rec As New ADODB.Recordset

rec.Open "Select Accountcode from tblAMIS_Begeningbalance where accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and fundcode = '" & fundcode & "' AND YEAR_ = '" & YEAR_ & "'", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        opndbaseFMIS.Execute "Update tblAMIS_Begeningbalance set debit = '" & Debit & "',credit = '" & Credit & "' where accountcode = '" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and fundcode = '" & fundcode & "' AND YEAR_ = '" & YEAR_ & "' "
    Else
        opndbaseFMIS.Execute "Insert into tblAMIS_Begeningbalance (accountcode,debit,Credit,fundcode,actioncode,YEAR_) values ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "','" & Debit & "','" & Credit & "','" & fundcode & "',1,'" & YEAR_ & "')"
    End If
rec.Close
End Function

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim sql As String
If KeyAscii = 13 Then
    If optCode.Value = True Then
        If IsNumeric(txtsearch.Text) = True Then
        sql = "EXECUTE  [fmis].[dbo].[MPproc_LoadJEVfrombeginIndi] @fundcode = '" & fundcode & "',@Accountcode = '" & Me.Caption & "',@lvl = " & col & ",@Type = 1,@code = '" & txtsearch.Text & "',@name = '',@year = '" & YEAR_ & "'"
        Else
            MsgBox "None Numeric entry"
            Exit Sub
        End If
    Else
            sql = "EXECUTE  [fmis].[dbo].[MPproc_LoadJEVfrombeginIndi] @fundcode= '" & fundcode & "',@Accountcode = '" & Me.Caption & "',@lvl = " & col & ",@Type = 2,@code = 0,@name = '" & txtsearch.Text & "',@year = '" & YEAR_ & "'"
        
    End If
        If Trim(txtsearch.Text) <> "" Then
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
End If
End Sub
