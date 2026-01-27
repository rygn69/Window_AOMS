VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AP_ImportInJEV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts payable Importing Wizard"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "frm_AP_ImportInJEV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_AP_ImportInJEV.frx":3AFA
   ScaleHeight     =   8520
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   9330
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      Begin VB.Frame Frame1 
         Caption         =   "NB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton OptDebit 
            Caption         =   "Debit"
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
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptCredit 
            Caption         =   "Credit"
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
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   99680259
         UpDown          =   -1  'True
         CurrentDate     =   41326
      End
      Begin VB.ComboBox cmb_fundtype 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   8040
         Width           =   3495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   11668
         _Version        =   393216
         BackColor       =   16777215
         BackColorSel    =   8454143
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         GridLinesUnpopulated=   1
         SelectionMode   =   1
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
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   615
         Left            =   7200
         TabIndex        =   2
         ToolTipText     =   "Close form"
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "Add to Journal"
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
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   14737632
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AP_ImportInJEV.frx":75F4
         ImgSize         =   24
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H11 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   4
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
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AP_ImportInJEV.frx":8646
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   5520
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Accountcode:"
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
         Left            =   3480
         TabIndex        =   5
         Top             =   8085
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Accountcode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   285
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_AP_ImportInJEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nme, accntcode, field, Searchname, REFF As String
Public Trnno As Long
Public inRec As Boolean
Public frm As Form
Private Sub lvButtons_H3_Click()
On Error GoTo bad
Dim sql As String
Dim rec As New ADODB.Recordset
sql = ExecFunction("select [fmis].[dbo].GetqueryfromRtable (" & Combo1.ItemData(Combo1.ListIndex) & ")")
Set rec = opndbaseFMIS.Execute(sql)
If rec.RecordCount > 0 Then

Set MSFlexGrid1.DataSource = rec
End If
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub lvButtons_H5_Click()
frm_relatedtableForCOA.Show 1
End Sub


Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = &HFF&
End Sub


Private Sub lvButtons_H1_Click()
medll.centerme frmCDClaimantRegistry
frmCDClaimantRegistry.Show 1
End Sub

Private Sub lvButtons_H11_Click()
Call LoadAP
End Sub

Private Sub lvButtons_H6_Click()
Dim x As Integer
Dim xx As Variant
Dim Debit, Credit As Currency
Dim z
    
    ProgressBar1.Visible = True
    ProgressBar1.Max = MSHFlexGrid1.Rows - 1
    If MsgBox("Are you sure do you want to add into Journal?", vbInformation + vbYesNo, "System Message") = vbYes Then
    'Me.Visible = False
        For x = 1 To MSHFlexGrid1.Rows - 1
                With frmSub3
                    .Picture2.Visible = False
                    .cmbEntry.Visible = False
                    If OptCredit.Value = True Then
                    Credit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    Debit = 0
                    Else
                    Credit = 0
                    Debit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    End If
                    opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(REFF) & "','" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "','" & Debit & "','" & Credit & "')"
                End With
                DoEvents
                ProgressBar1.Value = x
        Next x
    End If
    ProgressBar1.Visible = False
End Sub

Private Sub MSHFlexGrid1_Click()
'If MSHFlexGrid1.Row <> 0 Then
'txtcode.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
'txtnme.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
'End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
If MSHFlexGrid1.Rows > 1 Then
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
        If Right(Trim(Text1.Text), 1) = "-" Then
        Text1.Text = Text1.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
        Else
            If Len(Trim(Text1.Text)) < 3 Then
            Text1.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            Else
            Text1.Text = Trim(Text1.Text) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            End If
        End If
    End If
End If
txtfind.Text = ""
Text1.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub MSHFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label6.ForeColor = &H80000008
End Sub
Public Function LoadAP()
'On Error Resume Next
Dim rec As New ADODB.Recordset
Dim SRec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim tmpquery As String
 
rec.Open "SELECT [Accountcode],[OBRNO],rtrim(ltrim(Convert(nchar,[Amount],101))) as Amount FROM [fmis].[dbo].[tblAMIS_Accounts_Payable] where Fundcode = " & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & " and year_ = " & DTPicker1.Year & " and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 3
        
        If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
        Else
            MSHFlexGrid1.Cols = 4
        End If
            MSHFlexGrid1.TextMatrix(0, 1) = "Accountcode"
            MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
            MSHFlexGrid1.TextMatrix(0, 3) = "Amount"
    
            MSHFlexGrid1.ColWidth(0) = 0
            MSHFlexGrid1.ColWidth(1) = 2500
            MSHFlexGrid1.ColWidth(2) = 4000
            MSHFlexGrid1.ColWidth(3) = 2000
            MSHFlexGrid1.ColAlignment(3) = 4
            Call GettotalAP
'rec.Close

Set rec = Nothing
Exit Function
bad:
MsgBox "Note: " & err.description, vbCritical, "System Message"
End Function
Private Sub GettotalAP()
On Error Resume Next
Dim x As Long
Text2.Text = "0.00"
For x = 1 To MSHFlexGrid1.Rows - 1
    Text2.Text = CCur(Text2.Text) + CCur(MSHFlexGrid1.TextMatrix(x, 3))
Next x
Text2.Text = Format(Text2.Text, "#,##0.00")
End Sub

Private Sub btnsave_Click()
Dim Subcode As Long
Dim lvl As Integer
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyDesc] (" & IIf(((Val(GetLvlbyCode(Text1.Text)) + 1) = 1), 2, Val(GetLvlbyCode(Text1.Text)) + 1) & ",'" & Trim(Text1.Text) & "','" & txtnme.Text & "')") = 1 Then
        MsgBox "Acocuntname is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyCODE] (" & IIf(((Val(GetLvlbyCode(Text1.Text)) + 1) = 1), 2, Val(GetLvlbyCode(Text1.Text)) + 1) & ",'" & Trim(Text1.Text) & "-" & Trim(txtcode.Text) & "')") = 1 Then
        MsgBox "Code is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    
        Subcode = txtcode.Text
        lvl = GetLvlbyCode(Text1.Text)
        If lvl = 0 Then
        lvl = 2
        Else
        lvl = Val(lvl) + 1
        End If
        
        opndbaseFMIS.Execute "Exec [Proc_CheckIfExistSub] @lvl = " & lvl & ",@childcode = 'Empty',@accountcode = '" & Text1.Text & "'," & _
        " @subcode =" & Subcode & ",@subdesc = '" & txtnme.Text & "'"

End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_fundtype)
End Sub

