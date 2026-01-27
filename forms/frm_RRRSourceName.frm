VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_RRRSourceName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report of Revenue and Receipts Classification"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "frm_RRRSourceName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_RRRSourceName.frx":076A
   ScaleHeight     =   8685
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1080
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtSourcename 
      Appearance      =   0  'Flat
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2040
      Width           =   3855
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accountname"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Accountcode"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox cmb_Accountcode 
      Height          =   285
      Left            =   7680
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   8880
      TabIndex        =   0
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Cancel"
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
      ImgAlign        =   1
      Image           =   "frm_RRRSourceName.frx":AE19
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
      CapAlign        =   2
      BackStyle       =   2
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
      Image           =   "frm_RRRSourceName.frx":E923
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Add"
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
      Image           =   "frm_RRRSourceName.frx":1242D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   8880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Generate"
      CapAlign        =   2
      BackStyle       =   2
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
      Image           =   "frm_RRRSourceName.frx":15F37
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   9120
      Top             =   -120
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
            Picture         =   "frm_RRRSourceName.frx":19A41
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":1B3D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":1CD65
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":1E6F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":20089
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":21A1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":233AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":24D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":266D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":28065
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":28D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":29621
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":2A2FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":2AFD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":2BCB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":2C991
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_RRRSourceName.frx":2D66D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   5640
      Top             =   3120
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
      CapAlign        =   2
      BackStyle       =   2
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
      Image           =   "frm_RRRSourceName.frx":2DF49
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   9120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&New"
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
      Image           =   "frm_RRRSourceName.frx":31A53
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H7 
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   9120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Delete"
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
      Image           =   "frm_RRRSourceName.frx":3555D
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H8 
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   8880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Edit"
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
      Image           =   "frm_RRRSourceName.frx":39067
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H10 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&New"
      CapAlign        =   2
      BackStyle       =   4
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
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_RRRSourceName.frx":3CB71
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H9 
      Height          =   375
      Left            =   1125
      TabIndex        =   12
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   4
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
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_RRRSourceName.frx":3D7C3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H11 
      Height          =   375
      Left            =   2130
      TabIndex        =   15
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   4
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
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_RRRSourceName.frx":3DB15
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H12 
      Height          =   375
      Left            =   3130
      TabIndex        =   19
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   4
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
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_RRRSourceName.frx":4161F
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
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
      Left            =   240
      TabIndex        =   18
      Top             =   2079
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1600
      Width           =   1215
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue and Receipts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   390
      Width           =   5010
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10080
      Y1              =   7920
      Y2              =   7920
   End
End
Attribute VB_Name = "frm_RRRSourceName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public debitcredit As Integer
Public RRRCode As String
Public Code As String
Public Descrip As String
Public IsOK As Boolean

Private Sub Form_Load()
Call GetAccountNamebyorder(LstAccountcode, "SourceName")
End Sub
Public Function GetAccountNamebyorder(ByVal lst As ListView, ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Condition = Replace(Condition, "'", "")
rec.Open "Select trnno,[SourceName],[Accountcode] from fmis.dbo.tblAMIS_RRRSourceName where trnno not in ((SELECT [Subcode3] FROM [fmis].[dbo].[tblReff_RRRClassification] where subcode3 is not null group by subcode3))order by " & Condition & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                    Set x = lst.ListItems.Add(, , rec!Trnno)
                    x.SubItems(1) = Trim(rec.Fields!SourceName)
                    x.SubItems(2) = Trim(IIf(IsNull(rec.Fields!accountcode), "", rec.Fields!accountcode))
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
End Function

Private Sub LstAccountcode_Click()
lvButtons_H9.Enabled = False
txtID.Enabled = False
txtSourcename.Enabled = False
txtID.Text = LstAccountcode.SelectedItem.Text
txtSourcename.Text = LstAccountcode.SelectedItem.SubItems(1)
End Sub

Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader = "Accountname" Then
Call GetAccountNamebyorder(LstAccountcode, "SourceName")
ElseIf ColumnHeader = "Accountcode" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
ElseIf ColumnHeader = "ID" Then
Call GetAccountNamebyorder(LstAccountcode, "trnno")
End If
End Sub

Private Sub LstAccountcode_DblClick()
Dim rec As New ADODB.Recordset
'With frm_AccountView
'    Set .frm = Me
'    cmb_Accountcode.Text = LstAccountcode.SelectedItem.SubItems(2)
'    .Text1.Text = cmb_Accountcode.Text
'    .Show 1
'
'    LstAccountcode.SelectedItem.SubItems(2) = cmb_Accountcode.Text
'End With
'

With frmforCOA
    .nme = Trim(LstAccountcode.SelectedItem.SubItems(1))
    .accntcode = Trim(LstAccountcode.SelectedItem.SubItems(2))
    .Trnno = LstAccountcode.SelectedItem.Text
    Set .frm = Me
    .Show 1
End With
If IsOK = True Then
      rec.Open "select tables,columns,conditions from tblAMIS_RelatedTableforCOA where trnno = 8", opndbaseFMIS, adOpenStatic
      If rec.RecordCount > 0 Then
        Call UpdateExtractor(IIf(IsNull(rec!Tables), "", rec!Tables), IIf(IsNull(rec!columns), "", rec!columns), Trim(LstAccountcode.SelectedItem.SubItems(2)), IIf(IsNull(rec!Conditions), "", rec!Conditions), LstAccountcode.SelectedItem.Text)
      End If
      rec.Close
End If
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H10_Click()
lvButtons_H9.Enabled = True
txtID.Enabled = True
txtSourcename.Enabled = True
txtID.Text = ""
txtSourcename.Text = ""
lvButtons_H9.Caption = "&Save"
End Sub

Private Sub lvButtons_H11_Click()
lvButtons_H9.Enabled = True
txtSourcename.Enabled = True
lvButtons_H9.Caption = "&Update"
End Sub

Private Sub lvButtons_H12_Click()
If MsgBox("Are you sure do you want to Delete this entry?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "Delete from [tblAMIS_RRRSourceName]where trnno ='" & txtID.Text & "'"
    MsgBox "Successfully Deleted..!", vbInformation, "System Information"
    Call GetAccountNamebyorder(LstAccountcode, "SourceName")
End If
End Sub

Private Sub lvButtons_H3_Click()
Dim z As Long
Dim IfCheck As Boolean
IfCheck = False
For z = 1 To LstAccountcode.ListItems.Count
        If LstAccountcode.ListItems(z).Checked = True Then
            IfCheck = True
            Exit For
        End If
Next z
If IfCheck = False Then
    MsgBox "Please Check the account that you want to ADD.."
    Exit Sub
End If
If MsgBox("Are you sure Do you want to Add all checked account?", vbInformation + vbYesNo, "System Message") = vbYes Then
    For z = 1 To LstAccountcode.ListItems.Count
        If LstAccountcode.ListItems(z).Checked = True Then
'            If Trim(LstAccountcode.ListItems(z).ListSubItems(1).Text) <> "" Then
                opndbaseFMIS.Execute "EXECUTE  [fmis].[dbo].[MPproc_SaveToRRREstimated] @ID = '" & Val(LstAccountcode.ListItems(z).Text) & "',@rrrcode = '" & RRRCode & "',@descCrip = '" & Replace(LstAccountcode.ListItems(z).ListSubItems(1).Text, "'", "''") & "'"
            'End If
        End If
    Next z
End If
End Sub

Private Sub lvButtons_H9_Click()
If lvButtons_H9.Caption = "&Update" Then
    If MsgBox("Are you sure do you want to Update this entry?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Update [tblAMIS_RRRSourceName] set SourceName = '" & Replace(txtSourcename.Text, "'", "''") & "' where trnno = '" & txtID.Text & "'"
        MsgBox "Successfully Updated..!", vbInformation, "System Information"
        Call GetAccountNamebyorder(LstAccountcode, "SourceName")
    End If
ElseIf lvButtons_H9.Caption = "&Save" Then
    If MsgBox("Are you sure do you want to Save this entry?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Insert into [tblAMIS_RRRSourceName]([SourceName]) values ('" & Replace(txtSourcename.Text, "'", "''") & "')"
        MsgBox "Successfully Save..!", vbInformation, "System Information"
        Call GetAccountNamebyorder(LstAccountcode, "SourceName")
    End If
End If
End Sub
