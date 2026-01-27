VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AccountsPayable 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable Registry"
   ClientHeight    =   9375
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_AccountsPayable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_AccountsPayable.frx":076A
   ScaleHeight     =   9375
   ScaleWidth      =   14640
   Begin VB.Frame Frame2 
      Caption         =   "Details"
      Height          =   7215
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   14415
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   6720
         Width           =   2535
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   11033
         _Version        =   393216
         BackColorBkg    =   -2147483634
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
      Begin VB.ComboBox cmb_accountcode 
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
         Left            =   6240
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txt_Amount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtparticular 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   1560
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   795
         Width           =   9135
      End
      Begin VB.TextBox txtObrno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   435
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   19
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   14280
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Accountcode:"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular:"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "OBR No.:"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   8160
      Top             =   120
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
            Picture         =   "frm_AccountsPayable.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":2686
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":4018
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":59AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":733C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":8CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":A660
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":BFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":D984
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":F318
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":FFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":108D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":115B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":1228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":12F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":13C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountsPayable.frx":14920
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H lvlreset 
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Reset"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14737632
      cGradient       =   14737632
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_AccountsPayable.frx":151FC
      ImgSize         =   24
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   1482
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   14
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete all"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Criteria"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10695
      Begin lvButton.lvButtons_H lvlOK 
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&OK"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14737632
         cGradient       =   14737632
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountsPayable.frx":18D06
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   143654915
         UpDown          =   -1  'True
         CurrentDate     =   41307
      End
      Begin VB.ComboBox Cmb_fundtype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fundtype:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_AccountsPayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call LoadFundType(Cmb_fundtype)
End Sub

Private Sub GettotalAP()
On Error Resume Next
Dim x As Long
Text2.Text = "0.00"
For x = 1 To MSHFlexGrid1.Rows - 1
    Text2.Text = CCur(Text2.Text) + CCur(MSHFlexGrid1.TextMatrix(x, 3))
Next x
Text2.Text = Format(Text2.Text, "#,##0.00")
End Sub

Private Sub lvlOK_Click()
lvlOK.Enabled = False
lvlreset.Enabled = True
Frame3.Enabled = False
Frame2.Enabled = True
Toolbar1.Enabled = True
Call Loaddata
End Sub

Private Sub lvlreset_Click()
lvlOK.Enabled = True
Frame3.Enabled = True
lvlreset.Enabled = False
Frame2.Enabled = False
Toolbar1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button:
Case "Import"
With frm_AP_Import
    .fundcode = Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex)
    .YEAR_ = DTPicker1.Year
    .Show 1
    Call Loaddata
End With
Case "Save"
MsgBox "Save"
Case "Delete all"
    If MsgBox("Are You Sure do you want to Delete all List Below?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Delete FROM [fmis].[dbo].[tblAMIS_Accounts_Payable] where fundcode = " & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & " and year_ = " & DTPicker1.Year & ""
        Call Loaddata
    End If
End Select
End Sub
Public Sub Loaddata()
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("SELECT [OBRNO],[Particulars],Convert(nchar,[Amount],101) as Amount,[MainAccountcode],[SubAccountcode] FROM [fmis].[dbo].[tblAMIS_Accounts_Payable] where fundcode = " & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & " and year_ = " & DTPicker1.Year & "")
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
If rec.RecordCount > 0 Then
    Set MSHFlexGrid1.DataSource = rec
End If
Call SetMSHGrid(MSHFlexGrid1, 5)
Call GettotalAP
rec.Close
Set rec = Nothing
End Sub
