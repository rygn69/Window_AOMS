VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AccountcodeSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accountcode and Explaination Classification Utility"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "frm_AccountcodeSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   8760
      TabIndex        =   0
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "&Cancel"
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
      Image           =   "frm_AccountcodeSub.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   4440
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
      Image           =   "frm_AccountcodeSub.frx":4274
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Add"
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
      Image           =   "frm_AccountcodeSub.frx":7D7E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "&Save"
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
      Image           =   "frm_AccountcodeSub.frx":B888
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   7215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accountcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accountname"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   6000
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
            Picture         =   "frm_AccountcodeSub.frx":F392
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":10D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":126B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":14048
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":159DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1736C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":18CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1A690
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1C022
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1D9B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1E692
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1EF72
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":1FC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":2092A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":21606
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":222E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":22FBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
   End
End
Attribute VB_Name = "frm_AccountcodeSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
End Sub
Public Function GetAccountNamebyorder(ByVal lst As ListView, ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim X
Dim z As Integer
'Condition = Replace(Condition, "'", "")
rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother order by " & Condition & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
    
        For z = 1 To rec.RecordCount

                    Set X = lst.ListItems.Add(, , rec.Fields!Accountcode)
                    X.SubItems(1) = Trim(rec.Fields!accountname)
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
End Function


Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader = "Accountname" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountname")
ElseIf ColumnHeader = "Accountcode" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
End If
End Sub

Private Sub LstAccountcode_DblClick()
With frm_AccountcodeSub1
.Address = Trim(LstAccountcode.SelectedItem.SubItems(1))
.col = 2
.Accountcode = Trim(LstAccountcode.SelectedItem.Text)
.Subcode1 = Trim(LstAccountcode.SelectedItem.Text)
.Subdesc1 = Trim(LstAccountcode.SelectedItem.SubItems(1))
.Condition = "Subcode1 ='" & Trim(LstAccountcode.SelectedItem.Text) & "'"
.Show
End With
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H4_Click()
Dim X As Integer
For X = 1 To LstAccountcode.ListItems.Count
opndbaseFMIS.Execute "Insert into tblReff_CodeClassification(subcode1,subdesc1) values('" & LstAccountcode.ListItems(X).Text & "','" & Replace(LstAccountcode.ListItems(X).ListSubItems(1).Text, "'", "") & "') "
Next X
MsgBox "Successfully save", vbInformation, "System Message"
End Sub


