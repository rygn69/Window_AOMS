VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPCEngine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSystemUsers_new 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Users Restriction Level Registry"
   ClientHeight    =   9465
   ClientLeft      =   1890
   ClientTop       =   2430
   ClientWidth     =   13515
   Icon            =   "frm_Users_new.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   13515
   Begin VB.CheckBox Check1 
      Caption         =   "Administrator"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   2880
      Width           =   1575
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      ToolTipText     =   "Add Check Item"
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Users_new.frx":000C
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView lstuser 
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9763
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PW"
         Object.Width           =   0
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      FrameControl    =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4665
      Begin VB.TextBox txtRpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1335
         PasswordChar    =   "@"
         TabIndex        =   4
         Top             =   1640
         Width           =   3135
      End
      Begin VB.TextBox Txt_UserId 
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
         Left            =   1335
         TabIndex        =   1
         Top             =   345
         Width           =   3135
      End
      Begin VB.TextBox txt_Username 
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
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   765
         Width           =   3135
      End
      Begin VB.TextBox txt_Password 
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
         IMEMode         =   3  'DISABLE
         Left            =   1335
         PasswordChar    =   "@"
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retype-PW:"
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
         Left            =   90
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
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
         Left            =   90
         TabIndex        =   7
         Top             =   405
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Left            =   90
         TabIndex        =   6
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   90
         TabIndex        =   5
         Top             =   1275
         Width           =   945
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   7800
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
            Picture         =   "frm_Users_new.frx":3B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":54A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":6E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":87CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":A15E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":BAF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":D482
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":EE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":107A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":1213A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":12E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":136F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":143D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":150AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":15D8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":16A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Users_new.frx":17742
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   1508
      ButtonWidth     =   1138
      ButtonHeight    =   1455
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "New Account"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save the account"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit  Account Password"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete the Account"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close the form"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstfullnode 
      Height          =   2295
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "System Nodes"
         Object.Width           =   19403
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subcode1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subcode2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subcode3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "lvl"
         Object.Width           =   0
      EndProperty
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      ToolTipText     =   "Remove Check Item"
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   4
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
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Users_new.frx":1801E
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView lstUsernode 
      Height          =   5535
      Left            =   4920
      TabIndex        =   14
      Top             =   3840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User's Nodes"
         Object.Width           =   19403
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subcode1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subcode2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subcode3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "lvl"
         Object.Width           =   0
      EndProperty
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      ToolTipText     =   "Remove Check Item"
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Refresh"
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
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_Users_new.frx":1BB28
      cBack           =   -2147483633
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   1065
   End
End
Attribute VB_Name = "frmSystemUsers_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isNew As Boolean
Public isAdmin As Integer
Private Sub Check1_Click()
Dim val As Integer
If txt_userid.Text <> "" Then
    If Check1.Tag = "edit" Then
        If SystemAdmin = 1 Then
            If Check1.Value = 1 Then
                If MsgBox("Are you sure you want to make as ADMINISTRATOR this user?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                    opndbaseFMIS.Execute ("update tblAMIS_UserRegistry set admin = 1 where UserID = '" & txt_userid.Text & "' and actioncode = 1")
                Else
                Check1.Tag = "a"
                    Check1.Value = isAdmin
                End If
            Else
                If MsgBox("Are you sure you want to make as SIMPLE USER?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                    opndbaseFMIS.Execute ("update tblAMIS_UserRegistry set admin = 0 where UserID = '" & txt_userid.Text & "' and actioncode = 1")
                Else
                    Check1.Tag = "a"
                    Check1.Value = isAdmin
                End If
            End If
        Else
        
            Check1.Value = isAdmin
            MsgBox "Anauthorized Access", vbInformation, "System Message"
        End If
    Else
        Check1.Value = isAdmin
    End If
Else
    Check1.Value = isAdmin
End If
End Sub

Private Sub Form_Load()
LoadUser
End Sub
Private Sub lstuser_Click()
'On Error Resume Next
isNew = False
Frame1.Enabled = False
txt_userid.Text = lstuser.SelectedItem.Text
txt_Username.Text = lstuser.SelectedItem.ListSubItems(1).Text
txt_Password.Text = mydll.Decrypt(Trim(lstuser.SelectedItem.ListSubItems(3).Text))
txtRpassword.Text = mydll.Decrypt(Trim(lstuser.SelectedItem.ListSubItems(3).Text))
Check1.Tag = "a"
    If lstuser.SelectedItem.ListSubItems(2).Text = 1 Then
        Check1.Value = 1
        isAdmin = 1
    Else
        Check1.Value = 0
        isAdmin = 0
    End If
Check1.Tag = "edit"
Call LoadFullnode
Call LoadUserFullnode
End Sub

Private Sub lstuser_KeyUp(KeyCode As Integer, Shift As Integer)
Call lstuser_Click
End Sub

Private Sub lvButtons_H1_Click()
Dim x As Long
Dim y As Long
Dim CheckIFAlreadyInUserLst As Boolean
If MsgBox("Are you sure do you want to ADD the selected nodes?", vbInformation + vbYesNo, "Syste Message") = vbYes Then
    CheckIFAlreadyInUserLst = False
    For x = 1 To lstfullnode.ListItems.Count
        If lstfullnode.ListItems(x).Checked = True Then
        opndbaseFMIS.Execute "Insert into  [fmis].[dbo].[tblAMIS_Usernodes]([UserID],[SubnodeID]) values ('" & txt_userid.Text & "','" & lstfullnode.ListItems(x).ListSubItems(4).Text & "')"
        End If
    Next x
    opndbaseFMIS.Execute "EXECUTE  [fmis].[dbo].[MPfunc_LoadfullnodeByUserID] @userID = '" & txt_userid.Text & "'"
    Call LoadFullnode
    Call LoadUserFullnode
End If
End Sub
Private Function SaveInNOde()
Dim x As Long
Dim y As Long
opndbaseFMIS.Execute "delete from  [fmis].[dbo].[tblAMIS_Usernodes] where [UserID] ='" & txt_userid.Text & "'"
For x = 1 To lstfullnode.ListItems.Count
    opndbaseFMIS.Execute "Insert into  [fmis].[dbo].[tblAMIS_Usernodes]([UserID],[SubnodeID]) values ('" & txt_userid.Text & "','" & lstfullnode.ListItems(x).ListSubItems(4).Text & "')"
    DoEvents
Next x
End Function

Private Sub lvButtons_H2_Click()
Dim x As Long
Dim y
If MsgBox("Are you sure do you want to Delete the selected nodes?", vbInformation + vbYesNo, "Syste Message") = vbYes Then
    For x = 1 To lstUsernode.ListItems.Count
        If lstUsernode.ListItems(x).Checked = True Then
            opndbaseFMIS.Execute "Delete from  tblAMIS_Usernodes where userid = '" & txt_userid.Text & "' and subnodeID = '" & lstUsernode.ListItems(x).ListSubItems(4).Text & "'"
        End If
    Next x
    opndbaseFMIS.Execute "EXECUTE  [fmis].[dbo].[MPfunc_LoadfullnodeByUserID] @userID = '" & txt_userid.Text & "'"
    LoadUserFullnode
    LoadFullnode
End If
End Sub
Private Function CheckIFAlreadyInUserLst() As Boolean

End Function

Private Sub lvButtons_H3_Click()
opndbaseFMIS.Execute "exec fmis.dbo.MPfunc_Loadfullnode"
    Call LoadFullnode
    Call LoadUserFullnode
End Sub

'Private Sub lvButtons_H1_Click()
'Dim x As Long
'Dim y
'For x = 1 To lstfullnode.ListItems.Count
'    If lstfullnode.ListItems(x).Checked = True Then
'        Set y = lstUsernode.ListItems.Add(, , lstfullnode.ListItems(x).Text)
'                y.SubItems(1) = lstfullnode.ListItems(x).ListSubItems(1).Text
'                y.SubItems(2) = lstfullnode.ListItems(x).ListSubItems(2).Text
'                y.SubItems(3) = lstfullnode.ListItems(x).ListSubItems(3).Text
'                y.SubItems(4) = lstfullnode.ListItems(x).ListSubItems(4).Text
'                'lstfullnode.ListItems.Remove (lstfullnode.ListItems(x).Index)
'    End If
'Next x
'End Sub
'
'Private Sub lvButtons_H2_Click()
'Dim x As Long
'Dim y
'For x = 1 To lstUsernode.ListItems.Count
'    If lstUsernode.ListItems(x).Checked = True Then
'        Set y = lstfullnode.ListItems.Add(, , lstfullnode.ListItems(x).Text)
'                y.SubItems(1) = lstfullnode.ListItems(x).ListSubItems(1).Text
'                y.SubItems(2) = lstfullnode.ListItems(x).ListSubItems(2).Text
'                y.SubItems(3) = lstfullnode.ListItems(x).ListSubItems(3).Text
'                y.SubItems(4) = lstfullnode.ListItems(x).ListSubItems(4).Text
'                lstfullnode.ListItems.Remove (lstfullnode.ListItems(x).Index)
'    End If
'Next x
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Integer
Dim rec As New ADODB.Recordset
Dim Inpot
Dim admin As Integer
Check1.Tag = "a"
If Check1.Value = 1 Then
admin = 1
Else
admin = 0
End If
Check1.Tag = "edit"
'On Error GoTo bad
    Select Case Button:
    Case "New":
        txtclear
        txt_userid.Enabled = True
        Frame1.Enabled = True
        isNew = True
    Case "Save":
        If isNew = True Then
        
        Set rec = opndbaseFMIS.Execute("Select * from tblAMIS_UserRegistry where userid = '" & txt_userid.Text & "' and actioncode = 1")
                If rec.RecordCount > 0 Then
                MsgBox "User ID Already in the database", vbInformation, "System Message"
                Else
                    If MsgBox("Are you sure do you want to save?", vbInformation + vbYesNo, "System Message") = vbYes Then
                        If txt_Password.Text = txtRpassword Then
                            If Trim(txt_userid.Text) <> "" And Trim(txt_Password.Text) <> "" And Trim(txt_Username.Text) <> "" Then
                            opndbaseFMIS.Execute "Insert into tblAMIS_UserRegistry ([UserID],[UserName],[UserPassword],[Actioncode],[enteredbyUserid],[DateTimeEntered],admin) " & _
                            "values ('" & txt_userid.Text & "','" & Trim(txt_Username.Text) & "','" & mydll.Encrypt(UCase(txt_Password.Text)) & "',1,'" & ActiveUserID & "','" & Now & "','" & admin & "')"
                            LoadUser
                            Call LoadFullnode
                            Check1.Tag = "a"
                            If Check1.Value = 1 Then
                                SaveInNOde
                                LoadUserFullnode
                                opndbaseFMIS.Execute "EXECUTE  [fmis].[dbo].[MPfunc_LoadfullnodeByUserID] @userID = '" & txt_userid.Text & "'"
                            Else
                            MsgBox "Successfully save, Please assign to her/his nodes..", vbInformation, "System Message"
                            End If
                            Frame1.Enabled = False
                            Else
                                MsgBox "Please Complete the required field", vbInformation, "System Messagse"
                            End If
                            Check1.Tag = "edit"
                        Else
                        MsgBox "Password did not match, Please Check it", vbInformation, "System Message"
                        End If
                    End If
                End If
        Else
                    If MsgBox("Are you sure do you want to Update the password?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                    opndbaseFMIS.Execute "Update tblAMIS_UserRegistry set userpassword = '" & mydll.Encrypt(UCase(txt_Password.Text)) & "' where userid = '" & Trim(txt_userid) & "' and actioncode =1"
                    MsgBox "Successfully Update", vbInformation, "System Message"
                    End If
        End If
    Case "Edit":
    If isNew = False Then
ret:
             Inpot = InputBox("ENTER User Password:", "Password Verification")
                If UCase(Inpot) = txt_Password.Text Then
                     Frame1.Enabled = True
                     txt_userid.Enabled = False
                Else
                    If MsgBox("Invalid password", vbCritical + vbRetryCancel, "System Messagse") = vbRetry Then
                        GoTo ret
                    End If
                End If

    End If
    Case "Delete":
                If isNew = False Then
                    If MsgBox("Are you sure do you want to delete this account?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                        opndbaseFMIS.Execute "Update tblAMIS_UserRegistry set actioncode  =2 where userid = '" & txt_userid.Text & "' and actioncode = 1"
                        opndbaseFMIS.Execute "delete from  tblAMIS_Usernodes  where userid = '" & txt_userid.Text & "'"
                        txtclear
                        LoadUser
                    End If
                End If
    Case "Close":
                If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                    Unload Me
                End If
    End Select
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Function txtclear()
txt_Password.Text = ""
txt_userid.Text = ""
txt_Username.Text = ""
txtRpassword.Text = ""
lstfullnode.ListItems.Clear
lstUsernode.ListItems.Clear
End Function
Public Sub LoadFullnode()
Dim x As Long
Dim y
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("SELECT  [ID],[Fullnodes],[Subnode1],[Subnode2],[Subnode3],[lvl]  FROM [fmis].[dbo].[tblAMIS_tempNodes] where id not in (SELECT [SubnodeID] FROM [fmis].[dbo].[tblAMIS_Usernodes] where userID ='" & txt_userid.Text & "') order by subnode1,subnode2")
lstfullnode.ListItems.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        Set y = lstfullnode.ListItems.Add(, , Trim(IIf(IsNull(rec!Fullnodes), "", rec!Fullnodes)))
            y.SubItems(1) = Trim(IIf(IsNull(rec!Subnode1), "", rec!Subnode1))
            y.SubItems(2) = Trim(IIf(IsNull(rec!Subnode2), "", rec!Subnode2))
            y.SubItems(3) = IIf(IsNull(Trim(rec!Subnode3)), "", rec!Subnode3)
            y.SubItems(4) = Trim(rec!id)
            rec.MoveNext
            DoEvents
    Next x
End If
End Sub
Public Sub LoadUserFullnode()
Dim x As Long
Dim y
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("SELECT  [ID],[Fullnodes],[Subnode1],[Subnode2],[Subnode3],[lvl]  FROM [fmis].[dbo].[tblAMIS_tempNodes] where id in (SELECT [SubnodeID] FROM [fmis].[dbo].[tblAMIS_Usernodes] where userID ='" & txt_userid.Text & "') order by id")
lstUsernode.ListItems.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        Set y = lstUsernode.ListItems.Add(, , Trim(rec!Fullnodes))
            y.SubItems(1) = Trim(rec!Subnode1)
            y.SubItems(2) = Trim(rec!Subnode2)
            y.SubItems(3) = IIf(IsNull(Trim(rec!Subnode3)), "", rec!Subnode3)
            y.SubItems(4) = Trim(rec!id)
            rec.MoveNext
            DoEvents
    Next x
End If
End Sub
Public Sub LoadUser()
Dim x As Long
Dim y
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select userID,UserName,admin,userpassword from tblAMIS_UserRegistry where actioncode = 1 group by userid,Username,userpassword,admin order by username")
lstuser.ListItems.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        Set y = lstuser.ListItems.Add(, , Trim(rec!UserID))
            y.SubItems(1) = Trim(rec!UserName)
            y.SubItems(2) = Trim(rec!admin)
            y.SubItems(3) = Trim(rec!userPassword)
            rec.MoveNext
            DoEvents
    Next x
End If
End Sub

Private Sub LoadUserIntxt()
Dim opnuser As New ADODB.Recordset

opnuser.Open "Select * from pmis.dbo.Employee where SwipEmployeeID='" & txt_userid.Text & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    
    If opnuser.RecordCount <> 0 Then
        txt_Username.Text = UCase(opnuser!Firstname & " " & IIf(Len(Trim(opnuser!MI)) = 0, "", Left(opnuser!MI, 1) & ".") & " " & opnuser!Lastname & " " & IIf(Len(Trim(IIf(IsNull(opnuser!Suffix), "", opnuser!Suffix))) = 0, "", "," & opnuser!Suffix))
        
        If Len(Trim(txt_Password.Text)) <> 0 Then '--On got fucos on password_txt ----
            txt_Password.SelStart = 0
            txt_Password.SelLength = Len(txt_Password.Text)
            txt_Password.SetFocus
        Else
            txt_Password.SetFocus
        End If '-----------------------------------------------------------------
    Else
        MsgBox "User ID No. you have currently entered is not registered in the PMIS!", vbInformation, "System Information"
    End If
opnuser.Close
Set opnuser = Nothing
End Sub

Private Sub Txt_UserId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LoadUserIntxt
End If
End Sub
