VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmChkAuthorizedUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Authorization Registry"
   ClientHeight    =   3615
   ClientLeft      =   5340
   ClientTop       =   3030
   ClientWidth     =   7050
   Icon            =   "frmChkAuthorizedUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5070
      TabIndex        =   6
      Top             =   990
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   -105
      TabIndex        =   1
      Top             =   840
      Width           =   4830
      Begin VB.TextBox txt_code 
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
         IMEMode         =   3  'DISABLE
         Left            =   1170
         PasswordChar    =   "@"
         TabIndex        =   4
         Top             =   1455
         Width           =   3060
      End
      Begin VB.TextBox txt_userid 
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
         Left            =   1170
         TabIndex        =   2
         Top             =   645
         Width           =   3060
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Authority Code"
         Height          =   405
         Left            =   390
         TabIndex        =   5
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   735
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   1058
      ButtonWidth     =   953
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   360
      Top             =   2715
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
End
Attribute VB_Name = "frmChkAuthorizedUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveflag As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
WindowsXPC1.InitIDESubClassing
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Call LoadAllUser
End Sub

Private Function IfExist(ByVal UserID As String) As Boolean
Dim opnuser As New ADODB.Recordset

opnuser.Open "Select * from tblAMIS_UserAdvance where userid='" & txt_userid.Text & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnuser.RecordCount <> 0 Then
    IfExist = True
Else
    IfExist = False
End If
opnuser.Close
Set opnuser = Nothing


End Function
Private Sub Save()
If saveflag = 1 Then 'New
    If IfExist(txt_userid.Text) = False Then
        opndbaseFMIS.Execute "Insert into tblAMIS_UserAdvance (UserID,pword,actioncode) values('" & txt_userid.Text & "','" & mydll.Encrypt(txt_code.Text) & "',1)"
        MsgBox "Saving new Account, Successful!", vbInformation, "System Information"
    Else
        MsgBox "UserID already in used!", vbInformation, "System Information"
    End If
ElseIf saveflag = 2 Then 'Edit
    opndbaseFMIS.Execute "Update tblAMIS_UserAdvance set UserID='" & txt_userid.Text & "',pword='" & mydll.Encrypt(txt_code.Text) & "' where trnno=" & List1.ItemData(List1.ListIndex) & ""
    MsgBox "Updating Account, Successful!", vbInformation, "System Information"
End If
Call LoadAllUser
Call Clear
Frame1.Enabled = False
End Sub
Private Sub LoadAllUser()
Dim opnuser As New ADODB.Recordset
Dim xx As Integer

List1.Clear
opnuser.Open "Select * from tblAMIS_UserAdvance where actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnuser.RecordCount <> 0 Then
    Do Until opnuser.EOF
        List1.AddItem (opnuser!UserID)
        List1.ItemData(xx) = opnuser!Trnno
        xx = xx + 1
    opnuser.MoveNext
    Loop
End If
opnuser.Close
Set opnuser = Nothing

End Sub
Private Sub LoadBack(ByVal UserID As String)
Dim opnuser As New ADODB.Recordset

opnuser.Open "Select * from tblAMIS_UserAdvance where actioncode=1 and userid='" & UserID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnuser.RecordCount <> 0 Then
    txt_userid.Text = opnuser!UserID
    txt_code.Text = mydll.Decrypt(opnuser!Pword)
Else
    Call Clear
End If
opnuser.Close
Set opnuser = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
WindowsXPC1.EndWinXPCSubClassing
Set frmChkAuthorizedUser = Nothing
End Sub
Private Sub Clear()
txt_userid.Text = ""
txt_code = ""
End Sub

Private Sub List1_Click()
Call LoadBack(List1.List(List1.ListIndex))
saveflag = 0
Frame1.Enabled = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 2 Then
    If Len(List1.Text) > 0 Then
        If MsgBox("Are you sure want to Delete this User?", vbQuestion + vbYesNo, "system confirmation") = vbYes Then
            opndbaseFMIS.Execute "Update tblAMIS_UserAdvance set actioncode=4 where trnno=" & List1.ItemData(List1.ListIndex) & ""
            Call LoadAllUser
            Call Clear
            Frame1.Enabled = False

        End If
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        saveflag = 1
        Frame1.Enabled = True
        Call Clear
    Case 3
        Call Save
        saveflag = 0
        
    Case 5
        saveflag = 2
        Frame1.Enabled = True
    Case 7
        Unload Me
End Select
End Sub
