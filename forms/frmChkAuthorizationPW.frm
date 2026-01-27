VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmChkAuthorizationPW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Enter your Authorization Password"
   ClientHeight    =   2205
   ClientLeft      =   6180
   ClientTop       =   4620
   ClientWidth     =   5835
   Icon            =   "frmChkAuthorizationPW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
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
      Left            =   255
      PasswordChar    =   "@"
      TabIndex        =   4
      Top             =   855
      Width           =   3615
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   360
      Top             =   -210
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   4410
      TabIndex        =   2
      Top             =   735
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   420
      Left            =   4410
      TabIndex        =   1
      Top             =   255
      Width           =   1095
   End
   Begin VB.TextBox txt_pword 
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
      Left            =   255
      PasswordChar    =   "@"
      TabIndex        =   0
      Top             =   1590
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      Height          =   195
      Left            =   255
      TabIndex        =   5
      Top             =   585
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   1320
      Width           =   780
   End
End
Attribute VB_Name = "frmChkAuthorizationPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Private Sub Command1_Click()

counter = counter + 1
If counter <= 3 Then
    If Len(Trim(txt_pword.Text)) <> 0 Then
        If ActiveUserID = "0169" Then
            If ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
                frmEXVerifyCashAvailabilityNew.Timer4.Enabled = False
            End If
            
            counter = 0
            Unload Me
            frmExCashAllocator.Show vbModal
        Else
                If AuthorizedBKey(ActiveUserID, UCase(txt_pword.Text)) = True Then 'Authorized
                    If ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
                        frmEXVerifyCashAvailabilityNew.Timer4.Enabled = False
                    End If
                    counter = 0
                    Unload Me
                    frmExCashAllocator.Show vbModal
                Else
                    If MsgBox("Unauthorized Password!" & Chr(13) & Chr(13) & "Want to try again?", vbCritical + vbYesNo, "System Warning") = vbYes Then
                       txt_pword.SelStart = 0
                       txt_pword.SelLength = Len(txt_pword.Text)
                       txt_pword.SetFocus
                    Else
                        Unload Me
                    End If
                End If
        End If
    Else
        MsgBox "Please enter your password!", vbInformation, "System Information"
        txt_pword.SetFocus
    End If
Else
    MsgBox "You are not authorized for this procedure!", vbCritical, "System Warning"
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
WindowsXPC1.InitIDESubClassing
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
WindowsXPC1.EndWinXPCSubClassing
If ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
    frmEXVerifyCashAvailabilityNew.Timer4.Enabled = True
End If
Set frmExAuthorizationPW = Nothing
End Sub
Private Sub txt_pword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1_Click
End If
End Sub
