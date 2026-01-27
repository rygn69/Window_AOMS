VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmUserPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log on "
   ClientHeight    =   3675
   ClientLeft      =   3945
   ClientTop       =   2040
   ClientWidth     =   6600
   Icon            =   "FrmUserPassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton FlatBttn2 
      Caption         =   "  &Cancel "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CommandButton FlatBttn1 
      Caption         =   " &Log in  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   5040
      Picture         =   "FrmUserPassword.frx":3AFA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1185
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -360
      Top             =   4320
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
      TextControl     =   0   'False
   End
   Begin VB.TextBox Username_txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1485
      TabIndex        =   2
      Top             =   2370
      Width           =   3435
   End
   Begin VB.TextBox password_txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1485
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2910
      Width           =   3435
   End
   Begin VB.TextBox SwipeIDNo_txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1485
      TabIndex        =   0
      Top             =   1830
      Width           =   3435
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2565
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   4230
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   210
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2010"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblProductName 
      BackStyle       =   0  'Transparent
      Caption         =   "Agusan del Sur                                  ccounting      peration                       anagement      ystem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1215
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your access parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1500
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   5
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   4
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1905
      Width           =   1200
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   6135
      Left            =   0
      Picture         =   "FrmUserPassword.frx":43C4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image Image2 
      Height          =   6165
      Left            =   0
      Picture         =   "FrmUserPassword.frx":3BBC6
      Stretch         =   -1  'True
      Top             =   -1920
      Width           =   7440
   End
End
Attribute VB_Name = "frmUserPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Public whatLog As Integer
Dim Ok As Boolean

Private Sub FlatBttn1_Click()
On Error GoTo bad
'counter = 0
counter = counter + 1
FlatBttn1.Enabled = False
If counter <= 3 Then
    FlatBttn1.Enabled = False

    Call VerifyUser
    
    FlatBttn1.Enabled = True
Else
    MsgBox "Sorry, but you have no Access to the System!", vbInformation, "System Information"
    opndbasePMIS.Close
    opndbaseFMIS.Close
    Set opndbasePMIS = Nothing
    Set opndbaseFMIS = Nothing
    Unload Me
    End
End If
FlatBttn1.Enabled = True
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub FlatBttn2_Click()
If ShutDownMode = "Open System" Then
    opndbasePMIS.Close
    opndbaseFMIS.Close
    Set opndbasePMIS = Nothing
    Set opndbaseFMIS = Nothing
End If
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call FlatBttn2_Click
End If
End Sub

Private Sub Form_Load()

WindowsXPC1.InitSubClassing

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
'Timer1.Enabled = True
End Sub
Private Sub LoadUser()
Dim opnuser As New ADODB.Recordset
Dim opnuser1 As New ADODB.Recordset

    opnuser.Open "Select * from pmis.dbo.Employee where SwipEmployeeID='" & Replace(SwipeIDNo_txt.Text, "'", "") & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
   
        
        If opnuser.RecordCount <> 0 Then
            Username_txt.Text = UCase(opnuser!Firstname & " " & IIf(Len(Trim(opnuser!MI)) = 0, "", Left(opnuser!MI, 1) & ".") & " " & opnuser!Lastname & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", "," & opnuser!Suffix))
            
            If Len(Trim(password_txt.Text)) <> 0 Then '--On got fucos on password_txt ----
                password_txt.SelStart = 0
                password_txt.SelLength = Len(password_txt.Text)
                password_txt.SetFocus
            Else
                password_txt.SetFocus
            End If '-----------------------------------------------------------------
        
        Else
            opnuser1.Open "Select * from tblAMIS_UserRegistry where USERID='" & Replace(SwipeIDNo_txt.Text, "'", "") & "' AND ACTIONCODE = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
            If opnuser1.RecordCount > 0 Then
                Username_txt.Text = Trim(opnuser1!UserName)
                If Len(Trim(password_txt.Text)) <> 0 Then '--On got fucos on password_txt ----
                    password_txt.SelStart = 0
                    password_txt.SelLength = Len(password_txt.Text)
                    password_txt.SetFocus
                Else
                    password_txt.SetFocus
                End If
            Else
                MsgBox "User ID No. you have currently entered is not registered in the PMIS!", vbInformation, "System Information"
                If Len(Trim(SwipeIDNo_txt.Text)) <> 0 Then
                    SwipeIDNo_txt.SelStart = 0
                    SwipeIDNo_txt.SelLength = Len(SwipeIDNo_txt.Text)
                    SwipeIDNo_txt.SetFocus
                Else
                    SwipeIDNo_txt.SetFocus
                End If
            End If
        End If
    opnuser.Close
    Set opnuser = Nothing

End Sub
Private Sub LoadUserPic()
Dim opnuserpic As New ADODB.Recordset


On Error GoTo handler

'loading the computername where pictures is found
PicLocation = readTXTDATA("UserPictureLocation", "location", App.path & "\data\SystemDefault.ini")

opnuserpic.Open "Select photo from pmis.dbo.employee where swipemployeeid='" & SwipeIDNo_txt.Text & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnuserpic.RecordCount <> 0 Then
    frmLogInPicture.img_userpic.Visible = True
    Call LoadImageUser
    frmLogInPicture.img_userpic.Picture = LoadPicture(App.path & "\img.bmp")
Else
    frmLogInPicture.img_userpic.Visible = True
    Call LoadImageUser
    frmLogInPicture.img_userpic.Picture = LoadPicture(App.path & "\img.bmp")
End If
opnuserpic.Close
Set opnuserpic = Nothing

handler:
If err.Number <> 0 Then
    frmLogInPicture.img_userpic.Visible = False
    opnuserpic.Close
    Set opnuserpic = Nothing
End If
End Sub


Private Sub VerifyUser()
On Error Resume Next
Dim opnVerify As New ADODB.Recordset
opnVerify.Open "Select * from tblAMIS_UserRegistry where UserPassword='" & mydll.Encrypt(UCase(password_txt.Text)) & "'  and userid='" & SwipeIDNo_txt.Text & "' and Actioncode=1 ", opndbaseFMIS, adOpenStatic, adLockOptimistic
    
    If opnVerify.RecordCount <> 0 Then
    'SETTING ACTIVE PARAMETERS---------------------------------\\
        
        ActiveUser = UCase(opnVerify!UserName)
        ActiveUserID = opnVerify!UserID
        ActiveUserPass = password_txt.Text
        Call TransactionLogging("Log IN", "tblAMIS_log", Me.Caption, Winsock1.LocalIP)
        Call OnlineLogging(ActiveUserID, Winsock1.LocalIP, Winsock1.LocalPort)
        'Verify User Level-----------------------------------------
        frmLogInPicture.Label1.Caption = Trim(ActiveUser)
        frmLogInPicture.Label2.Caption = Trim(opnVerify!Position)
        '----Loading the active user picture---------
        Ok = True
        Call LoadUserPic
        Log = "In"
        
        'Writing to Log In History ------------------------
            
         
        Unload Me 'unloading this form
        '-----------------------------------
            Call frm_toolwindows.loadUsernode
            frm_toolwindows.lvButtons_H3.Caption = "Log-out"
            MDIFrm_MAIN.Show
        '-----------------------------------
    Else
        MsgBox "Please correct your entries!", vbCritical, "System Warning!"
        If Len(Trim(password_txt)) <> 0 Then
            password_txt.SelStart = 0
            password_txt.SelLength = Len(password_txt.Text)
            password_txt.SetFocus
        Else
            password_txt.SetFocus
        End If
        'frmLogInPicture.img_userpic.Visible = False
    End If
opnVerify.Close
Set opnVerify = Nothing
Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If Ok = True Then
'        If whatLog = 0 Then
'        Cancel = 1
'        ElseIf whatLog = 2 Then
'        Unload Me
'        End If
'    Else
'    End
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUserPassword = Nothing
    WindowsXPC1.EndWinXPCSubClassing
End Sub

Private Sub password_txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call FlatBttn1_Click
End If
End Sub

Private Sub SwipeIDNo_txt_Click()
If Len(Trim(SwipeIDNo_txt)) <> 0 Then
    SwipeIDNo_txt.SelStart = 0
    SwipeIDNo_txt.SelLength = Len(SwipeIDNo_txt.Text)
    SwipeIDNo_txt.SetFocus
End If
End Sub

Private Sub SwipeIDNo_txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SwipeIDNo_txt.Text = UCase(SwipeIDNo_txt.Text)
    Call LoadUser
End If
End Sub
