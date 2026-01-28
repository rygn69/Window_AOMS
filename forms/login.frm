VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUserPassword 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   4815
   ClientLeft      =   3840
   ClientTop       =   3045
   ClientWidth     =   6300
   Icon            =   "login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":09EA
   ScaleHeight     =   4815
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   550
      Left            =   50
      ScaleHeight     =   555
      ScaleWidth      =   6195
      TabIndex        =   14
      Top             =   4180
      Width           =   6200
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"login.frx":185D
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   50
         TabIndex        =   15
         Top             =   30
         Width           =   6135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   5760
      Top             =   1440
   End
   Begin VB.PictureBox picCompany 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   240
      ScaleHeight     =   2055
      ScaleWidth      =   7680
      TabIndex        =   9
      Top             =   6360
      Width           =   7680
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   8400
      Picture         =   "login.frx":1925
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   3840
      Width           =   15
   End
   Begin VB.PictureBox log_pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   50
      Picture         =   "login.frx":1BA8
      ScaleHeight     =   2235
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   1900
      Width           =   6195
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1800
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox SwipeIDNo_txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   0
         Top             =   120
         Width           =   2985
      End
      Begin VB.TextBox password_txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2985
      End
      Begin VB.TextBox Username_txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   1
         Top             =   603
         Width           =   2985
      End
      Begin lvButton.lvButtons_H FlatBttn1 
         Height          =   615
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "&Log in"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648384
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "login.frx":1E2B
         ImgSize         =   32
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H FlatBttn2 
         Height          =   615
         Left            =   5520
         TabIndex        =   13
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12632319
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "login.frx":2705
         ImgSize         =   32
         cBack           =   -2147483633
      End
      Begin VB.Label blnkFullName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   6840
         TabIndex        =   11
         Top             =   120
         Width           =   2985
      End
      Begin VB.Label Username_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID NUMBER:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   240
         Index           =   1
         Left            =   1515
         TabIndex        =   7
         Top             =   165
         Width           =   1545
      End
      Begin VB.Label Password_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   345
         Left            =   1995
         TabIndex        =   6
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Username_lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   240
         Index           =   0
         Left            =   1995
         TabIndex        =   5
         Top             =   645
         Width           =   1065
      End
      Begin VB.Image Image2 
         Height          =   1995
         Left            =   0
         Picture         =   "login.frx":2FDF
         Top             =   0
         Width           =   1995
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   45
      Picture         =   "login.frx":5028
      Top             =   45
      Width           =   6180
   End
   Begin VB.Label progress_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "progress"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   285
      TabIndex        =   3
      Top             =   3675
      Width           =   1680
   End
   Begin VB.Image imgPicture 
      Height          =   15
      Index           =   0
      Left            =   120
      Picture         =   "login.frx":91EF
      Top             =   1080
      Width           =   15
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ePay"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   885
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   5640
      Width           =   3135
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
Dim rec As New ADODB.Recordset


Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
'Timer1.Enabled = True
'Call LoadSysteProperties
'Image3.Picture = LoadPicture(App.path & "\images\Login_Form.jpg")
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
'On Error Resume Next
On Error GoTo bad
Dim opnVerify As New ADODB.Recordset
opnVerify.Open "Select * from tblAMIS_UserRegistry where UserPassword='" & mydll.Encrypt(UCase(password_txt.Text)) & "'  and userid='" & SwipeIDNo_txt.Text & "' and Actioncode=1 ", opndbaseFMIS, adOpenStatic, adLockOptimistic
    
    If opnVerify.RecordCount <> 0 Then
    'SETTING ACTIVE PARAMETERS---------------------------------\\
        
        ActiveUser = UCase(opnVerify!UserName)
        ActiveUserID = opnVerify!UserID
        ActiveUserPass = password_txt.Text
        SystemAdmin = opnVerify!admin
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
bad:
MsgBox err.description
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
End Sub

Private Sub Password_txt_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub text1_Change()
Text2.Text = mydll.Decrypt(UCase(Text1.Text))
End Sub
