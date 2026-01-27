VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_UserDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_UserDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10350
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   1050
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Change Pic"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8438015
      cGradient       =   8438015
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_UserDetails.frx":0E42
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Password"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtUserID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "Password"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   4095
         Begin VB.TextBox txt_RnewPW 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txt_NewPW 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox txt_OldPW 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   3960
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Retype New Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Note: To change your password, please type your old password first. Then enter a new password twice."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "New Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Old Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPosition 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1482
      ButtonWidth     =   1191
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   3720
         Top             =   0
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
               Picture         =   "frm_UserDetails.frx":20C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":3A56
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":53E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":6D7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":870C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":A09E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":BA30
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":D3C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":ED54
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":106E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":113C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":11CA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":12980
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":1365C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":14338
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":15014
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_UserDetails.frx":15CF0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:Please Complete the Required field."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblpath 
      BackStyle       =   0  'Transparent
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
      Left            =   4560
      TabIndex        =   17
      Top             =   960
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   6435
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   5685
   End
End
Attribute VB_Name = "frm_UserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isChange As Boolean
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Frame2.Enabled = True
End If
End Sub

Private Sub Form_Load()
On Error GoTo bad
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim strSQL As String
    
    strSQL = "SELECT [UserID],[UserName],[UserPassword],[Position],[Pic],admin FROM [fmis].[dbo].[tblAMIS_UserRegistry] where userid = '" & Trim(ActiveUserID) & "' and actioncode =1"
                
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = opndbaseFMIS
        .Source = strSQL
        .Open
    End With
    
    If Not (rs.BOF And rs.EOF) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
        txtUserID.Text = rs!UserID
        txtname.Text = rs!UserName
        txtposition.Text = IIf(IsNull(rs!Position), "", rs!Position)
        admin = rs!admin
        If IsNull(rs!pic) = False Then
        strStream.Write rs!pic
        strStream.SaveToFile App.path & "\img.bmp", adSaveCreateOverWrite
        Image1.Picture = LoadPicture(App.path & "\img.bmp")
        strStream.Close
        End If
        Set strStream = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
bad:
    MsgBox "Noted: " & err.description, vbCritical, "System Message"
    
End Sub

Private Sub lvButtons_H1_Click()
On Error GoTo bad
With CommonDialog1
    .CancelError = False
    .Filter = "Image Files(*.jpg;*.jpeg)"
    'Set .FileTitle = "*.jpg"
    .ShowOpen
    .DialogTitle = "Browse Picture"
    
    If .FileName <> "" Then
        Image1.Picture = LoadPicture(.FileName)
        lblpath.Caption = "Path: " & .FileName
        isChange = True
    End If
End With
Exit Sub
bad:
MsgBox err.description
End Sub
Public Function SaveImageDB()
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim Trnno  As Long
Dim pw As String
    'Add the image to the database
    
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    If CommonDialog1.FileName <> "" Then
        strStream.LoadFromFile CommonDialog1.FileName
    Else
        strStream.LoadFromFile App.path & "\img.bmp"
    End If
    
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = opndbaseFMIS
        .Source = "SELECT [UserID],[UserName],[UserPassword],[Position],[Pic],datetimeentered,admin,actioncode,trnno FROM [fmis].[dbo].[tblAMIS_UserRegistry] where userid = '" & Trim(ActiveUserID) & "' and actioncode = 1"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    If rs.RecordCount > 0 Then
    pw = Trim(rs!userPassword)
    trrno = rs!Trnno
    
    If Check1.Value = 1 Then
        If mydll.Encrypt(txt_OldPW.Text) <> Trim(rs!userPassword) Then
            MsgBox "Invalid old Password...!", vbCritical, "System Message"
        Else
            If Trim(txt_NewPW.Text) <> Trim(txt_RnewPW.Text) Then
                MsgBox "Password Did not Match...!", vbCritical, "System Message"
            Else
                If MsgBox("Are you sure do you want to upload the Photo?", vbInformation + vbYesNo, "System Message") = vbYes Then
                    rs.AddNew
                    rs.Fields("UserId").Value = txtUserID.Text
                    rs.Fields("userName").Value = txtname.Text
                    rs.Fields("userpassword").Value = mydll.Encrypt(txt_NewPW.Text)
                    rs.Fields("Position").Value = txtposition.Text
                    rs.Fields("datetimeentered").Value = Now
                    rs.Fields("pic").Value = strStream.Read
                    rs.Fields("admin").Value = admin
                    rs.Fields("actioncode").Value = 1
                    rs.update
                    MsgBox "Successfully Updated in Database."
                    txt_OldPW.Text = ""
                    txt_NewPW.Text = ""
                    txt_RnewPW.Text = ""
                    Check1.Value = 0
                    Frame2.Enabled = False
                    opndbaseFMIS.Execute "Update [tblAMIS_UserRegistry] set actioncode =2 where trnno <= " & trrno & " and userid = '" & ActiveUserID & "' and actioncode = 1"
                    frmLogInPicture.img_userpic.Visible = True
                    Call LoadImageUser
                    frmLogInPicture.img_userpic.Picture = LoadPicture(App.path & "\img.bmp")
                End If
            End If
        End If
    Else
    If MsgBox("Are you sure do you want to upload the Photo?", vbInformation + vbYesNo, "System Message") = vbYes Then
        rs.AddNew
        rs.Fields("UserId").Value = txtUserID.Text
        rs.Fields("userName").Value = txtname.Text
        rs.Fields("Position").Value = txtposition.Text
        rs.Fields("userpassword").Value = pw
        rs.Fields("datetimeentered").Value = Now
        rs.Fields("pic").Value = strStream.Read
        rs.Fields("admin").Value = admin
        rs.Fields("actioncode").Value = 1
        rs.update
        MsgBox "Successfully Updated in Database."
        txt_OldPW.Text = ""
        txt_NewPW.Text = ""
        txt_RnewPW.Text = ""
        Check1.Value = 0
        Frame2.Enabled = False
        opndbaseFMIS.Execute "Update [tblAMIS_UserRegistry] set actioncode =2 where trnno <= " & trrno & " and userid = '" & ActiveUserID & "' and actioncode = 1"
        frmLogInPicture.img_userpic.Visible = True
        Call LoadImageUser
        frmLogInPicture.img_userpic.Picture = LoadPicture(App.path & "\img.bmp")
    End If
    End If
    
    End If
    
    strStream.Close
    rs.Close
    'Cleanup
    Set strStream = Nothing
    Set rs = Nothing
    Set cn = Nothing
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'New
        
'            lblstatus.Visible = True
            'Call PlayAVI(Me.Animation1, "horizontaloading.avi")
            Call SaveImageDB
            'Call StopAvi(Me.Animation1)
            'lblstatus.Visible = False
    Case 5:
    Unload Me
End Select
End Sub
