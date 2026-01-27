VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6270
   ClientLeft      =   2805
   ClientTop       =   3225
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   6270
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3360
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   4920
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   2865
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Junior Programmer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MAR PAUL M. AJERO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   3525
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6615
      TabIndex        =   8
      Top             =   4815
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Agusan del Sur"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   465
      Left            =   180
      TabIndex        =   7
      Top             =   105
      Width           =   2325
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   195
      TabIndex        =   1
      Top             =   2925
      Width           =   5115
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Provincial Information Management Office(PIMO)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   225
      TabIndex        =   6
      Top             =   5760
      Width           =   5190
   End
   Begin VB.Label Label2 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincial Government of Agusan del Sur"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   6045
      TabIndex        =   4
      Top             =   5340
      Width           =   2985
   End
   Begin VB.Label lblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Licensed to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   7080
      TabIndex        =   3
      Top             =   5085
      Width           =   1950
   End
   Begin VB.Label lblProductName 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":8EF1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   2160
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   6135
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2012"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public isOKtoUpdate As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If MsgBox("Want to Cancel Establishment of Connection?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
        Unload Me
        End
    End If
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Label9.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

'---LOADING TXT FILES---------------
    LogLocation = readTXTDATA("Location", "log", App.path & "\data\SystemDefault.ini")
    AuditLog = readTXTDATA("Location", "audit", App.path & "\data\SystemDefault.ini")
    AViLocation = readTXTDATA("Location", "Avis", App.path & "\data\SystemDefault.ini")
    ReportLocation = readTXTDATA("Location", "reports", App.path & "\data\SystemDefault.ini")
'------------------------------------
lblWarning.Caption = "Checking and Establishing Connection . . ."
'Animation1.Open AViLocation & "\FrontPage site scan.avi"
'Animation1.Visible = True
'Animation1.Play
Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set frmSplash = Nothing
End Sub
Private Sub SetDate()
Dim opnTime As New ADODB.Recordset

On Error GoTo handler

opnTime.Open "Select GEtdate() as SysDate", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnTime.RecordCount <> 0 Then
    Date = opnTime!SysDate
    Time = Format(opnTime!SysDate, "h:mm:ss ampm")
Else
    Date = Date
End If
opnTime.Close
Set opnTime = Nothing

handler:
If err.Number <> 0 Then
    Date = Date
    Exit Sub
End If
End Sub
Public Sub reconnect()
Dim ResolutionChange As Integer
Dim ServerTimeSync As Integer
Dim server As String
Dim login As String
Dim password As String
On Error GoTo handler
'---------------------------Connecting to Databases -----------------------------------------------
server = readTXTDATA("Database Type", "SSS", App.path & "\data\SystemDefault.ini")
login = readTXTDATA("Database Type", "LLL", App.path & "\data\SystemDefault.ini")
password = readTXTDATA("Database Type", "PPP", App.path & "\data\SystemDefault.ini")

EMode = readTXTDATA("SEM", "mode", App.path & "\data\SystemDefault.ini") 'System encryption mode
If EMode = 1 Then
    server = DecryptString(server)
    login = DecryptString(login)
    password = DecryptString(password)
End If

dbPMIS = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & login & ";Initial Catalog=pmis;Data Source=" & server & ""
dbFMIS = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & login & ";Initial Catalog=fmis;Data Source=" & server & ""


opndbaseFMIS.CommandTimeout = 0
opndbasePMIS.CursorLocation = adUseClient
opndbasePMIS.Open dbPMIS
InitErrMsgType = 1 'Connected to PMIS


opndbaseFMIS.CommandTimeout = 0
opndbaseFMIS.CursorLocation = adUseClient
opndbaseFMIS.Open dbFMIS
InitErrMsgType = 2 'Connected to FMIS
Exit Sub
handler:
MsgBox err.description
End Sub
Private Sub Timer1_Timer()
Dim ResolutionChange As Integer
Dim ServerTimeSync As Integer
Dim server As String
Dim login As String
Dim password As String
On Error GoTo handler
'---------------------------Connecting to Databases -----------------------------------------------
server = readTXTDATA("Database Type", "SSS", App.path & "\data\SystemDefault.ini")
login = readTXTDATA("Database Type", "LLL", App.path & "\data\SystemDefault.ini")
password = readTXTDATA("Database Type", "PPP", App.path & "\data\SystemDefault.ini")

EMode = readTXTDATA("SEM", "mode", App.path & "\data\SystemDefault.ini") 'System encryption mode
If EMode = 1 Then
    server = DecryptString(server)
    login = DecryptString(login)
    password = DecryptString(password)
End If

dbPMIS = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & login & ";Initial Catalog=pmis;Data Source=" & server & ""
dbFMIS = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & login & ";Initial Catalog=fmis;Data Source=" & server & ""

opndbasePMIS.ConnectionTimeout = 120
opndbasePMIS.CursorLocation = adUseClient
opndbasePMIS.Open dbPMIS
InitErrMsgType = 1 'Connected to PMIS

opndbaseFMIS.ConnectionTimeout = 120
opndbaseFMIS.CursorLocation = adUseClient
opndbaseFMIS.Open dbFMIS
InitErrMsgType = 2 'Connected to FMIS

'Setting the right workstation date--------'
ServerTimeSync = readTXTDATA("TimeSync", "mode", App.path & "\data\SystemDefault.ini")
Timer1.Enabled = False
 If CheckTheUpdate(App.Major & "." & App.Minor & "." & App.Revision) = True Then
        isOKtoUpdate = False
        frm_updatedesc.Show 1
        If isOKtoUpdate = True Then
            Shell App.path & "\AOMSUpdate.exe", vbNormalFocus
            End
        End If
 End If
 
Unload Me
frmUserPassword.whatLog = 2
frmUserPassword.Show
Exit Sub
handler:
MsgBox err.description
If err.Number <> 0 Then
    Select Case InitErrMsgType
        Case 0: 'Unable to Connect to PMIS Database
            lblWarning.Caption = "Unable to Connect PMIS Server!"
            MsgBox "PMIS Server Connection, Failed!", vbCritical, "System Warning"
        Case 1: 'Unable to Connect to FMIS Database
            lblWarning.Caption = "Unable to Connect FMIS Server!"
            MsgBox "FMIS Server Connection, Failed!", vbCritical, "System Warning"
        Case 2: 'Unable to change date
            lblWarning.Caption = "Unable to Set System Date and Time!"
            MsgBox "Setting System Date and Time, Failed!", vbCritical, "System Warning"
        Case Else
            lblWarning.Caption = ""
            MsgBox err.description, vbCritical, "System Warning"
    End Select
    Unload Me
    Timer1.Enabled = False
    End
End If
End Sub
Private Sub Timer2_Timer()
DoEvents
End Sub

