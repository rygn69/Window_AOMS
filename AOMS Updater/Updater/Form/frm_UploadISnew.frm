VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloading the update"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   Icon            =   "frm_UploadISnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_UploadISnew.frx":3519A
   ScaleHeight     =   1395
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   0
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      FullWidth       =   97
      FullHeight      =   17
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait...., the system will Freeze for a moment to download the update."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7455
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbFMIS As String
'dbFMIS = readTXTDATA("Database Type", "FMIS", App.Path & "\data\SystemDefault.ini")
'strConnString = "Provider=SQLOLEDB.1;Password=(@/51u0#2@3n8D0e1L1#0u1R;Persist Security Info=True;User ID=pgasis;Data Source=192.168.2.1\pgas"
strConnString = "Provider=SQLOLEDB.1;Password=T9v#qE1r@Lx8Zp!f;Persist Security Info=True;User ID=proc;Initial Catalog=fmis;Data Source=vm_db_pmis.pgas.ph"
'-----------------------------
'PAYROLL
'exeID = 2
'exeName = "Epay.exe"
'-----------------------------

'AOMS
exeID = 1
exeName = "Accounting Operation Management System.exe"
End Sub

Private Sub Timer1_Timer()
Dim strTempPath As String
Dim strTempName As String
Dim strTempFile As String
Dim blnShow As Boolean

    'Create a temp file name
        Call PlayAVI(Me.Animation1, "horizontaloading.avi")
    strTempPath = App.Path & "\"
    strTempName = exeName
    strTempFile = strTempPath & strTempName
bck:
On Error GoTo bad:
    blnShow = ViewFromDB(1, strTempFile)
    
    If blnShow = True Then
        MsgBox "Update Successfully Downloaded....", vbInformation, "System Information"
    End If
        Call StopAvi(Me.Animation1)
        Timer1.Enabled = False
        Shell App.Path & "\" & exeName, vbNormalFocus
        End
Exit Sub
bad:
 If Err.Number = 3004 Then
    'MsgBox Err.Number
        If MsgBox("The " & exeName & " is Already open, Please close all AOMS System to Receive the Update..", vbCritical + vbRetryCancel, "System Information") = vbRetry Then
            GoTo bck
        End If
    End If
    
End Sub

