VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_listener 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading the Update"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   Icon            =   "frm_UploadIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_UploadIS.frx":1042
   ScaleHeight     =   1365
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6480
      Top             =   0
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      FullWidth       =   97
      FullHeight      =   17
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait...., the System will Freeze the Moment to download the update."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7455
   End
End
Attribute VB_Name = "frm_listener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbFMIS As String
dbFMIS = readTXTDATA("Database Type", "FMIS", App.Path & "\data\SystemDefault.ini")
strConnString = dbFMIS
End Sub

Private Sub Timer1_Timer()
Dim strTempPath As String
Dim strTempName As String
Dim strTempFile As String
Dim blnShow As Boolean

    'Create a temp file name
        Call PlayAVI(Me.Animation1, "horizontaloading.avi")
    strTempPath = App.Path & "\"
    strTempName = "Accounting Operation Management System.exe"
    strTempFile = strTempPath & strTempName
bck:
On Error GoTo bad:
    blnShow = ViewFromDB(1, strTempFile)
    
    If blnShow = True Then
        MsgBox "Update Successfully Downloaded....", vbInformation, "System Information"
    End If
        Call StopAvi(Me.Animation1)
        Timer1.Enabled = False
        Shell App.Path & "\Accounting Operation Management System.exe", vbNormalFocus
        End
Exit Sub
bad:
 If Err.Number = 3004 Then
    'MsgBox Err.Number
        If MsgBox("The AOMS is Already open, Please close all AOMS System to Receive the Update..", vbCritical + vbRetryCancel, "System Information") = vbRetry Then
            GoTo bck
        End If
    End If
    
End Sub
