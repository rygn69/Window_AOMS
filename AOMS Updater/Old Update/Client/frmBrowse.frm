VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBrowse.frx":1272
   ScaleHeight     =   6615
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
   End
   Begin VB.FileListBox fileFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   2760
      System          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
   End
   Begin VB.DirListBox fileDirectory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.DriveListBox fileDrive 
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
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8175
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   Dim tmpDir As String
      
   If fileFile.FileName <> vbNullString Then
      'Process here!
      
      tmpDir = fileFile.Path
      If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
      MDIForm1.lblsend = fileFile.FileName
      frmClient.EnterFileData tmpDir, fileFile.FileName
   End If
   Unload Me
End Sub

Private Sub fileDirectory_Change()
   fileFile.Path = fileDirectory.Path
End Sub

Private Sub fileDrive_Change()
   On Error GoTo ErrorHandle
   
   fileDirectory.Path = fileDrive.Drive
   fileDirectory.SetFocus
   
   Exit Sub
ErrorHandle:
   MsgBox "Drive not available.", vbCritical, "Error"
   fileDrive.Drive = "c:\"
   fileDirectory.Path = "c:\"
End Sub

Private Sub fileFile_DblClick()
   Dim tmpDir As String
   
   tmpDir = fileFile.Path
   If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
   MDIForm1.lblsend = fileFile.FileName
   frmClient.EnterFileData tmpDir, fileFile.FileName
   
   Unload Me
End Sub

Private Sub fileFile_KeyPress(KeyAscii As Integer)
   Dim tmpDir As String
   
   If KeyAscii = vbKeyReturn Then
      
      KeyAscii = 0
      
      If fileFile.FileName <> vbNullString Then
         tmpDir = fileFile.Path
         If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
         frmClient.EnterFileData tmpDir, fileFile.FileName
         Unload Me
         
      End If
   
   
   End If
   
End Sub

Private Sub Form_Load()
   Dim TempPath As String
   TempPath = App.Path
   
   If Left$(TempPath, 1) = "\" Then TempPath = "C:\"
      
   DoEvents
      
   'Default directory
   fileDrive.Drive = Left$(TempPath, 3)
   fileDirectory.Path = Left$(TempPath, 3)
   
   'Set focus to directory list
   frmBrowse.Visible = True
   DoEvents
   fileDirectory.SetFocus
   
End Sub
