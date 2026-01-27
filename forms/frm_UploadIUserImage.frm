VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_UploadIS 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_UploadIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    On Error GoTo cmdSave_Error
    With CommonDialog1
        .CancelError = True
        .Filter = "Image Files (*.exe"
        .ShowOpen
    End With
    
    AddImageToDB CommonDialog1.FileName, 1, "File added to database"
    
Exit Sub
cmdSave_Error:
End Sub

Private Sub cmdView_Click()
Dim strTempPath As String
Dim strTempName As String
Dim strTempFile As String
Dim blnShow As Boolean

    'Create a temp file name
    strTempPath = App.Path & "\"
    strTempName = Format(Now, "MMDDYYHHNNSS") & ".bmp"
    strTempFile = strTempPath & strTempName
    
    blnShow = ViewFromDB(1, strTempFile)
    
    If blnShow Then
        Picture1.Picture = LoadPicture(strTempFile)
        DoEvents
        Kill (strTempFile)
    End If
End Sub

Private Sub Form_Load()
    'Set the connectionstring to your database
    strConnString = "Provider=SQLOLEDB.1;Password=flamex;Persist Security Info=True;User ID=sa;Initial Catalog=master;timeout expired=0;Data Source= 192.168.2.1"
End Sub

