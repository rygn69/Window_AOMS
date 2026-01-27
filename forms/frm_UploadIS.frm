VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_UploadIS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AOMS System Update Uploading"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   Icon            =   "frm_UploadIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6375
   Visible         =   0   'False
   Begin VB.TextBox txtISID 
      Appearance      =   0  'Flat
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
      Left            =   840
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2400
      Width           =   6015
   End
   Begin VB.TextBox txtversion 
      Appearance      =   0  'Flat
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
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Upload"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_UploadIS.frx":3AFA
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Browse"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_UploadIS.frx":4B4C
      cBack           =   16777215
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   688
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   273
      FullHeight      =   26
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IS ID:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   735
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
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
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
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Uploading System Update. Please wait....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frm_UploadIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cc_Click()

End Sub

Private Sub cmdSave_Click()
On Error GoTo cmdSave_Error
If txtISID.Text = 0 Then
    MsgBox "Invalid IS ID", vbCritical, "System Message"
    Exit Sub
End If
If MsgBox("Are you sure do you want to upload?", vbInformation + vbYesNo, "System Message") = vbYes Then
    lblstatus.Visible = True
    Call PlayAVI(Me.Animation1, "horizontaloading.avi")
    AddImageToDB CommonDialog1.FileName, 1, Text1.Text, txtname.Text, txtversion.Text, txtISID.Text
    Call StopAvi(Me.Animation1)
    lblstatus.Visible = False
    MsgBox "Successfully Uploaded in Database."
End If
Exit Sub
cmdSave_Error: MsgBox err.description
End Sub

Private Sub lvButtons_H1_Click()
On Error GoTo bad
With CommonDialog1
    .CancelError = True
    .Filter = "Exe File (*.exe"
    .ShowOpen
    lblpath.Caption = .FileName
    txtname.Text = .FileTitle
    Call getExeDetails
    'txtversion.Text = App.Major & "." & App.Minor & "." & (App.Revision - 1)
End With
Exit Sub
bad:
MsgBox err.description
End Sub
'Private Function getExeID(ExeName As String) As Integer
'    Dim rec As New ADODB.Recordset
'    getExeID = 0
'  Set rec = opndbaseFMIS.Execute("select ISID,Name from [dbo].[tblAMIS_SystemUpdate] where Name = '" & txtName.Text & "'")
'    If rec.RecordCount > 0 Then
'        getExeID = rec!isID
'    End If
'    rec.Close
'End Function
Private Sub getExeDetails()
    Dim rec As New ADODB.Recordset
   
  Set rec = opndbaseFMIS.Execute("select ISID,Description,[Version] from [dbo].[tblAMIS_SystemUpdate] where Name = '" & txtname.Text & "'")
    If rec.RecordCount > 0 Then
        Text1.Text = rec!description
        txtISID.Text = rec!isID
        txtversion.Text = rec!Version
    End If
    rec.Close
End Sub

