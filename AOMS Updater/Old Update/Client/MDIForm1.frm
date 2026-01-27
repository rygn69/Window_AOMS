VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "AOMS System Update Logger"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7260
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":09EA
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   7230
      TabIndex        =   0
      Top             =   0
      Width           =   7260
      Begin VB.Label lblsend 
         AutoSize        =   -1  'True
         Caption         =   "Accounting Operation Management System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   3870
      End
      Begin VB.Label Label1 
         Caption         =   "File to Send:"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3480
      Top             =   2640
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu browse 
         Caption         =   "Browse File to Send"
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu si 
         Caption         =   "System Interval to Read"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub browse_Click()
frmBrowse.Show
End Sub


Private Sub si_Click()
Timer1.Interval = InputBox("Enter Interval: ", "System Interval")
End Sub

Private Sub Timer1_Timer()
Dim rec As New ADODB.Recordset
Dim frm As New frmClient
Set rec = opndbase.Execute("Select top 1 * from tblAMIS_UserUpdate where actioncode =0 order by ID asc")
If rec.RecordCount > 0 Then

      m_FileCompletePath = App.Path & "\Accounting Operation Management System.exe"
      m_FileName = "Accounting Operation Management System.exe"

        frm.LocalIP = rec!IP
        opndbase.Execute ("Update tblAMIS_UserUpdate set actioncode =1 Where ip = '" & Trim(rec!IP) & "'")
        rec.Close
        frm.Show
End If
End Sub
