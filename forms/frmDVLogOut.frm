VERSION 5.00
Begin VB.Form frmDVLogOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Out (DV Numbering)"
   ClientHeight    =   3015
   ClientLeft      =   1380
   ClientTop       =   4140
   ClientWidth     =   6480
   Icon            =   "frmDVLogOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6480
   Begin VB.ComboBox cmb_approved 
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
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
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
      Height          =   540
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton btnOut 
      Caption         =   " &Log out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   5160
      Picture         =   "frmDVLogOut.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1185
   End
   Begin VB.CheckBox chkReturn 
      BackColor       =   &H80000012&
      Caption         =   "Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1950
      Width           =   855
   End
   Begin VB.TextBox txtDetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   200
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "ok"
      Top             =   2265
      Width           =   4560
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1365
      Width           =   4590
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Audited By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DV Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   990
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2880
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmDVLogOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnOut_Click()
    If Trim(cmb_approved.Text) = "" Then
               MsgBox "Please Specify the Approved Name of the transaction", vbInformation, "System Message"
              ' cmb_approved.SetFocus
               Exit Sub
    End If
    If MsgBox("Are you sure you want to Log Out this DV No. " & txtDVNo.Text & "?" & vbNewLine & "Audited by: " & cmb_approved.Text & "", vbQuestion + vbYesNo) = vbYes Then
        opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns Set PAout=1, PAoutDate='" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "', PADesc='" & Replace(Trim(txtDetail.Text), "'", "''") & "', OutBy='" & ActiveUserID & "',ReturnFlag=" & chkReturn.Value & " Where DVNo='" & txtDVNo.Text & "' And ActionCode=1"
        Call LogApprovedAndAudit(txtDVNo.Text, "Auditby", cmb_approved.ItemData(cmb_approved.ListIndex))
        frmIncomingTrn.lblRefresh.Caption = "True"
        Unload Me
    End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    txtDVNo.Text = DVNoOut
    Call GetSignatory(cmb_approved, "Audit by")
    'cmb_approved.ListIndex = 0
End Sub

