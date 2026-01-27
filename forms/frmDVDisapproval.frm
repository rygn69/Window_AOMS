VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDVDisapproval 
   Caption         =   "Log Out (DV Numbering)"
   ClientHeight    =   9360
   ClientLeft      =   1395
   ClientTop       =   4155
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   12225
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   13150
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox chkReturn 
      Caption         =   "Return"
      Height          =   375
      Left            =   420
      TabIndex        =   4
      Top             =   10560
      Width           =   1215
   End
   Begin VB.TextBox txtDetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   2
      Top             =   11145
      Width           =   4680
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   10920
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check no"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   405
      TabIndex        =   1
      Top             =   630
      Width           =   690
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   960
      Left            =   120
      Top             =   240
      Width           =   5550
   End
End
Attribute VB_Name = "frmDVDisapproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCLose_Click()
    Unload Me
End Sub

Private Sub btnOut_Click()
    If MsgBox("Are you sure you want to Log Out this DV No. " & txtDVNo.Text & "?", vbQuestion + vbYesNo) = vbYes Then
        opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns Set actioncode=4, PAoutDate='" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "', PADesc='" & Replace(Trim(txtDetail.Text), "'", "''") & "', OutBy='" & ActiveUserID & "',ReturnFlag=" & chkReturn.Value & " Where DVNo='" & txtDVNo.Text & "' And ActionCode=1"
        frmIncomingTrn.lblRefresh.Caption = "True"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    txtDVNo.Text = DVNoOut
End Sub
Private Function loaddv()
Dim rec As New ADODB.Recordset
rec.Open "Select "
End Function

