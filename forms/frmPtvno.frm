VERSION 5.00
Begin VB.Form frmPtvno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter PTV Number"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "frmPtvno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmPtvno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public checkno, Trnno As String

Private Sub Command1_Click()
    If CheckIF(Text1.Text) = True Then
    MsgBox "PTV number already exist on the database..!", vbInformation, "System Message"
    Exit Sub
    End If
If MsgBox("Are you sure do you want to Save??", vbInformation + vbYesNo, "System Message") = vbYes Then
opndbaseFMIS.Execute "update tblCMS_CDRCIReport set dvno = '" & Text1.Text & "',released = 2 where trnno = '" & Trnno & "' and actioncode = 1"
frmJEVNumberingThruRCI.grd_details.TextMatrix(ForTheGridRowNo, 5) = Text1.Text
frmJEVNumberingThruRCI.grd_details.TextMatrix(ForTheGridRowNo, 4) = "2"
Unload Me
End If
End Sub

Private Function CheckIF(ByVal ptvNo As String) As Boolean
Dim rec As New ADODB.Recordset
CheckIF = False
rec.Open "Select dvno from tblCMS_CDRCIReport where dvno = '" & ptvNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        CheckIF = True
    End If
rec.Close
Set rec = Nothing
End Function

Private Sub Command2_Click()
Unload Me
End Sub
