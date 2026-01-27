VERSION 5.00
Begin VB.Form frmApprovedJEVLogOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DV (with Approved JEV) Log Out Registry"
   ClientHeight    =   2355
   ClientLeft      =   4470
   ClientTop       =   2880
   ClientWidth     =   7245
   Icon            =   "frmApprovedJEVLogOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7245
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
      Height          =   420
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1185
   End
   Begin VB.CommandButton cmdLogOut 
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
      Left            =   5880
      Picture         =   "frmApprovedJEVLogOut.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1185
   End
   Begin VB.TextBox txt_Remark 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   135
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1350
      Width           =   4620
   End
   Begin VB.TextBox Text1 
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
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   4590
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
      Left            =   135
      TabIndex        =   3
      Top             =   1005
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter DV Number:"
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
      Top             =   120
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   3435
      Left            =   0
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "frmApprovedJEVLogOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLogOut_Click()
If Len(txt_Remark.Text) > 0 Then
    Select Case DVApprovedForLogOut(Text1.Text)
        Case 0 'For Approval
            MsgBox "Specified DV No. is Still For Approval!", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
        Case 2 'For Log-Out
            If MsgBox("Are you sure want to Log-Out this Voucher having an Approved JEV?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
                
                If Len(ActiveUserID) > 0 Then
                    opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set LogOutBy='" & ActiveUserID & "',LogOutDateTime='" & Now & "',LogOutRemark='" & txt_Remark.Text & "' " & _
                        " where DVNo='" & Text1.Text & "' and actioncode=1"
                    MsgBox "Voucher Logged-Out, Successfully!", vbInformation, "Sytem Information"
                    Call ClearEntry
                Else
                    Exit Sub
                End If
            End If
        Case 3 'Already Log-Out
            MsgBox "Specified DV No. was Already Logged-Out!", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
        Case 4 'Unregistered
            MsgBox "Specified DV No. was not yet Registered!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
    End Select
Else
    MsgBox "Specify Any Remarks for this Transaction!", vbInformation, "System Information"
    txt_Remark.SelStart = 0
    txt_Remark.SelLength = Len(txt_Remark.Text)
    txt_Remark.SetFocus
End If
End Sub
Private Sub ClearEntry()
txt_Remark.Text = ""
Text1.Text = ""
End Sub

Private Function DVApprovedForLogOut(ByVal DVNo As String) As Integer
Dim opnDV As New ADODB.Recordset

opnDV.Open "Select DVNo,ApprovedByID,LogOutBy from tblAMIS_JournalEntry where DVNo='" & DVNo & "' and actioncode=1 group by DVNo,ApprovedByID,LogOutBy", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opnDV.RecordCount <> 0 Then
        If Len(opnDV!ApprovedByID) > 0 Then
            'DVApprovedForLogOut = 1 'Already Approved
            If Len(opnDV!Logoutby) > 0 Then
                DVApprovedForLogOut = 3 'Already Log-Out
            Else
                DVApprovedForLogOut = 2 'For Log-Out
            End If
        Else
            DVApprovedForLogOut = 0 'For Approval
        End If
    Else
        DVApprovedForLogOut = 4 'Unregistered
    End If
opnDV.Close
Set opnDV = Nothing
End Function


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdLogOut_Click
End If
End Sub


Private Sub txt_Remark_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     cmdLogOut_Click
End If
End Sub
