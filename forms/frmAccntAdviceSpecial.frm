VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAccntAdviceSpecial 
   Caption         =   "Check Details for Manually Issued Check"
   ClientHeight    =   4620
   ClientLeft      =   2010
   ClientTop       =   2580
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   9765
   Begin VB.Frame Frame1 
      Caption         =   "Authorized by :"
      Height          =   2355
      Left            =   6000
      TabIndex        =   14
      Top             =   300
      Width           =   3300
      Begin VB.CommandButton Command2 
         Caption         =   "Include to Selection"
         Height          =   480
         Left            =   1350
         TabIndex        =   19
         Top             =   1725
         Width           =   1785
      End
      Begin VB.TextBox txt_pword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1050
         PasswordChar    =   "@"
         TabIndex        =   16
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox txt_userid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1050
         PasswordChar    =   "@"
         TabIndex        =   15
         Top             =   465
         Width           =   2085
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   1275
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   555
         Width           =   630
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1095
      TabIndex        =   13
      Top             =   1950
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   17563649
      CurrentDate     =   40360
   End
   Begin VB.TextBox txt_amount 
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
      Left            =   1095
      TabIndex        =   11
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   5085
      TabIndex        =   9
      Top             =   3225
      Width           =   270
   End
   Begin VB.TextBox txt_payee 
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
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3225
      Width           =   3960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   600
      Left            =   7230
      TabIndex        =   5
      Top             =   2955
      Width           =   2055
   End
   Begin VB.ComboBox cmb_BankAccntNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1125
      TabIndex        =   4
      Top             =   2565
      Width           =   4245
   End
   Begin VB.TextBox txt_BankName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   945
      Width           =   4110
   End
   Begin VB.TextBox txt_CheckNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   4110
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Chk. Date"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   375
      TabIndex        =   12
      Top             =   1890
      Width           =   705
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Chk. Amount"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   330
      TabIndex        =   10
      Top             =   3780
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payee"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   345
      TabIndex        =   8
      Top             =   3345
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account No."
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2490
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   390
      TabIndex        =   3
      Top             =   945
      Width           =   705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Check No"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   375
      TabIndex        =   1
      Top             =   270
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1620
      Left            =   0
      Top             =   0
      Width           =   5340
   End
End
Attribute VB_Name = "frmAccntAdviceSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function PWAuthentic(ByVal UserID As String, ByVal Pword As String) As Boolean
Dim opn As New ADODB.Recordset

opn.Open "Select * from tblAMIS_UserAdvance where actioncode=1 and UserID='" & UserID & "' and pword='" & mydll.Encrypt(Pword) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opn.RecordCount <> 0 Then
    PWAuthentic = True
Else
    PWAuthentic = False
End If
opn.Close
Set opn = Nothing

End Function
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Len(cmb_BankAccntNo.Text) > 0 And Len(txt_payee.Text) > 0 And Len(txt_amount.Text) > 0 Then
    If PWAuthentic(txt_userid.Text, txt_pword.Text) = True Then
        
                        If Len(frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 0)) Then
                            frmAccountantsAdvice.MSHFlexGrid1.Rows = frmAccountantsAdvice.MSHFlexGrid1.Rows + 1 'Adding New Row
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 0) = cmb_BankAccntNo.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 1) = txt_CheckNo.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 2) = DTPicker1.Value
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 3) = txt_payee.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 4) = txt_amount.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 5) = GetBankIDbyBankName(txt_BankName.Text)
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 7) = 1
                        Else
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 0) = cmb_BankAccntNo.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 1) = txt_CheckNo.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 2) = DTPicker1.Value
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 3) = txt_payee.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 4) = txt_amount.Text
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 5) = GetBankIDbyBankName(txt_BankName.Text)
                            frmAccountantsAdvice.MSHFlexGrid1.TextMatrix(frmAccountantsAdvice.MSHFlexGrid1.Rows - 1, 7) = 1
                        End If
                        
                        frmAccountantsAdvice.txt_TotalAmt.Text = Format(GetTotalEnteredAmtInGrid(frmAccountantsAdvice.MSHFlexGrid1, 4, 1), "###,##0.00")
                        
                        opndbaseFMIS.Execute "Insert Into tblAMIS_AcctntAdviceAuthorizedCheck(checkNo,AuthorizedBy,DateAuthorized,actioncode) " & _
                                " values ('" & txt_CheckNo.Text & "','" & txt_userid.Text & "','" & DTPicker1.Value & "',1)"
                        Unload Me
    Else
        MsgBox "Invalid Authorization Code!", vbCritical, "System Warning"
    End If
Else
    MsgBox "Some Field Missing!", vbInformation, "System Information"
End If
End Sub

Private Sub Command3_Click()
ActiveFormCaller = Me.Name
frmCDClaimantRegistry.Show vbModal
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAccntAdviceSpecial = Nothing
End Sub

Private Sub txt_amount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(txt_amount.Text) <> 0 Then
        txt_amount.Text = Format(txt_amount.Text, "#,##0.00")
    Else
        txt_amount.Text = Format(0, "#,##0.00")
    End If
End If
End Sub

Private Sub txt_amount_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txt_amount_LostFocus()
    If Len(txt_amount.Text) <> 0 Then
        txt_amount.Text = Format(txt_amount.Text, "#,##0.00")
    Else
        txt_amount.Text = Format(0, "#,##0.00")
    End If

End Sub

Private Sub txt_BankName_Change()
Call LoadAllBankAccountNos(txt_BankName.Text, cmb_BankAccntNo)
End Sub
