VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_JEVApproval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Approval Registry"
   ClientHeight    =   3975
   ClientLeft      =   4020
   ClientTop       =   5055
   ClientWidth     =   6795
   Icon            =   "frm_JEVApproval.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6795
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
      Left            =   165
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2745
      Width           =   4500
   End
   Begin VB.CommandButton cmdlogout 
      Caption         =   "&Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   5400
      Picture         =   "frm_JEVApproval.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1305
   End
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
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   " &Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   5400
      Picture         =   "frm_JEVApproval.frx":4264
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   375
      Width           =   1305
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
      Height          =   420
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1305
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   465
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "View JEV Details"
      Top             =   840
      Width           =   375
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   582
      ButtonWidth     =   3334
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&View Approved JEV"
         EndProperty
      EndProperty
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
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   4470
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
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Approved By"
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
      Left            =   165
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
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
      Top             =   480
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   3675
      Left            =   0
      Top             =   360
      Width           =   5325
   End
End
Attribute VB_Name = "frm_JEVApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim tmpAppID As String
    Select Case DVApproved(Text1.Text)
        Case 0 'For Approval
            If Trim(cmb_approved.Text) = "" Then
                MsgBox "Please Specify the Approved Name of the transaction", vbInformation, "System Message"
                cmb_approved.SetFocus
                Exit Sub
            End If
            If MsgBox("Are you sure want to Approved this JEV", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
                
                If Len(ActiveUserID) > 0 Then
                    Call LogApprovedAndAudit(Text1.Text, "Approvedby", cmb_approved.ItemData(cmb_approved.ListIndex))
                    opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set ApprovedByID='" & ActiveUserID & "',DateTimeApproved='" & Now & "' " & _
                        " where DVNo='" & Text1.Text & "' and actioncode=1"
                    MsgBox "Transaction Approved!", vbInformation, "Sytem Information"
                    cmb_approved.ListIndex = 0
                Else
                    Exit Sub
                End If
            End If
        Case 1 'Approved
            MsgBox "Specified DV No. was Already Approved!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
        Case 4 'Not Yet Assigned
            MsgBox "Specified DV No. was not yet Registered!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
    End Select
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Len(Trim(Text1.Text)) > 0 Then
    frmJEVNumberingAssignment.txt_DVNo.Text = Text1.Text
    frmJEVNumberingAssignment.Label1.Visible = False
    frmJEVNumberingAssignment.txt_Jevno.Visible = False
    frmJEVNumberingAssignment.Shape1.Visible = False
    frmJEVNumberingAssignment.Command1.Visible = False
    frmJEVNumberingAssignment.Show vbModal
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Call GetSignatory(cmb_approved, "Approved 2")
cmb_approved.ListIndex = 0
End Sub
Private Function DVApproved(ByVal DVNo As String) As Integer
Dim opnDV As New ADODB.Recordset
opnDV.Open "Select DVNo,ApprovedByID from tblAMIS_JournalEntry where DVNo='" & DVNo & "' and actioncode=1 group by DVNo,ApprovedByID", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opnDV.RecordCount <> 0 Then
        If Len(opnDV!ApprovedByID) > 0 Then
            DVApproved = 1
        Else
            DVApproved = 0
        End If
    Else
        DVApproved = 4
    End If
opnDV.Close
Set opnDV = Nothing
End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Approved JEV
        frmListOfApprovedJEV.Show vbModal
End Select
End Sub
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

Private Sub txt_Remark_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     cmdLogOut_Click
End If
End Sub

