VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSystemUsers 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Users Restriction Level Registry"
   ClientHeight    =   8640
   ClientLeft      =   1890
   ClientTop       =   2430
   ClientWidth     =   11580
   Icon            =   "frm_Users.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11580
   Begin VB.CommandButton FlatBttn1 
      Caption         =   "&New/Reset"
      Height          =   795
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   330
      Width           =   1290
   End
   Begin VB.CommandButton FlatBttn2 
      Caption         =   "&Save"
      Height          =   795
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   330
      Width           =   1290
   End
   Begin VB.CommandButton FlatBttn3 
      Caption         =   "&Edit"
      Height          =   795
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   330
      Width           =   1290
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   630
      Top             =   8370
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      FrameControl    =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   180
      TabIndex        =   10
      Top             =   1395
      Width           =   4080
      Begin VB.TextBox Txt_UserId 
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
         Left            =   855
         TabIndex        =   14
         Top             =   345
         Width           =   3135
      End
      Begin VB.TextBox txt_Username 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   885
         Width           =   3135
      End
      Begin VB.TextBox txt_Password 
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
         IMEMode         =   3  'DISABLE
         Left            =   855
         PasswordChar    =   "@"
         TabIndex        =   12
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserID"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   1515
         Width           =   690
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5505
      Top             =   -90
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "System Restriction Level"
      Height          =   3375
      Left            =   4995
      TabIndex        =   1
      Top             =   360
      Width           =   6315
      Begin VB.CheckBox chk_AcctAdvice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Accountant's Advice Preparation"
         Height          =   375
         Left            =   390
         TabIndex        =   29
         Top             =   2790
         Width           =   2235
      End
      Begin VB.CheckBox chk_Maintenance 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Maintenance"
         Height          =   195
         Left            =   2850
         TabIndex        =   24
         Top             =   1935
         Width           =   1365
      End
      Begin VB.Frame fra_Maintenance 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   3000
         TabIndex        =   25
         Top             =   1935
         Width           =   2685
         Begin VB.OptionButton opn_CAccnt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Chart of Accounts"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   225
            Width           =   1620
         End
         Begin VB.OptionButton opn_Users 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Users Account Registry"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   540
            Width           =   2055
         End
         Begin VB.OptionButton opn_MaintenanceAll 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Allow All"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   870
            Width           =   1005
         End
      End
      Begin VB.CheckBox chk_JEVApproval 
         BackColor       =   &H00E0E0E0&
         Caption         =   "JEV Approval"
         Height          =   195
         Left            =   390
         TabIndex        =   21
         Top             =   1485
         Width           =   1785
      End
      Begin VB.CheckBox chk_JEVPreparation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "JEV Preparation"
         Height          =   195
         Left            =   390
         TabIndex        =   11
         Top             =   1020
         Width           =   1800
      End
      Begin VB.CheckBox chk_JEVNumbering 
         BackColor       =   &H00E0E0E0&
         Caption         =   "JEV Numbering"
         Height          =   195
         Left            =   2895
         TabIndex        =   5
         Top             =   555
         Width           =   1545
      End
      Begin VB.CheckBox chk_SUtilities 
         BackColor       =   &H00E0E0E0&
         Caption         =   "System Utilities"
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   2400
         Width           =   1620
      End
      Begin VB.CheckBox chk_JEVLogOut 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Approved JEV Log Out"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   1935
         Width           =   2040
      End
      Begin VB.CheckBox chk_Incoming 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incoming (Pre-Audit)"
         Height          =   195
         Left            =   390
         TabIndex        =   2
         Top             =   555
         Width           =   1830
      End
      Begin VB.Frame fra_JEV 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   2985
         TabIndex        =   6
         Top             =   570
         Width           =   2685
         Begin VB.OptionButton opn_Others 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Others"
            Height          =   195
            Left            =   1380
            TabIndex        =   23
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton opn_Collection 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Collection"
            Height          =   195
            Left            =   195
            TabIndex        =   22
            Top             =   840
            Width           =   1065
         End
         Begin VB.OptionButton opn_JEVAll 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Allow All"
            Height          =   195
            Left            =   1380
            TabIndex        =   9
            Top             =   510
            Width           =   1005
         End
         Begin VB.OptionButton opn_Check 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   540
            Width           =   795
         End
         Begin VB.OptionButton opn_CAsh 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cash"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   225
            Width           =   810
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4110
      Left            =   150
      TabIndex        =   0
      Top             =   4140
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   7250
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorBkg    =   14737632
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmSystemUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim trn As String



Private Sub chk_JEVNumbering_Click()
If chk_JEVNumbering.Value = 1 Then
    opn_JEVAll.Value = True 'Set as Default
Else
    Call ResetOptions("JEVNumbering")
End If
End Sub

Private Sub chk_Maintenance_Click()
If chk_Maintenance.Value = 1 Then
    opn_MaintenanceAll.Value = True 'Set as Default
Else
    Call ResetOptions("Maintenance")
End If
End Sub

Private Sub FlatBttn1_Click()


On Error GoTo handler

trn = 0
Call Reset
Call Enable
txt_userid.SetFocus
    

handler:
If err.Number <> 0 Then
    Exit Sub
End If
End Sub

Private Sub FlatBttn2_Click()
FlatBttn2.Enabled = False
Call SavEntry
FlatBttn2.Enabled = True
End Sub

Private Sub FlatBttn3_Click()
trn = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
Call Enable
Call Reset
Call LoadSelectedRec2TxtBoxes
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Dep As Variant

If KeyCode = vbKeyEscape Then
    Unload Me
ElseIf KeyCode = vbKeyF10 And Shift = 2 Then
    Dep = InputBox("Access Code:", "Advance Security Access Right Registry")

    If Len(Dep) = 0 Then
        Exit Sub
    Else
        'Enter Form for Granting Access Code for Authorization
        If Trim(ActiveUserID) = "0169" And Dep = "firebird" Then
            frmChkAuthorizedUser.Show vbModal
        End If
    End If
End If

'09306670080
End Sub
Private Function VerifyMaintenancelevel()

If chk_Maintenance.Value = 1 Then
    If opn_CAccnt.Value = True Then 'Chart of Accounts
        VerifyMaintenancelevel = 1
    ElseIf opn_Users.Value = True Then 'User's Registry
        VerifyMaintenancelevel = 2
    ElseIf opn_MaintenanceAll.Value = True Then 'All
        VerifyMaintenancelevel = 3
    Else
        VerifyMaintenancelevel = 0
    End If
Else
    VerifyMaintenancelevel = 0
End If
End Function
Private Function VerifyJEVNumberingLevel()
If chk_JEVNumbering.Value = 1 Then
    If opn_CAsh.Value = True Then 'Cash Disbursement
        VerifyJEVNumberingLevel = 1
    ElseIf opn_Check.Value = True Then 'Check Disbursement
        VerifyJEVNumberingLevel = 2
    ElseIf opn_Collection.Value = True Then 'Collection
        VerifyJEVNumberingLevel = 3
    ElseIf opn_Others.Value = True Then 'Others
        VerifyJEVNumberingLevel = 4
    ElseIf opn_JEVAll.Value = True Then 'All
        VerifyJEVNumberingLevel = 5
    Else 'Disallow
        VerifyJEVNumberingLevel = 0
    End If
Else
    VerifyJEVNumberingLevel = 0
End If
End Function

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
WindowsXPC1.InitSubClassing

'If UserLevel = "System Administrator" Then
'    txt_Password.PasswordChar = ""
'Else
'    txt_Password.PasswordChar = "@"
'End If


fra_JEV.Enabled = False
fra_Maintenance.Enabled = False

Timer1.Enabled = True
End Sub
Private Function FDuplicated(ByVal userno As String) As Boolean
Dim opnID As New ADODB.Recordset

opnID.Open "Select Userid from tblAMIS_UserRegistry where actioncode=1 and userid='" & userno & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnID.RecordCount <> 0 Then
    FDuplicated = True
Else
    FDuplicated = False
End If
opnID.Close
Set opnID = Nothing
End Function

Private Sub SavEntry()

If trn <> 0 Then 'Update Records
    
            
    opndbaseFMIS.Execute "Update tblAMIS_UserRegistry set ActionCode=2 where userid=" & txt_userid.Text & ""
            
    opndbaseFMIS.Execute "Insert into tblAMIS_UserRegistry (DateTimeEntered,enteredbyUserid,Actioncode,Userid,Username,UserPassword,Mod1,Mod2,Mod3,Mod4,Mod5,Mod6,Mod7,Mod8) " & _
                " values ('" & Now & "','" & ActiveUserID & "',1,'" & UCase(txt_userid.Text) & "','" & txt_Username.Text & "','" & mydll.Encrypt(UCase(txt_Password.Text)) & "', " & _
                " " & chk_Incoming.Value & "," & chk_JEVPreparation.Value & "," & chk_JEVApproval.Value & "," & chk_JEVLogOut.Value & "," & chk_SUtilities.Value & ", " & _
                " " & VerifyMaintenancelevel & "," & VerifyJEVNumberingLevel & "," & chk_AcctAdvice.Value & ")"
            
    Call Reset
    Call LoadSavedUser2grid
    Call Disable
    MsgBox "Updating Successful!", vbInformation, "System Information"
    

Else 'Save new Record
    If FDuplicated(txt_userid.Text) = False Then 'If is not yet existing------------
        If Len(Trim(txt_userid.Text)) = 0 Or Len(Trim(txt_Password.Text)) = 0 Then
                MsgBox "Please Complete Your entries!", vbInformation, "System Information"
        Else
        
            opndbaseFMIS.Execute "Insert into tblAMIS_UserRegistry (DateTimeEntered,enteredbyUserid,Actioncode,Userid,Username,UserPassword,Mod1,Mod2,Mod3,Mod4,Mod5,Mod6,Mod7,Mod8) " & _
                        " values ('" & Now & "','" & ActiveUserID & "',1,'" & UCase(txt_userid.Text) & "','" & txt_Username.Text & "','" & mydll.Encrypt(UCase(txt_Password.Text)) & "', " & _
                        " " & chk_Incoming.Value & "," & chk_JEVPreparation.Value & "," & chk_JEVApproval.Value & "," & chk_JEVLogOut.Value & "," & chk_SUtilities.Value & ", " & _
                        " " & VerifyMaintenancelevel & "," & VerifyJEVNumberingLevel & "," & chk_AcctAdvice.Value & ")"
            Call Reset
            Call LoadSavedUser2grid
            Call Disable
            MsgBox "Saving of New User Successful!", vbInformation, "System Information"
        End If
    Else
        MsgBox "The User you want to Add is already existing!", vbInformation, "System Information"
    End If
End If

End Sub
Private Sub LoadBackMaintenanceOptionSelection(ByVal OptionNo As Integer)
Select Case OptionNo
    Case 1 'opn_CAccnt
        opn_CAccnt.Value = True
    Case 2 'opn_Users
        opn_Users.Value = True
    Case 3 'opn_MaintenanceAll
        opn_MaintenanceAll.Value = True
End Select
End Sub
Private Sub LoadBackJEVNumberingOptionSelection(ByVal OptionNo As Integer)
Select Case OptionNo
    Case 1 'opn_CAsh
        opn_CAsh.Value = True
    Case 2 'opn_Check
        opn_Check.Value = True
    Case 3 'opn_Collection
        opn_Collection.Value = True
    Case 4 'opn_Others
        opn_Others.Value = True
    Case 5 'opn_JEVAll
        opn_JEVAll.Value = True
End Select
End Sub
Private Sub LoadSelectedRec2TxtBoxes()

txt_userid.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
txt_Username.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
txt_Password.Text = mydll.ReverseTxt(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3))
txt_Password.Text = mydll.Convert2Chr(txt_Password.Text)


chk_Incoming.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
chk_JEVPreparation.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5)
chk_JEVApproval.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 6)
chk_JEVLogOut.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 7)
chk_SUtilities.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)
chk_AcctAdvice.Value = IIf(Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 11)) = 0, 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 11))

'------for Maintenance
If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 9) > 0 Then
    chk_Maintenance.Value = 1
    Call LoadBackMaintenanceOptionSelection(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 9))
Else
    chk_Maintenance.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 9)
    Call ResetOptions("Maintenance")
End If

'------for JEVNumbering
If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10) > 0 Then
    chk_JEVNumbering.Value = 1
    Call LoadBackJEVNumberingOptionSelection(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10))
Else
    chk_JEVNumbering.Value = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10)
    Call ResetOptions("JEVNumbering")
End If

End Sub

Private Sub LoadSavedUser2grid()
Dim opnuser As New ADODB.Recordset
Dim xx, yy As Integer


If Len(Trim(txt_userid.Text)) = 0 Then
    opnuser.Open "Select * from tblAMIS_UserRegistry where Actioncode=1 order by userid", opndbaseFMIS, adOpenStatic, adLockOptimistic
Else
    opnuser.Open "Select * from tblAMIS_UserRegistry where userid like '" & txt_userid.Text & "%' and Actioncode=1 order by userid", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If


If opnuser.RecordCount <> 0 Then
    
    Set MSHFlexGrid1.DataSource = opnuser
    
    'Set the ColWidth of the Gridlines------
    MSHFlexGrid1.ColWidth(0) = 0
    MSHFlexGrid1.ColWidth(1) = 1000
    MSHFlexGrid1.ColWidth(2) = 2000
    MSHFlexGrid1.ColWidth(3) = 1000
    MSHFlexGrid1.ColWidth(4) = 800
    MSHFlexGrid1.ColWidth(5) = 800
    MSHFlexGrid1.ColWidth(6) = 800
    MSHFlexGrid1.ColWidth(7) = 800
    MSHFlexGrid1.ColWidth(8) = 800
    MSHFlexGrid1.ColWidth(9) = 800
    MSHFlexGrid1.ColWidth(10) = 800
    MSHFlexGrid1.ColWidth(11) = 800
    MSHFlexGrid1.ColWidth(12) = 0
    MSHFlexGrid1.ColWidth(13) = 0
    MSHFlexGrid1.ColWidth(14) = 0
    

    MSHFlexGrid1.Visible = False
    '------align---
        For xx = 7 To MSHFlexGrid1.Cols - 1
            For yy = 0 To MSHFlexGrid1.Rows - 1
                MSHFlexGrid1.col = xx
                MSHFlexGrid1.Row = yy
                MSHFlexGrid1.CellAlignment = 4
            Next yy
        Next xx
    '--------------
    MSHFlexGrid1.Visible = True

Else
    MSHFlexGrid1.Visible = False
End If
opnuser.Close
Set opnuser = Nothing
End Sub
Private Sub Disable()
txt_userid.Enabled = False
'txt_Username.Enabled = False
txt_Password.Enabled = False
chk_Incoming.Enabled = False
chk_JEVPreparation.Enabled = False
chk_JEVApproval.Enabled = False
chk_JEVLogOut.Enabled = False
chk_SUtilities.Enabled = False
chk_JEVNumbering.Enabled = False
chk_AcctAdvice.Enabled = False
Call EnableOptions("JEVNumbering", False)
fra_JEV.Enabled = False
chk_Maintenance.Enabled = False
Call EnableOptions("Maintenance", False)
fra_Maintenance.Enabled = False
End Sub
Private Sub Enable()
txt_userid.Enabled = True
'txt_Username.Enabled = False
txt_Password.Enabled = True

chk_Incoming.Enabled = True
chk_JEVPreparation.Enabled = True
chk_JEVApproval.Enabled = True
chk_JEVLogOut.Enabled = True
chk_SUtilities.Enabled = True
chk_JEVNumbering.Enabled = True
chk_AcctAdvice.Enabled = True
Call EnableOptions("JEVNumbering", True)
fra_JEV.Enabled = True
chk_Maintenance.Enabled = True
Call EnableOptions("Maintenance", True)
fra_Maintenance.Enabled = True
End Sub
Private Sub Reset()
txt_userid.Text = ""
txt_Username.Text = ""
txt_Password.Text = ""

chk_AcctAdvice.Value = 0
chk_Incoming.Value = 0
chk_JEVPreparation.Value = 0
chk_JEVApproval.Value = 0
chk_JEVLogOut.Value = 0
chk_SUtilities.Value = 0

chk_JEVNumbering.Value = 0
Call ResetOptions("JEVNumbering")

chk_Maintenance.Value = 0
Call ResetOptions("Maintenance")

End Sub
Private Sub EnableOptions(ByVal ChkName As String, ByVal Enable As Boolean)
Select Case ChkName
    Case "Maintenance"
        opn_CAccnt.Enabled = Enable
        opn_Users.Enabled = Enable
        opn_MaintenanceAll.Enabled = Enable
    Case "JEVNumbering"
        opn_CAsh.Enabled = Enable
        opn_Check.Enabled = Enable
        opn_Collection.Enabled = Enable
        opn_Others.Enabled = Enable
        opn_JEVAll.Enabled = Enable
End Select
End Sub
Private Sub ResetOptions(ByVal ChkName As String)
Select Case ChkName
    Case "Maintenance"
        opn_CAccnt.Value = False
        opn_Users.Value = False
        opn_MaintenanceAll.Value = False
    Case "JEVNumbering"
        opn_CAsh.Value = False
        opn_Check.Value = False
        opn_Collection.Value = False
        opn_Others.Value = False
        opn_JEVAll.Value = False
End Select
End Sub


Private Sub LoadUser()
Dim opnuser As New ADODB.Recordset

opnuser.Open "Select * from Employee where SwipEmployeeID='" & txt_userid.Text & "'", opndbasePMIS, adOpenStatic, adLockOptimistic
    
    If opnuser.RecordCount <> 0 Then
        txt_Username.Text = UCase(opnuser!Firstname & " " & IIf(Len(Trim(opnuser!MI)) = 0, "", Left(opnuser!MI, 1) & ".") & " " & opnuser!Lastname & " " & IIf(Len(Trim(opnuser!Suffix)) = 0, "", "," & opnuser!Suffix))
        
        If Len(Trim(txt_Password.Text)) <> 0 Then '--On got fucos on password_txt ----
            txt_Password.SelStart = 0
            txt_Password.SelLength = Len(txt_Password.Text)
            txt_Password.SetFocus
        Else
            txt_Password.SetFocus
        End If '-----------------------------------------------------------------
    
    Else
        MsgBox "User ID No. you have currently entered is not registered in the PMIS!", vbInformation, "System Information"
    End If
opnuser.Close
Set opnuser = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMother.StatusBar1.Panels(3) = "Active Module :"
Set frmSystemUsers = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub

Private Sub MSHFlexGrid1_Click()
'Call Reset
trn = MSHFlexGrid1.ColWidth(0)
Call Disable
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyDelete Then
    If trn <> 0 Then
        If MsgBox("Are you sure want to delete this User?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
            opndbaseFMIS.Execute "update tblCMS_UserDetails set actioncode=3,datetimeentered='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & "," & Now & "',enteredbyuserid='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2) & "," & ActiveUserID & "' where trnno=" & trn & ""
            Call LoadSavedUser2grid
        End If
    End If
End If
End Sub


Private Sub Timer1_Timer()
trn = 0
Call LoadSavedUser2grid
Call Disable
Timer1.Enabled = False
End Sub




Private Sub Txt_UserId_Change()
If trn = 0 Then
    Call LoadSavedUser2grid
End If
End Sub

Private Sub Txt_UserId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt_userid.Text = UCase(txt_userid.Text)
    Call LoadUser
End If
End Sub

