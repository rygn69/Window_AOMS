VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmChartOfAccounts 
   Caption         =   "Chart of Accounts Maintenance"
   ClientHeight    =   10035
   ClientLeft      =   1965
   ClientTop       =   1695
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   15150
   Begin VB.Frame Frame3 
      Caption         =   "Search (Enter Account Code)"
      Height          =   975
      Left            =   315
      TabIndex        =   25
      Top             =   1965
      Width           =   5715
      Begin VB.TextBox txt_AccntNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   26
         Top             =   300
         Width           =   5220
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Name Description"
      Height          =   2355
      Left            =   315
      TabIndex        =   16
      Top             =   3315
      Width           =   14505
      Begin VB.TextBox txt_MainAccntCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6465
         TabIndex        =   23
         Top             =   615
         Width           =   3555
      End
      Begin VB.TextBox txt_SubAccntNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6900
         TabIndex        =   21
         Top             =   1395
         Width           =   3120
      End
      Begin VB.TextBox txt_sub 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1155
         TabIndex        =   19
         Top             =   1395
         Width           =   5685
      End
      Begin VB.TextBox txt_main 
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
         Left            =   810
         TabIndex        =   17
         Top             =   615
         Width           =   5595
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NGAS Account Code |"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   10395
         TabIndex        =   28
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   10605
         TabIndex        =   27
         Top             =   1050
         Width           =   825
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   2250
         Left            =   10035
         Top             =   60
         Width           =   4500
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   1110
         Top             =   1350
         Width           =   13740
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   780
         Top             =   570
         Width           =   13740
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Account Code"
         Height          =   225
         Left            =   6480
         TabIndex        =   24
         Top             =   315
         Width           =   3435
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Account Code"
         Height          =   225
         Left            =   6930
         TabIndex        =   22
         Top             =   1125
         Width           =   3435
      End
      Begin VB.Line Line8 
         X1              =   390
         X2              =   390
         Y1              =   570
         Y2              =   2160
      End
      Begin VB.Line Line11 
         X1              =   585
         X2              =   1155
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line10 
         X1              =   585
         X2              =   585
         Y1              =   795
         Y2              =   2085
      End
      Begin VB.Line Line9 
         X1              =   390
         X2              =   795
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Account Name "
         Height          =   195
         Left            =   1170
         TabIndex        =   20
         Top             =   1155
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Account Name"
         Height          =   195
         Left            =   810
         TabIndex        =   18
         Top             =   375
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Height          =   600
      Left            =   3555
      TabIndex        =   15
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Edit"
      Height          =   600
      Left            =   2400
      TabIndex        =   14
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   600
      Left            =   1245
      TabIndex        =   13
      Top             =   1125
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   600
      Left            =   90
      TabIndex        =   12
      Top             =   1125
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Level of Category"
      Height          =   2535
      Left            =   9120
      TabIndex        =   3
      Top             =   720
      Width           =   5820
      Begin VB.ComboBox cmb_fourth 
         Height          =   315
         Left            =   1290
         TabIndex        =   7
         Top             =   1830
         Width           =   3585
      End
      Begin VB.ComboBox cmb_third 
         Height          =   315
         Left            =   1095
         TabIndex        =   6
         Top             =   1310
         Width           =   3765
      End
      Begin VB.ComboBox cmb_Second 
         Height          =   315
         Left            =   855
         TabIndex        =   5
         Top             =   790
         Width           =   3990
      End
      Begin VB.ComboBox cmb_First 
         Height          =   315
         Left            =   675
         TabIndex        =   4
         Top             =   270
         Width           =   4185
      End
      Begin VB.Line Line7 
         X1              =   450
         X2              =   720
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line6 
         X1              =   855
         X2              =   1320
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line5 
         X1              =   840
         X2              =   840
         Y1              =   1455
         Y2              =   2220
      End
      Begin VB.Line Line4 
         X1              =   630
         X2              =   1125
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Line Line3 
         X1              =   630
         X2              =   630
         Y1              =   930
         Y2              =   2220
      End
      Begin VB.Line Line2 
         X1              =   450
         X2              =   900
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   450
         Y1              =   390
         Y2              =   2205
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1st"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   315
         Width           =   210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2nd"
         Height          =   195
         Left            =   5040
         TabIndex        =   10
         Top             =   845
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3rd"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   1375
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4th"
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   1905
         Width           =   225
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3900
      Left            =   105
      TabIndex        =   2
      Top             =   6000
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   6879
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cmb_Fund 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   465
      Width           =   4005
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2640
      Left            =   9000
      Top             =   660
      Width           =   6225
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2490
      Left            =   60
      Top             =   3225
      Width           =   15045
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   315
      TabIndex        =   1
      Top             =   210
      Width           =   765
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   -90
      Top             =   -15
      Width           =   4695
   End
End
Attribute VB_Name = "frmChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TmpFMISCode As Long
Dim TmpRowNo As Long
Dim saveflag As Integer


Private Sub cmb_Fund_click()
Call LoadGroupLevel(cmb_Fund.Text, "FirstLevelGroup", cmb_fourth)
Call LoadGroupLevel(cmb_Fund.Text, "SecondLevelGroup", cmb_third)
Call LoadGroupLevel(cmb_Fund.Text, "ThirdLevelGroup", cmb_Second)
Call LoadGroupLevel(cmb_Fund.Text, "FourthLevelGroup", cmb_First)
End Sub

Private Sub Command1_Click()
saveflag = 1
Call Clear
Call FrameState(True)
End Sub

Private Sub Command2_Click()
Call SaveRec
Call LoadFilteredAccnts(txt_AccntNo.Text, cmb_Fund.Text)
End Sub

Private Sub Command3_Click()
saveflag = 2
Call FrameState(True)
Call LoadBackToEdit(TmpRowNo)
End Sub
Private Sub Command4_Click()
If Val(TmpFMISCode) > 0 Then
    If MsgBox("Are you Sure, Want to Remove this Account?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Update tblREF_AIS_ChartofAccounts set active=0 where FMISAccountCode=" & TmpFMISCode & ""
        MsgBox "Account Removed Successfuly!", vbInformation, "System Information"
    
        saveflag = 0
        Call FrameState(False)
        Call LoadGroupLevel(cmb_Fund.Text, "FirstLevelGroup", cmb_fourth)
        Call LoadGroupLevel(cmb_Fund.Text, "SecondLevelGroup", cmb_third)
        Call LoadGroupLevel(cmb_Fund.Text, "ThirdLevelGroup", cmb_Second)
        Call LoadGroupLevel(cmb_Fund.Text, "FourthLevelGroup", cmb_First)
        Call LoadFilteredAccnts(txt_AccntNo.Text, cmb_Fund.Text)
    
    End If
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Label1.Caption = ""
LoadFund
Call FrameState(False)
End Sub
Private Sub LoadFund()
Dim Frec As New ADODB.Recordset
Dim x As Integer

cmb_Fund.Clear

Frec.Open ("Select * from tblRefBMS_Funds Order By FundMedium"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    For x = 1 To Frec.RecordCount
        cmb_Fund.AddItem Frec!FundName
        cmb_Fund.ItemData(cmb_Fund.NewIndex) = Frec!FundCode
        Frec.MoveNext
    Next x
End If
Frec.Close
Set Frec = Nothing

End Sub

Private Sub SaveRec()
Dim SAC As String
            If Left(Trim(txt_SubAccntNo.Text), 1) = "-" Then
                SAC = Trim(txt_MainAccntCode.Text) & "" & Trim(txt_SubAccntNo.Text)
            Else
                SAC = Trim(txt_MainAccntCode.Text) & "-" & Trim(txt_SubAccntNo.Text)
            End If
Select Case saveflag
            
    Case 1 'New
        opndbaseFMIS.Execute "Insert into tblREF_AIS_ChartofAccounts (DateEntered,FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup, " & _
                " AccountCode,ChildSeriesNumber,ChildAccountCode,AccountName,AccountNameFull,NormalBalance,FundType, " & _
                " UserID,Active) " & _
                " values ('" & Date & "','" & cmb_First.Text & "','" & cmb_Second.Text & "','" & cmb_third.Text & "','" & cmb_fourth.Text & "', " & _
                " '" & txt_MainAccntCode.Text & "','" & txt_SubAccntNo.Text & "','" & SAC & "', " & _
                " '" & txt_main.Text & "','" & txt_sub.Text & "',1,'" & cmb_Fund.Text & "','" & ActiveUserID & "',1)"
                
                MsgBox "Saving New Account, Successful!", vbInformation, "System Information"
                Call Clear
    Case 2 'Edit
        opndbaseFMIS.Execute "update tblREF_AIS_ChartofAccounts set FourthLevelGroup='" & cmb_First.Text & "',ThirdLevelGroup='" & cmb_Second.Text & "', " & _
                " SecondLevelGroup='" & cmb_third.Text & "',FirstLevelGroup='" & cmb_fourth.Text & "', " & _
                " AccountCode='" & txt_MainAccntCode.Text & "',ChildSeriesNumber='" & txt_SubAccntNo.Text & "', " & _
                " ChildAccountCode='" & SAC & "', " & _
                " AccountName='" & txt_main.Text & "',AccountNameFull='" & txt_sub.Text & "',FundType='" & cmb_Fund.Text & "' " & _
                " where FMISAccountCode=" & TmpFMISCode & ""
                
                MsgBox "Saving New Account, Successful!", vbInformation, "System Information"
                Call Clear
End Select

End Sub
Private Sub Clear()
'Label1.Caption = ""
'txt_main.Text = ""
'txt_sub.Text = ""
'txt_SubAccntNo.Text = ""
'txt_MainAccntCode.Text = ""

Call LoadGroupLevel(cmb_Fund.Text, "FirstLevelGroup", cmb_fourth)
Call LoadGroupLevel(cmb_Fund.Text, "SecondLevelGroup", cmb_third)
Call LoadGroupLevel(cmb_Fund.Text, "ThirdLevelGroup", cmb_Second)
Call LoadGroupLevel(cmb_Fund.Text, "FourthLevelGroup", cmb_First)

End Sub
Private Sub FrameState(ByVal Enable As Boolean)

Frame1.Enabled = Enable
Frame2.Enabled = Enable
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set frmChartOfAccounts = Nothing
End Sub
Private Sub SetGrid()
MSHFlexGrid1.ColWidth(0) = 1000 'FMISAccountCode
MSHFlexGrid1.ColWidth(1) = 0 'FourthLevelGroup
MSHFlexGrid1.ColWidth(2) = 0 'ThirdLevelGroup
MSHFlexGrid1.ColWidth(3) = 0 'SecondLevelGroup
MSHFlexGrid1.ColWidth(4) = 0 'FirstLevelGroup
MSHFlexGrid1.ColWidth(5) = 0 'AccountCode
MSHFlexGrid1.ColWidth(6) = 0 'ChildSeriesNumber
MSHFlexGrid1.ColWidth(7) = 3000 'ChildAccountCode
MSHFlexGrid1.ColWidth(8) = 5500 'AccountName
MSHFlexGrid1.ColWidth(9) = 5000 'AccountNameFull
MSHFlexGrid1.ColWidth(10) = 2000 'FundType
MSHFlexGrid1.ColWidth(11) = 1000 'UserID
End Sub
Private Sub LoadBackToEdit(ByVal RowNo As Long)
cmb_First.ListIndex = GetIndex(cmb_First, MSHFlexGrid1.TextMatrix(RowNo, 1))
cmb_Second.ListIndex = GetIndex(cmb_Second, MSHFlexGrid1.TextMatrix(RowNo, 2))
cmb_third.ListIndex = GetIndex(cmb_third, MSHFlexGrid1.TextMatrix(RowNo, 3))
cmb_fourth.ListIndex = GetIndex(cmb_fourth, MSHFlexGrid1.TextMatrix(RowNo, 4))
txt_main.Text = MSHFlexGrid1.TextMatrix(RowNo, 8)
Label1.Caption = MSHFlexGrid1.TextMatrix(RowNo, 7)

If Trim(UCase(MSHFlexGrid1.TextMatrix(RowNo, 8))) = Trim(UCase(MSHFlexGrid1.TextMatrix(RowNo, 9))) Then
    txt_sub.Text = ""
Else
    txt_sub.Text = MSHFlexGrid1.TextMatrix(RowNo, 9)
End If

txt_SubAccntNo.Text = MSHFlexGrid1.TextMatrix(RowNo, 6)
txt_MainAccntCode.Text = MSHFlexGrid1.TextMatrix(RowNo, 5)
End Sub
Private Sub LoadFilteredAccnts(ByVal AccntNo As String, ByVal fund As String)
Dim opnaccnt As New ADODB.Recordset

opnaccnt.Open "Select FMISAccountCode as FMISID,FourthLevelGroup,ThirdLevelGroup,SecondLevelGroup,FirstLevelGroup,AccountCode,ChildSeriesNumber, " & _
    " ChildAccountCode as AccountCode,AccountName, " & _
      " AccountNameFull as SubAccountName,FundType as Fund,UserID from tblREF_AIS_ChartofAccounts " & _
      " where ChildAccountCode like '" & txt_AccntNo.Text & "%' and fundtype='" & fund & "' and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnaccnt.RecordCount <> 0 Then
    Set MSHFlexGrid1.DataSource = opnaccnt
    Call SetGrid
End If
opnaccnt.Close
Set opnaccnt = Nothing

End Sub






Private Sub MSHFlexGrid1_Click()
On Error GoTo bad
TmpRowNo = MSHFlexGrid1.Row
TmpFMISCode = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) 'FmisCode
Exit Sub
bad:
MsgBox err.Description, vbInformation, "System Message"
End Sub

Private Sub txt_AccntNo_Change()
Call LoadFilteredAccnts(txt_AccntNo.Text, cmb_Fund.Text)
End Sub
Private Sub LoadGroupLevel(ByVal fund As String, ByVal grouplevel As String, ByVal cmb As ComboBox)
Dim opnLevel As New ADODB.Recordset

cmb.Clear
opnLevel.Open "Select " & grouplevel & " as level from tblREF_AIS_ChartofAccounts where fundtype='" & fund & "' group by " & grouplevel & " order by " & grouplevel & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnLevel.RecordCount <> 0 Then
    Do Until opnLevel.EOF
        If IsNull(opnLevel!Level) = False Then
            cmb.AddItem (opnLevel!Level)
        End If
    opnLevel.MoveNext
    Loop
End If
opnLevel.Close
Set opnLevel = Nothing

End Sub
Private Sub txt_main_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_MainAccntCode.Text)) > 0 Then
        txt_MainAccntCode.SelStart = 0
        txt_MainAccntCode.SelLength = Len(txt_MainAccntCode.Text)
        txt_MainAccntCode.SetFocus
    Else
        txt_MainAccntCode.SetFocus
    End If
End If
End Sub
Private Sub txt_MainAccntCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_sub.Text)) > 0 Then
        txt_sub.SelStart = 0
        txt_sub.SelLength = Len(txt_sub.Text)
        txt_sub.SetFocus
    Else
        txt_sub.SetFocus
    End If
End If
End Sub

Private Sub txt_sub_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_SubAccntNo.Text)) > 0 Then
        txt_SubAccntNo.SelStart = 0
        txt_SubAccntNo.SelLength = Len(txt_SubAccntNo.Text)
        txt_SubAccntNo.SetFocus
    Else
        txt_SubAccntNo.SetFocus
    End If
End If
End Sub

Private Sub txt_SubAccntNo_LostFocus()
If Left(txt_SubAccntNo.Text, 1) = "-" Then
  
End If
End Sub
