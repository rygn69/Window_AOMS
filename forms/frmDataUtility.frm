VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDataUtility 
   Caption         =   "Dbase Utility"
   ClientHeight    =   10050
   ClientLeft      =   840
   ClientTop       =   2640
   ClientWidth     =   15180
   Icon            =   "frmDataUtility.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   15180
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3495
      Top             =   2685
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   195
      TabIndex        =   44
      Top             =   6555
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox pic_UpdateClaimantCode 
      Height          =   210
      Left            =   3735
      ScaleHeight     =   150
      ScaleWidth      =   11190
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   11250
      Begin VB.TextBox txt_REplace 
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
         Left            =   1995
         TabIndex        =   36
         Top             =   1035
         Width           =   1755
      End
      Begin VB.TextBox txt_OrigPrefixCode 
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
         Left            =   1995
         TabIndex        =   34
         Top             =   405
         Width           =   1755
      End
      Begin VB.CommandButton cmd_close 
         Caption         =   "Close"
         Height          =   525
         Left            =   9570
         TabIndex        =   32
         Top             =   150
         Width           =   1485
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "Update"
         Height          =   525
         Left            =   9570
         TabIndex        =   31
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lbl_Percent 
         Caption         =   "Label15"
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
         Height          =   285
         Left            =   6060
         TabIndex        =   43
         Top             =   1155
         Width           =   3000
      End
      Begin VB.Label lbl_table 
         Caption         =   "Label15"
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
         Height          =   285
         Left            =   6060
         TabIndex        =   42
         Top             =   810
         Width           =   3000
      End
      Begin VB.Label lbl_action 
         Caption         =   "Label15"
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
         Height          =   285
         Left            =   6060
         TabIndex        =   41
         Top             =   465
         Width           =   3000
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage Completion :"
         Height          =   195
         Left            =   4200
         TabIndex        =   40
         Top             =   1185
         Width           =   1740
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Table :"
         Height          =   195
         Left            =   4200
         TabIndex        =   39
         Top             =   832
         Width           =   1005
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recent Action :"
         Height          =   195
         Left            =   4200
         TabIndex        =   38
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace Prefix Code"
         Height          =   195
         Left            =   375
         TabIndex        =   37
         Top             =   1155
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Original Prefix Code"
         Height          =   195
         Left            =   375
         TabIndex        =   35
         Top             =   525
         Width           =   1380
      End
      Begin VB.Shape Shape5 
         Height          =   1455
         Left            =   210
         Top             =   255
         Width           =   9075
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   195
      TabIndex        =   27
      Top             =   8940
      Width           =   3270
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   135
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   375
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SQL Statement"
      Height          =   1020
      Left            =   180
      TabIndex        =   24
      Top             =   7935
      Width           =   3285
      Begin VB.CommandButton Command1 
         Caption         =   "Formulate"
         Height          =   450
         Left            =   240
         TabIndex        =   26
         Top             =   375
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Execute"
         Height          =   450
         Left            =   1725
         TabIndex        =   25
         Top             =   375
         Width           =   1335
      End
   End
   Begin VB.CheckBox chk_tables 
      Caption         =   "Table Names"
      Height          =   195
      Left            =   225
      TabIndex        =   22
      Top             =   4020
      Width           =   1665
   End
   Begin VB.CheckBox chk_Field 
      Caption         =   "Field Parameters"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   6900
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6510
      Left            =   3720
      ScaleHeight     =   6480
      ScaleWidth      =   11235
      TabIndex        =   18
      Top             =   3105
      Visible         =   0   'False
      Width           =   11265
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1170
         TabIndex        =   20
         Top             =   1500
         Visible         =   0   'False
         Width           =   2040
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6480
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   11430
         _Version        =   393216
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SQL Statement"
      Height          =   765
      Left            =   3735
      TabIndex        =   17
      Top             =   2190
      Width           =   11250
      Begin VB.TextBox txt_sql 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   105
         TabIndex        =   23
         Top             =   270
         Width           =   11040
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      TabIndex        =   16
      Top             =   7140
      Width           =   3300
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3690
      Top             =   -90
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Use Parameter of :"
      Height          =   1755
      Left            =   195
      TabIndex        =   11
      Top             =   2100
      Width           =   3300
      Begin VB.OptionButton opn_Controlno 
         Caption         =   "Control / Voucher No."
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   690
         Width           =   2070
      End
      Begin VB.OptionButton opn_checkno 
         Caption         =   "Check No."
         Height          =   195
         Left            =   330
         TabIndex        =   14
         Top             =   1395
         Width           =   1245
      End
      Begin VB.OptionButton opn_fmisno 
         Caption         =   "Record ID / FMIS No."
         Height          =   195
         Left            =   330
         TabIndex        =   13
         Top             =   1050
         Width           =   2025
      End
      Begin VB.OptionButton opn_Reviewno 
         Caption         =   "Review / Rec No."
         Height          =   195
         Left            =   330
         TabIndex        =   12
         Top             =   345
         Width           =   1875
      End
   End
   Begin VB.TextBox txt_CheckNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   9510
      TabIndex        =   9
      Top             =   1013
      Width           =   5250
   End
   Begin VB.TextBox txt_FmisNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   6735
      TabIndex        =   7
      Top             =   1013
      Width           =   2610
   End
   Begin VB.TextBox txt_controlNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   3945
      TabIndex        =   2
      Top             =   1013
      Width           =   2610
   End
   Begin VB.TextBox txt_Reviewno 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   345
      TabIndex        =   1
      Top             =   660
      Width           =   2955
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   4260
      Width           =   3270
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "F12 - For Mass Update of Claimant Code"
      Height          =   210
      Left            =   11805
      TabIndex        =   33
      Top             =   9660
      Width           =   3045
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check No."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9480
      TabIndex        =   10
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record ID / FMIS No."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6735
      TabIndex        =   8
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control / Voucher No."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3930
      TabIndex        =   6
      Top             =   750
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Review / Rec No."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   375
      TabIndex        =   5
      Top             =   405
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   390
      TabIndex        =   4
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please fill-up this first among others."
      Height          =   420
      Left            =   630
      TabIndex        =   3
      Top             =   1485
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   930
      Left            =   195
      Top             =   30
      Width           =   3300
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   3750
      Top             =   450
      Width           =   11220
   End
   Begin VB.Shape Shape3 
      Height          =   1530
      Left            =   195
      Top             =   495
      Width           =   3300
   End
   Begin VB.Shape Shape4 
      Height          =   1530
      Left            =   3750
      Top             =   30
      Width           =   11220
   End
End
Attribute VB_Name = "frmDataUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MSSQL_SECURE_LOGIN = True   'login type (True for NT security)
Const MSSQL_LOGIN_NAME = ""       'login name (for NT security use "" here)
Const MSSQL_PASSWORD = ""         'password   (for NT security use "" here)



Public dmoSrv    'As New SQLDMO.SQLServer    'SQLDMO Server object


Dim tmpData As Variant

Private Sub cmd_close_Click()
lbl_action.Caption = ""
lbl_table.Caption = ""
lbl_Percent.Caption = ""
txt_OrigPrefixCode.Text = ""
txt_REplace.Text = ""
pic_UpdateClaimantCode.Visible = False
'List1.Style = 0
End Sub

Private Sub cmd_update_Click()
Dim cc, xx As Integer

For cc = 0 To List1.ListCount - 1
    DoEvents
    If List1.Selected(cc) = True Then
        lbl_action.Caption = "Updating Claimant Codes..."
        lbl_table.Caption = List1.List(cc)
        For xx = 1 To MSHFlexGrid1.Rows - 1
            DoEvents
            lbl_Percent.Caption = CInt((xx / (MSHFlexGrid1.Rows - 1)) * 100) & " % Completed.... " & GetNewCode(MSHFlexGrid1.TextMatrix(xx, 2), txt_OrigPrefixCode.Text, txt_REplace.Text)
            Call UpdateClaimantCode(MSHFlexGrid1.TextMatrix(xx, 2), GetNewCode(MSHFlexGrid1.TextMatrix(xx, 2), txt_OrigPrefixCode.Text, txt_REplace.Text), List1.List(cc), "ClaimantCode")
        Next xx
    End If
Next cc
MsgBox "Updating Claimant Codes... Successful!", vbInformation, "System Information"
lbl_action.Caption = ""
lbl_table.Caption = ""
lbl_Percent.Caption = ""
End Sub
Private Function GetNewCode(ByVal OrigCode As String, ByVal OrigPrefixCode As String, ByVal NewPrefixCode As String) As String
GetNewCode = NewPrefixCode & Mid(OrigCode, Len(OrigPrefixCode) + 1, Len(OrigCode) - Len(OrigPrefixCode))

End Function
Private Sub UpdateClaimantCode(ByVal OrigCode As String, ByVal NewCode As String, ByVal TableName As String, ByVal Fieldname As String)
opndbaseFMIS.Execute "Update " & TableName & " set " & Fieldname & " = '" & NewCode & "' where " & Fieldname & " = '" & OrigCode & "'"
End Sub
Private Sub Command1_Click()
Call FormulateSQL
End Sub
Private Function OpnSelected() As Boolean
If opn_Reviewno.Value = True Or opn_Controlno.Value = True Or opn_fmisno.Value = True Or opn_checkno.Value = True Then
    OpnSelected = True
Else
    OpnSelected = False
End If
End Function
Private Function ReformatMultipleCheckPosition(ByVal CheckSources As String) As String
Dim tmpText As Variant
Dim cc As Integer

If InStr(CheckSources, ",") <> 0 Then
    tmpText = Split(CheckSources, ",")
    For cc = 0 To UBound(tmpText)
        If Len(ReformatMultipleCheckPosition) = 0 Then
            ReformatMultipleCheckPosition = "'" & tmpText(cc) & "'"
        Else
            ReformatMultipleCheckPosition = ReformatMultipleCheckPosition & ",'" & tmpText(cc) & "'"
        End If
    Next cc
End If

If Len(ReformatMultipleCheckPosition) <> 0 Then
    ReformatMultipleCheckPosition = "(" & ReformatMultipleCheckPosition & ")"
End If
End Function

Private Sub FormulateSQL()
If chk_tables.Value = 1 And chk_Field.Value = 1 Then
    If OpnSelected = True Then
        If Len(Trim(txt_Reviewno.Text)) <> 0 And Len(Trim(txt_controlNo.Text)) <> 0 Then
                If opn_Reviewno.Value = True Then
                    txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " where " & List2.List(List2.ListIndex) & "='" & txt_Reviewno.Text & "'" & " order by trnno"
                ElseIf opn_Controlno.Value = True Then
                    txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " where " & List2.List(List2.ListIndex) & "='" & txt_controlNo.Text & "'" & " order by trnno"
                ElseIf opn_fmisno.Value = True Then
                    txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " where " & List2.List(List2.ListIndex) & "='" & txt_FmisNo.Text & "'" & " order by trnno"
                ElseIf opn_checkno.Value = True Then
                    If InStr(txt_CheckNo.Text, ",") <> 0 Then
                        txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " where " & List2.List(List2.ListIndex) & " in " & ReformatMultipleCheckPosition(txt_CheckNo.Text) & " order by trnno"
                    Else
                        txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " where " & List2.List(List2.ListIndex) & "='" & txt_CheckNo.Text & "'" & " order by trnno"
                    End If
                End If
        Else
            MsgBox "There is no Review No. or Control No. being specified!", vbInformation, "System Information"
            If Len(Trim(txt_Reviewno.Text)) <> 0 Then
                txt_Reviewno.SelStart = 0
                txt_Reviewno.SelLength = Len(txt_Reviewno.Text)
                txt_Reviewno.SetFocus
            Else
                txt_Reviewno.SetFocus
            End If
        End If
    Else
        MsgBox "Please select type of Parameter to be used in the SQL Statement!", vbInformation, "System Information"
    End If
ElseIf chk_tables.Value = 1 And chk_Field.Value = 0 Then
    txt_sql.Text = "Select * from " & List1.List(List1.ListIndex) & " order by trnno"
ElseIf chk_tables.Value = 0 And chk_Field.Value = 1 Then
    MsgBox "To formulate a valid SQL Statement" & Chr(13) & "you have to Check the table First" & Chr(13) & "Then Select what table you like to open!", vbInformation, "System Information"
ElseIf chk_tables.Value = 0 And chk_Field.Value = 0 Then
    MsgBox "Unable to formulate SQL Statement!", vbInformation, "System Information"
End If
End Sub
Private Sub Command2_Click()
On Error GoTo handler

Picture1.Visible = True
MSHFlexGrid1.Clear


If InStr(txt_sql.Text, "Select") <> 0 Then
    If opndbaseFMIS.Execute(txt_sql.Text).RecordCount <> 0 Then
        Set MSHFlexGrid1.Recordset = opndbaseFMIS.Execute(txt_sql.Text)
        MSHFlexGrid1.ColWidth(0) = 300
        MSHFlexGrid1.Refresh
        Text1.Visible = False
        Label1.Caption = "Actual No. of Record/s Found : " & Format(opndbaseFMIS.Execute(txt_sql.Text).RecordCount, "###,##0")
        Label8.Caption = "Max. no of Record/s Shown    : " & Format(MSHFlexGrid1.Rows - 1, "###,##0")
    Else
        MSHFlexGrid1.Rows = 2
    End If

Else
    opndbaseFMIS.Execute (txt_sql.Text)
    MsgBox "Sql Statement Successfuly Executed!", vbInformation, "System Information"
    txt_sql.Text = ""
End If

handler:
If err.Number <> 0 Then
    MsgBox err.Description
    Picture1.Visible = False
    If Len(Trim(txt_sql.Text)) <> 0 Then
        txt_sql.SelStart = 0
        txt_sql.SelLength = Len(txt_sql.Text)
        txt_sql.SetFocus
    Else
        txt_sql.SetFocus
    End If
End If
End Sub
Private Function getFieldType(ByVal Fieldname As String, ByVal TableName As String) As String
Dim opntable As New ADODB.Recordset

opntable.Open "Select * from " & TableName & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    getFieldType = getGeneralDataType(opntable(Fieldname).Type)
opntable.Close
Set opntable = Nothing
End Function






Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
ElseIf KeyCode = vbKeyF12 Then
    lbl_action.Caption = ""
    lbl_table.Caption = ""
    lbl_Percent.Caption = ""
    txt_OrigPrefixCode.Text = ""
    txt_REplace.Text = ""
    pic_UpdateClaimantCode.Visible = True
    'List1.Style = 1
End If
End Sub

Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = ((Screen.Height - Me.Height) / 2)
WindowsXPC1.InitSubClassing
Label1.Caption = ""
Label8.Caption = ""
Timer1.Enabled = True
End Sub
Private Function getGeneralDataType(ByVal DataType As Integer) As String
Select Case DataType

Case 0, 2, 3, 4, 5, 6, 20, 14, 16, 17, 18, 19, 21, 131, 128, 139, 204, 205, 201
    getGeneralDataType = "Number"
Case 11
    getGeneralDataType = "Boolean"
Case 10, 132, 12, 9, 13, 72, 8, 129, 200, 130, 202, 203, 136, 138
    getGeneralDataType = "String"
Case 7, 133, 134, 135, 64
    getGeneralDataType = "DateNTime"
End Select

End Function
Private Sub ClearOptions()
opn_Reviewno.Value = False
opn_Controlno.Value = False
opn_fmisno.Value = False
opn_checkno.Value = False
End Sub
Private Sub ClearCheckBoxes()
chk_tables.Value = 0
chk_Field.Value = 0
End Sub
Private Sub ClearAllKeys()
txt_controlNo.Text = ""
txt_FmisNo.Text = ""
txt_CheckNo.Text = ""
End Sub
Private Sub LoadAllKeys(ByVal Reviewno As String)
txt_controlNo.Text = getControlNo(Reviewno)
txt_FmisNo.Text = GetFMISNo(txt_controlNo.Text)
txt_CheckNo.Text = GetCheckNo(GettingFinalMixcode)
End Sub
Private Function GettingFinalMixcode() As String
If Len(Trim(txt_FmisNo.Text)) <> 0 Then
    GettingFinalMixcode = txt_FmisNo.Text
Else
    GettingFinalMixcode = txt_controlNo.Text
End If
End Function
Private Function GetFMISNo(ByVal controlno As String) As String
Dim opnFMiSNo As New ADODB.Recordset

opnFMiSNo.Open "select RecordID from tblCMS_CDVoucherPreparation4CA where controlno='" & controlno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnFMiSNo.RecordCount <> 0 Then
    GetFMISNo = opnFMiSNo!RecordID
End If
opnFMiSNo.Close
Set opnFMiSNo = Nothing
End Function
Private Function GetCheckNo(ByVal MixCode As String) As String
Dim opnCheckNo As New ADODB.Recordset

opnCheckNo.Open "Select checkno from tblCMS_CDPreparedCheck where mixcode='" & MixCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnCheckNo.RecordCount <> 0 Then
    Do Until opnCheckNo.EOF
        If Len(GetCheckNo) = 0 Then
            GetCheckNo = opnCheckNo!checkno
        Else
            GetCheckNo = GetCheckNo & "," & opnCheckNo!checkno
        End If
        opnCheckNo.MoveNext
    Loop
End If
opnCheckNo.Close
Set opnCheckNo = Nothing
End Function
Private Function getControlNo(ByVal Reviewno As String) As String
Dim opnControlNo As New ADODB.Recordset

opnControlNo.Open "Select VoucherNo from tblCMS_EXCashVerification where recno='" & Reviewno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnControlNo.RecordCount <> 0 Then
    getControlNo = IIf(IsNull(opnControlNo!VoucherNo), "", opnControlNo!VoucherNo)
End If
opnControlNo.Close
Set opnControlNo = Nothing

End Function
Private Sub LoadFieldCriteria(ByVal TableName As String)
Dim tmpField As Variant
Dim cc As Integer

tmpField = readTXTDATA(TableName, "Params", App.Path & "\data\Attachment.ini")
tmpField = Split(tmpField, ",")

List2.Clear
For cc = 0 To UBound(tmpField)
    List2.AddItem (tmpField(cc))
Next cc
List2.ListIndex = 0 'Setting as default
End Sub


Private Sub Form_Unload(Cancel As Integer)
WindowsXPC1.EndWinXPCSubClassing
Set frmDataUtility = Nothing
End Sub

Private Sub List1_Click()
Call LoadFieldCriteria(List1.List(List1.ListIndex))
End Sub

Private Sub MSHFlexGrid1_Click()
Text1.Text = ""
Text1.Visible = False
MSHFlexGrid1.Refresh
End Sub

Private Sub MSHFlexGrid1_DblClick()
Select Case MSHFlexGrid1.col
    Case 0, 1 'No action
        Text1.Text = ""
        Text1.Visible = False
    Case Else
        tmpData = MSHFlexGrid1.Text 'For purposes of comparing between the Old Date against the New Data.......
        Text1.Move MSHFlexGrid1.CellLeft, MSHFlexGrid1.CellTop, MSHFlexGrid1.CellWidth, MSHFlexGrid1.CellHeight
        Text1.Visible = True
        If Len(Trim(MSHFlexGrid1.Text)) <> 0 Then
            Text1.Text = MSHFlexGrid1.Text
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
        Else
            Text1.SetFocus
        End If
End Select
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = vbKeyDelete Then
    If Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> 0 Then
        If MsgBox("Are you sure want to Delete this Record?", vbQuestion + vbYesNo, "System Query Confirmation") = vbYes Then
            opndbaseFMIS.Execute "Delete from " & List1.List(List1.ListIndex) & " where trnno=" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & ""
            MsgBox "Delete Procedure Succesful!", vbInformation, "System Information"
            Command2_Click 'executing updates---------------
        End If
    End If
End If
End Sub

Private Sub MShFlexGrid1_Scroll()
Text1.Text = ""
Text1.Visible = False
End Sub
Private Function CompleteEntries(ByVal TargetRowNo As Integer) As Boolean
Dim cc As Integer

For cc = 2 To MSHFlexGrid1.Cols - 1
    If Len(Trim(MSHFlexGrid1.TextMatrix(TargetRowNo, cc))) = 0 Then
        CompleteEntries = False
        Exit Function
    Else
        CompleteEntries = True
    End If
Next cc

End Function
Private Function getFormulateSavingSQL(ByVal TableName As String, ByVal TargetRowNo As Integer) As String
Dim cc As Integer
Dim tmpFlds, tmpValues As String

For cc = 2 To MSHFlexGrid1.Cols - 1

    If Len(tmpFlds) <> 0 Then 'Formulating the Field Names of the Table------------------------\
        tmpFlds = tmpFlds & "," & MSHFlexGrid1.TextMatrix(0, cc)
    Else
        tmpFlds = MSHFlexGrid1.TextMatrix(0, cc)
    End If '-----------------------------------------------------------------------------------\

    'Formulation of Field Values of the Table-----------------------------------/
    Select Case getFieldType(MSHFlexGrid1.TextMatrix(0, cc), TableName)
        Case "Number"
            If Len(tmpValues) <> 0 Then
                tmpValues = tmpValues & "," & MSHFlexGrid1.TextMatrix(TargetRowNo, cc)
            Else
                tmpValues = MSHFlexGrid1.TextMatrix(TargetRowNo, cc)
            End If
        Case "Boolean"
            If Len(tmpValues) <> 0 Then
                tmpValues = tmpValues & "," & IIf(MSHFlexGrid1.TextMatrix(TargetRowNo, cc), 1, 0)
            Else
                tmpValues = " & IIf(MSHFlexGrid1.TextMatrix(TargetRowNo, cc), 1, 0) & "
            End If
        Case "String", "DateNTime"
            If Len(tmpValues) <> 0 Then
                tmpValues = tmpValues & ",'" & MSHFlexGrid1.TextMatrix(TargetRowNo, cc) & "'"
            Else
                tmpValues = "'" & MSHFlexGrid1.TextMatrix(TargetRowNo, cc) & "'"
            End If
    End Select
    '----------------------------------------------------------------------------/
Next cc

getFormulateSavingSQL = "Insert into " & TableName & " (" & tmpFlds & ") " & " values (" & tmpValues & ")"

End Function
Private Sub SaveUpdate(ByVal Trnno As Long, ByVal Colno As Integer, ByVal FldName As String, ByVal TableName As String, ByVal FldData As Variant)

Select Case getFieldType(MSHFlexGrid1.TextMatrix(0, Colno), TableName)
    Case "Number"
        opndbaseFMIS.Execute "update " & TableName & " set " & FldName & "=" & FldData & " where trnno=" & Trnno & ""
    Case "Boolean"
        opndbaseFMIS.Execute "update " & TableName & " set " & FldName & "=" & IIf(FldData, 1, 0) & " where trnno=" & Trnno & ""
    Case "String"
        opndbaseFMIS.Execute "update " & TableName & " set " & FldName & "='" & FldData & "' where trnno=" & Trnno & ""
    Case "DateNTime"
        opndbaseFMIS.Execute "update " & TableName & " set " & FldName & "='" & FldData & "' where trnno=" & Trnno & ""
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handler

If KeyCode = 13 Then
    Select Case MSHFlexGrid1.col
        Case 0, 1 ' no action
        Case Else
            MSHFlexGrid1.Text = Text1.Text
            Text1.Text = ""
            Text1.Visible = False
            
            If MSHFlexGrid1.col < MSHFlexGrid1.Cols - 1 Then
                'Saving Procedure will be set here...........
                    If Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> 0 Then 'Only for Edition for Existing Record---
                        If tmpData <> MSHFlexGrid1.Text Then
                            If MsgBox("Want to Save changes?", vbQuestion + vbYesNo, "Sysm Confirmation Query") = vbYes Then
                                Call SaveUpdate(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1), MSHFlexGrid1.col, MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.col), List1.List(List1.ListIndex), MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col))
                            End If
                        End If
                    End If
                '............................................
                MSHFlexGrid1.col = MSHFlexGrid1.col + 1
                MSHFlexGrid1.Row = MSHFlexGrid1.Row
                MSHFlexGrid1_DblClick
                
            ElseIf MSHFlexGrid1.col = MSHFlexGrid1.Cols - 1 Then
                    If MSHFlexGrid1.Row = MSHFlexGrid1.Rows - 1 Then
                            'Saving Procedure will be set here...........
                            If Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> 0 Then 'Only for Edition for Existing Record---
                                If tmpData <> MSHFlexGrid1.Text Then
                                    If MsgBox("Want to Save changes?", vbQuestion + vbYesNo, "Sysm Confirmation Query") = vbYes Then
                                        Call SaveUpdate(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1), MSHFlexGrid1.col, MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.col), List1.List(List1.ListIndex), MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col))
                                    End If
                                End If
                                
                                Text1.Text = ""
                                Text1.Visible = False
                                
                                If MsgBox("Want to Add New Record?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
                                    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                                    MSHFlexGrid1.col = 2
                                    MSHFlexGrid1.Row = MSHFlexGrid1.Rows - 1
                                    MSHFlexGrid1_DblClick
                                End If
                                
                            Else 'For New Record---------------------
                                If MsgBox("Want to save Newly Created Record?", vbInformation + vbYesNo, "System Confirmation Query") = vbYes Then
                                    'If CompleteEntries(MSHFlexGrid1.Rows - 1) = True Then
                                            opndbaseFMIS.Execute getFormulateSavingSQL(List1.List(List1.ListIndex), MSHFlexGrid1.Rows - 1)
                                            If MsgBox("Saving of New Record Successful!" & Chr(13) & Chr(13) & "Want to Add New Record?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
                                                Text1.Text = ""
                                                Text1.Visible = False
                                                Command2_Click
                                                
                                                If MsgBox("Want to Add New Record?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
                                                    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                                                    MSHFlexGrid1.col = 2
                                                    MSHFlexGrid1.Row = MSHFlexGrid1.Rows - 1
                                                    MSHFlexGrid1_DblClick
                                                End If
                                            End If
                                    'Else
                                    '    MsgBox "If you want to save your Newly Created Record," & Chr(13) & "Please complete your entries!", vbCritical, "Sytem Information"
                                    'End If
                                Else
                                    Text1.Text = ""
                                    Text1.Visible = False
                                    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
                                End If
                            End If
                            '............................................
                            
                    Else
                            'Saving Procedure will be set here...........
                            If Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> 0 Then 'Only for Edition for Existing Record---
                                If tmpData <> MSHFlexGrid1.Text Then
                                    If MsgBox("Want to Save changes?", vbQuestion + vbYesNo, "Sysm Confirmation Query") = vbYes Then
                                        Call SaveUpdate(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1), MSHFlexGrid1.col, MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.col), List1.List(List1.ListIndex), MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col))
                                    End If
                                End If
                            End If
                            '............................................
                            Text1.Text = ""
                            Text1.Visible = False
                    End If
            End If
    End Select
End If

handler:
If err.Number <> 0 Then
    MsgBox err.Description
End If
End Sub

Private Sub Timer1_Timer()
Call ClearCheckBoxes
Call ClearOptions
Call LoadTableNames(List1)
Timer1.Enabled = False
End Sub

Private Sub txt_OrigPrefixCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_REplace.Text)) <> 0 Then
        txt_REplace.SelStart = 0
        txt_REplace.SelLength = Len(txt_REplace.Text)
        txt_REplace.SetFocus
    Else
        txt_REplace.SetFocus
    End If
End If
End Sub

Private Sub txt_Reviewno_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Call ClearAllKeys
    If Len(Trim(txt_Reviewno.Text)) <> 0 Then
        Call LoadAllKeys(txt_Reviewno.Text)
    End If
End If
End Sub

