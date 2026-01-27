VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form f_COA_SL_sub 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "f_COA_SL_sub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   6840
      Top             =   0
   End
   Begin VB.TextBox txtcode 
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
      Left            =   1560
      TabIndex        =   6
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtadd 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   6960
      Width           =   6375
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   8655
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Import"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":4274
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accountcode"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Accountname"
         Object.Width           =   11758
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdupdate 
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   7680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "&Update"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":7D7E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3600
      Top             =   1440
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
      TextControl     =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   7680
      TabIndex        =   10
      Top             =   7680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Back"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":8DD0
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":C8DA
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   7680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "&Hide"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":103E4
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "&Back"
      CapAlign        =   2
      BackStyle       =   4
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":13EEE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H7 
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   7680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "&UnHide"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_sub.frx":179F8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label lbl_ChartAccountChildID 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "New"
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
      Left            =   6360
      TabIndex        =   19
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Add Others"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lbldeselect 
      Caption         =   "Deselect All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblselect 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label blnkStat 
      Caption         =   "New"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   6570
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   6550
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Code    Explaination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   6550
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8760
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label lblstat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "f_COA_SL_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Address, accountcode, accntcode, Field1, Field2, Condition, Accountname, childcode As String
Public ChartAccountChildID, code As Long
Public col, levelno As Integer
Public IfEdited As Boolean
Public Function Get_Account_by_ChartAccountParentID(ByVal AccountChildParentID As Long, ByVal Order As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select [ChartAccountChildID],[code],AccountChildName from [Accounting].[tbl_l_ChartOfAccountsChild] where AccountChildParentID = " & AccountChildParentID & " " & Order & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    LstAccountcode.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                    Set x = LstAccountcode.ListItems.Add(, , rec.Fields!ChartAccountChildID)
                    x.SubItems(1) = Trim(rec.Fields!code)
                    x.SubItems(2) = Trim(rec.Fields!AccountChildName)
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
End Function
Private Sub cmdupdate_Click()
If Trim(txtcode.Text) = "" And Trim(txtadd.Text) = "" Then
    MsgBox "Please Complete the parameter..!", vbInformation, "System Message"
    Exit Sub
End If
If lbl_ChartAccountChildID.Caption = "" Then
MsgBox "Please select item on the list..!", vbInformation, "System Message"
    Exit Sub
End If
'If ExecFunction("SELECT [fmis].[dbo].[MPfunc_CheckIfAccntIsUse] ('" & Trim(Me.Caption) & "-" & Trim(txtcode.Text) & "' )") = 1 Then
'MsgBox , vbCritical, "Error Updating Accounts"
'MsgBox "Cannot Update " & txtcode.Text & "-" & txtadd.Text & ": Access is denied." & vbNewLine & vbNewLine & "Make sure that Accounts is not currently in Use.", vbCritical, "Error Updating Accounts"
'End If
    If MsgBox("Accountcode " & txtcode.Text & "-" & txtadd.Text & " is have Sub Account..!" & vbNewLine & "Are you sure do you want to Update?", vbCritical + vbYesNo, "System Message") = vbYes Then
        opndbaseFMIS.Execute "update [Accounting].[tbl_l_ChartOfAccountsChild] set AccountChildName = '" & txtadd.Text & "'  where ChartAccountChildID = '" & lbl_ChartAccountChildID.Caption & "' "
        MsgBox "Successfully Updated", vbInformation, "System Message"
        Call Form_Load
    End If
End Sub

Private Sub Form_Load()
txtaddress.Text = Accountname
Call Get_Account_by_ChartAccountParentID(ChartAccountChildID, "")
Me.Caption = childcode
Label5.Caption = LstAccountcode.ListItems.Count & " Records Found"
End Sub


Private Sub Label4_Click()
txtadd.Text = "Others"
txtcode.Text = 0
Call lvButtons_H3_Click
End Sub

Private Sub Lblselect_Click()
Dim x As Long
For x = 1 To LstAccountcode.ListItems.Count
    LstAccountcode.ListItems(x).Checked = True
    DoEvents
Next
End Sub

Private Sub lbldeselect_Click()
Dim x As Long
For x = 1 To LstAccountcode.ListItems.Count
    LstAccountcode.ListItems(x).Checked = False
    DoEvents
Next
End Sub



Private Sub LstAccountcode_Click()
If LstAccountcode.ListItems.Count <> 0 Then
    lbl_ChartAccountChildID.Caption = LstAccountcode.SelectedItem.Text
    txtcode.Text = LstAccountcode.SelectedItem.SubItems(1)
    txtadd.Text = LstAccountcode.SelectedItem.SubItems(2)
    blnkStat.Caption = "EDIT"
Else
    blnkStat.Caption = "NEW"
    txtcode.Text = ""
    txtadd.Text = ""
    lbl_ChartAccountChildID.Caption = ""
End If

End Sub

Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader
Case "Accountcode"
Call Get_Account_by_ChartAccountParentID(ChartAccountChildID, "order by code")
Case "Accountname"
Call Get_Account_by_ChartAccountParentID(ChartAccountChildID, "order by AccountChildName")
End Select
End Sub

Private Sub LstAccountcode_DblClick()
If LstAccountcode.ListItems.Count > 0 Then
Dim Newform As New f_COA_SL_sub
    With Newform
    .Accountname = Accountname & " ~ " & Trim(LstAccountcode.SelectedItem.SubItems(2))
    .levelno = levelno + 1
    .ChartAccountChildID = Trim(LstAccountcode.SelectedItem.Text)
    .code = Trim(LstAccountcode.SelectedItem.SubItems(1))
    .childcode = childcode & "-" & Trim(LstAccountcode.SelectedItem.SubItems(1))
    .Show 1
    End With
End If
End Sub

Private Sub lvButtons_H1_Click()
Dim Subcode As String
Dim x As Long
Dim isdel As Boolean
    Subcode = "Subcode" & col
    isdel = False
    
    
    
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
                isdel = True
            End If
        Next x
    If isdel = False Then
        MsgBox "Please Select/Check first the list above that you want to delete.", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Are you sure do you want to Delete the checked Account?", vbCritical + vbYesNo, "System Information") = vbYes Then
  
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
'                If ExecFunction("SELECT [fmis].[dbo].[MPfunc_CheckIfAccntIsUse] ('" & Trim(Me.Caption) & "-" & Trim(LstAccountcode.ListItems(x).Text) & "' )") = 0 Then
'
                lblstat.Caption = "Deleting Account " & LstAccountcode.ListItems(x).Text
                    opndbaseFMIS.Execute "DELETE FROM [Accounting].[tbl_l_ChartOfAccountsChild] where [ChartAccountChildID] = " & Trim(LstAccountcode.ListItems(x).Text) & ""
'                Else
'                    If MsgBox("Cannot Delete " & Trim(LstAccountcode.ListItems(x).ListSubItems(1).Text) & "-" & Trim(LstAccountcode.ListItems(x).ListSubItems(2).Text) & ": Access is denied." & vbNewLine & vbNewLine & "Make sure that Accounts is not currently in Use." & vbNewLine & "Yes = Ignore" & vbNewLine & "No = Cancel Deleting", vbCritical + vbYesNo, "Error Deleting Accounts") = vbYes Then
'                        'no action
'                    Else
'                        Exit For
'                    End If
'                End If
            End If
                DoEvents
        Next x
        lblstat.Caption = ""
        Call Form_Load
    End If
End Sub
Private Function IfCodeIsUses(ByVal id As Double) As Boolean
Dim rec As New ADODB.Recordset
IfCodeIsUses = False
rec.Open "Select top 1 id from tblAMIS_FinalJEV where id = " & id & ""
    If rec.RecordCount > 0 Then
        IfCodeIsUses = True
    End If
rec.Close
End Function
Private Sub lvButtons_H2_Click()
With f_COA_SL_import
    .Accountname = Accountname
    .levelno = levelno
    .ChartAccountChildID = ChartAccountChildID
    .code = code
    .childcode = childcode
    .Show 1
    Call Form_Load
End With
End Sub
Private Sub lvButtons_H3_Click()
If Trim(txtcode.Text) = "" And Trim(txtadd.Text) = "" Then
    MsgBox "Please Complete the parameter..!", vbInformation, "System Message"
    Exit Sub
End If

If IfExistname(txtadd.Text) = True Then
    MsgBox "Name Already Exist on the database..!", vbInformation, "System Message"
    Exit Sub
End If

If IfExistcode(txtcode.Text) = True Then
    MsgBox "Code Already Exist on the database..!", vbInformation, "System Message"
    GoTo proceed
    Exit Sub
End If

If Trim(txtcode.Text) = "" Then
proceed:
    If MsgBox("System Generate Code for your entry, Do you want to proceed?", vbInformation + vbYesNo, "System Information") = vbYes Then
        txtcode.Text = Get_MaxValue_COA(ChartAccountChildID)
    Else
        MsgBox "Please Specify the code, to Proceed...!", vbInformation, "System Message"
        txtcode.SetFocus
    Exit Sub
    End If
End If
If MsgBox("Are you sure do want to save?", vbInformation + vbYesNo, "System Message") = vbYes Then
        opndbaseFMIS.Execute "exec Accounting.usp_Save_ChartOfAccountsChild @code = '" & txtcode.Text & "',@AccountChildParentID = '" & ChartAccountChildID & "' ,@AccountChildName= '" & Replace(txtadd.Text, "'", "''") & "',@parentchildcode = '" & childcode & "' ,@parentLevelno = '" & levelno & "',@ModifiedByID = '" & ActiveUserID & "'"
        Call Get_Account_by_ChartAccountParentID(ChartAccountChildID, "order by AccountChildName")
    txtadd.Text = ""
    txtcode.Text = ""
End If
End Sub
Public Function IfExistname(ByVal name As String) As Boolean
Dim x As Integer
IfExistname = False
    For x = 1 To LstAccountcode.ListItems.Count
        If UCase(name) = Trim(LstAccountcode.ListItems(x).SubItems(2)) Then
            IfExistname = True
            Exit For
        End If
    Next x
End Function
Public Function IfExistcode(ByVal code As String) As Boolean
Dim x As Integer
IfExistcode = False
    For x = 1 To LstAccountcode.ListItems.Count
        If Trim(code) = Trim(LstAccountcode.ListItems(x).SubItems(1)) Then
            IfExistcode = True
            Exit For
        End If
    Next x
End Function
Public Function GetmaxID()
Dim rec As New ADODB.Recordset
Dim sql As String
sql = "select Accounting.fn_getMaxValue_COA(" & code & ") as maxid"
rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        GetmaxID = IIf(IsNull(rec!maxid), 0, rec!maxid) + 1
    Else
        GetmaxID = 1
    End If
rec.Close
Set rec = Nothing
End Function


Private Sub lvButtons_H4_Click()
Unload Me
End Sub

Private Sub lvButtons_H5_Click()
Dim Subcode As String
Dim x As Long
Dim isdel As Boolean
    Subcode = "Subcode" & col
    isdel = False
    
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
                isdel = True
            End If
        Next x
    If isdel = False Then
        MsgBox "Please Select/Check first the list above that you want to HIDE.", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Are you sure do you want to HIDE the checked Account?", vbInformation + vbYesNo, "System Information") = vbYes Then
  
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
                
                lblstat.Caption = "Updating Account " & LstAccountcode.ListItems(x).Text
                    opndbaseFMIS.Execute "update [tblReff_CodeClassification] set actioncode = 1  where " & Getcondition & " and " & Subcode & " = '" & Trim(LstAccountcode.ListItems(x).Text) & "' and actioncode = 0 "
            End If
                DoEvents
        Next x
        lblstat.Caption = ""
        Call Form_Load
    End If
End Sub

Private Sub lvButtons_H7_Click()
Dim Subcode As String
Dim x As Long
Dim isdel As Boolean
    Subcode = "Subcode" & col
    isdel = False
    
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
                isdel = True
            End If
        Next x
    If isdel = False Then
        MsgBox "Please Select/Check first the list above that you want to HIDE.", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Are you sure do you want to UNHIDE the checked Account?", vbInformation + vbYesNo, "System Information") = vbYes Then
  
        For x = 1 To LstAccountcode.ListItems.Count
            If LstAccountcode.ListItems(x).Checked = True Then
                
                lblstat.Caption = "Updating Account " & LstAccountcode.ListItems(x).Text
                    opndbaseFMIS.Execute "update [tblReff_CodeClassification] set actioncode = 0  where " & Getcondition & " and " & Subcode & " = '" & Trim(LstAccountcode.ListItems(x).Text) & "' and actioncode = 1 "
            End If
                DoEvents
        Next x
        lblstat.Caption = ""
        Call Form_Load
    End If
End Sub

'Private Sub lvButtons_H5_Click()
'Dim x As Long
'If IsCopy = False Then
'
'        For x = 1 To LstAccountcode.ListItems.Count
'            If LstAccountcode.ListItems(x).Checked = True Then
'                IsCopy = True
'                lvButtons_H5.Caption = "Paste"
'            End If
'        Next x
'Else
'    If MsgBox("Are you sure do you want to paste here?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
'        lvButtons_H5.Caption = "Copy"
'    Else
'        lvButtons_H5.Caption = "Copy"
'    End If
'IsCopy = False
'End If
'End Sub

Private Sub txtadd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call lvButtons_H3_Click
End If
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtcode.Text = GetmaxID
End If
End Sub
