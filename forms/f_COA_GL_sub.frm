VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form f_COA_GL_sub 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "f_COA_GL_sub.frx":0000
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
      TabIndex        =   1
      Top             =   420
      Width           =   8655
   End
   Begin MSComctlLib.ListView LstAccountcode 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
         Text            =   "ChartAccountID"
         Object.Width           =   0
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
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   4800
      TabIndex        =   3
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
      Image           =   "f_COA_GL_sub.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   7680
      TabIndex        =   5
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
      Image           =   "f_COA_GL_sub.frx":4274
      ImgSize         =   24
      cBack           =   -2147483633
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
      TabIndex        =   4
      Top             =   120
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
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "f_COA_GL_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Address, accountcode, accntcode, Field1, Field2, Condition As String
Public ChartAccountID As Long
Public Subcode1, Subdesc1, Subcode2, Subdesc2, Subcode3, Subdesc3, Subcode4, Subdesc4, Subcode5, Subdesc5, Subcode6, Subdesc6, Subcode7, Subdesc7 As String
Public col As Integer
Public IfEdited As Boolean
Public Function Get_Account_by_ChartAccountParentID(ByVal ChartAccountParentID As Long, ByVal Order As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select [ChartAccountID],[Accountcode],Accountname from [Accounting].[tbl_l_ChartOfAccountsParent] where ChartAccountParentID = " & ChartAccountParentID & " " & Order & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    LstAccountcode.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                    Set x = LstAccountcode.ListItems.Add(, , rec.Fields!ChartAccountID)
                    x.SubItems(1) = Trim(rec.Fields!accountcode)
                    x.SubItems(2) = Trim(rec.Fields!Accountname)
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
'If ExecFunction("SELECT [fmis].[dbo].[MPfunc_CheckIfAccntIsUse] ('" & Trim(Me.Caption) & "-" & Trim(txtcode.Text) & "' )") = 1 Then
'MsgBox , vbCritical, "Error Updating Accounts"
'MsgBox "Cannot Update " & txtcode.Text & "-" & txtadd.Text & ": Access is denied." & vbNewLine & vbNewLine & "Make sure that Accounts is not currently in Use.", vbCritical, "Error Updating Accounts"
'End If
    If MsgBox("Accountcode " & txtcode.Text & "-" & txtadd.Text & " is have Sub Account..!" & vbNewLine & "Are you sure do you want to Update?", vbCritical + vbYesNo, "System Message") = vbYes Then
        opndbaseFMIS.Execute "update tblReff_CodeClassification set " & Field2 & " = '" & UCase(Trim(txtadd.Text)) & "',migrated = null  where  " & Getcondition & " and " & Field1 & " = '" & Trim(txtcode.Text) & "' "
        MsgBox "Successfully Updated", vbInformation, "System Message"
        Call Form_Load
    End If
End Sub

Private Sub Form_Load()
txtaddress.Text = Address
Field1 = "Subcode" & col
Field2 = "Subdesc" & col
Call Get_Account_by_ChartAccountParentID(ChartAccountID, "order by accountname")
Me.Caption = accntcode
'txtcode.Text = GetmaxID
Label5.Caption = LstAccountcode.ListItems.Count & " Records Found"
Call lOADcopyStat
End Sub
Public Function lOADcopyStat()
If IsCopy = False Then
'    lvButtons_H5.Caption = "Copy"
Else
    lvButtons_H5.Caption = "Paste"
End If
End Function

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



Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader
Case "Accountcode"
Call Get_Account_by_ChartAccountParentID(ChartAccountID, "order by accountcode")
Case "Accountname"
Call Get_Account_by_ChartAccountParentID(ChartAccountID, "order by accountname")
End Select
End Sub

Private Sub LstAccountcode_DblClick()
If LstAccountcode.ListItems.Count > 0 Then
Dim Newform As New f_COA_GL_sub
    With Newform
    .Address = Address & "~" & Trim(LstAccountcode.SelectedItem.SubItems(2))
    .col = 2
    .ChartAccountID = Trim(LstAccountcode.SelectedItem.Text)
    .Subcode1 = Trim(LstAccountcode.SelectedItem.SubItems(1))
    .Subdesc1 = Trim(LstAccountcode.SelectedItem.SubItems(2))
    .accntcode = accntcode & "-" & Trim(LstAccountcode.SelectedItem.SubItems(1))
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
                If ExecFunction("SELECT [fmis].[dbo].[MPfunc_CheckIfAccntIsUse] ('" & Trim(Me.Caption) & "-" & Trim(LstAccountcode.ListItems(x).Text) & "' )") = 0 Then
                
                lblstat.Caption = "Deleting Account " & LstAccountcode.ListItems(x).Text
                    opndbaseFMIS.Execute "Delete from [tblReff_CodeClassification]  where " & Getcondition & " and " & Subcode & " = '" & Trim(LstAccountcode.ListItems(x).Text) & "' and actioncode = 0 "
                Else
                    If MsgBox("Cannot Delete " & Trim(LstAccountcode.ListItems(x).Text) & "-" & Trim(LstAccountcode.ListItems(x).ListSubItems(1).Text) & ": Access is denied." & vbNewLine & vbNewLine & "Make sure that Accounts is not currently in Use." & vbNewLine & "Yes = Ignore" & vbNewLine & "No = Cancel Deleting", vbCritical + vbYesNo, "Error Deleting Accounts") = vbYes Then
                        'no action
                    Else
                        Exit For
                    End If
                End If
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
Private Function IfUsecode(ByVal accountcode As String)
Dim rec As New ADODB.Recordset

End Function
Private Sub lvButtons_H2_Click()
With frm_Import
    .Subcode1 = Subcode1
    .Subdesc1 = Subdesc1
    
    .Subcode2 = Subcode2
    .Subdesc2 = Subdesc2
    
    .Subcode3 = Subcode3
    .Subdesc3 = Subdesc3
    
    .Subcode4 = Subcode4
    .Subdesc4 = Subdesc4
    
    .Subcode5 = Subcode5
    .Subdesc5 = Subdesc5
    
    .Subcode6 = Subcode6
    .Subdesc6 = Subdesc6
    
    .Subcode7 = Subcode7
    .Subdesc7 = Subdesc7
    .col = col
    .accntcode = accntcode
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
    txtcode.Text = GetmaxID
    Else
    MsgBox "Please Specify the code, to Proceed...!", vbInformation, "System Message"
    txtcode.SetFocus
    Exit Sub
    End If
End If
If MsgBox("Are you sure do want to save?", vbInformation + vbYesNo, "System Message") = vbYes Then
    opndbaseFMIS.Execute "Insert into tblReff_CodeClassification (Subcode1, Subdesc1, Subcode2, Subdesc2,lvl) values ('" & Subcode1 & "','" & Replace(Subdesc1, "'", "''") & "','" & txtcode & "','" & UCase(Replace(txtadd.Text, "'", "''")) & "'," & col & ")"
    Call GetAccountNameForSetUp(LstAccountcode, "tblReff_CodeClassification", Field1, Field2, Condition, Field1)
    txtadd.Text = ""
    txtcode.Text = ""
End If
End Sub
Public Function IfExistname(ByVal name As String) As Boolean
Dim x As Integer
IfExistname = False
    For x = 1 To LstAccountcode.ListItems.Count
        If UCase(name) = Trim(LstAccountcode.ListItems(x).SubItems(1)) Then
            IfExistname = True
            Exit For
        End If
    Next x
End Function
Public Function IfExistcode(ByVal Code As String) As Boolean
Dim x As Integer
IfExistcode = False
    For x = 1 To LstAccountcode.ListItems.Count
        If Trim(Code) = Trim(LstAccountcode.ListItems(x).Text) Then
            IfExistcode = True
            Exit For
        End If
    Next x
End Function
Public Function GetmaxID()
Dim rec As New ADODB.Recordset
Dim sql As String

Select Case col:
Case 1
Case 2
    sql = "Select   max(subcode2) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 3
    sql = "Select max(subcode3) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 4
sql = "Select max(subcode4) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 5
    sql = "Select max(subcode5) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 6
    sql = "Select max(subcode6) as maxid from tblReff_CodeClassification where " & Getcondition & ""
Case 7
    sql = "Select max(subcode7) as maxid from tblReff_CodeClassification where " & Getcondition & ""
End Select

rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        GetmaxID = IIf(IsNull(rec!maxid), 0, rec!maxid) + 1
    Else
        GetmaxID = 1
    End If
rec.Close
Set rec = Nothing
End Function
Private Function Getcondition() As String
        Select Case (col):
        Case 2
            Getcondition = "Subcode1 = " & Subcode1
        Case 3
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "subcode2 = " & Subcode2
        Case 4
           Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "subcode3 = " & Subcode3
        Case 5
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "subcode4 = " & Subcode4
        Case 6
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "subcode5 = " & Subcode5
        Case 7
            Getcondition = "Subcode1 = " & Subcode1 & " and " & "Subcode2 = " & Subcode2 & " and " & "Subcode3 = " & Subcode3 & " and " & "Subcode4 = " & Subcode4 & " and " & "Subcode5 = " & Subcode5 & " and " & "subcode6 = " & Subcode6
        End Select
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
