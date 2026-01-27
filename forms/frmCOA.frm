VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmforCOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Conversion"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "frmCOA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCOA.frx":3AFA
   ScaleHeight     =   8520
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   9090
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      Begin VB.TextBox txtnme 
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
         Left            =   1440
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   120
         Width           =   6975
      End
      Begin VB.TextBox Text1 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtdetails 
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
         Height          =   615
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1035
         Width           =   6975
      End
      Begin VB.TextBox txtfind 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   8040
         Width           =   3375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   11033
         _Version        =   393216
         BackColor       =   16777215
         BackColorSel    =   8454143
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   495
         Left            =   8520
         TabIndex        =   4
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmCOA.frx":75F4
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnsave 
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         ToolTipText     =   "Save the Entry as New"
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   5
         Image           =   "frmCOA.frx":774E
         cBack           =   16777215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Click me to Add Others in the list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Name:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Accountcode:"
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
         Left            =   0
         TabIndex        =   8
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Account Details:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Press ENTER "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Search Name:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   8085
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmforCOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nme, accntcode As String
Public Trnno As Long
Public inRec As Boolean
Public frm As Form
Private Sub lvButtons_H3_Click()
On Error GoTo bad
Dim sql As String
Dim rec As New ADODB.Recordset
sql = ExecFunction("select [fmis].[dbo].GetqueryfromRtable (" & Combo1.ItemData(Combo1.ListIndex) & ")")
Set rec = opndbaseFMIS.Execute(sql)
If rec.RecordCount > 0 Then

Set MSFlexGrid1.DataSource = rec
End If
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub lvButtons_H5_Click()
frm_relatedtableForCOA.Show 1
End Sub

Private Sub Label6_Click()
Dim Subcode As Long
Dim x As Long
Dim lvl As Integer
If Trim(txtdetails.Text) = "" Then
MsgBox "Invalid Accountcode", vbInformation, "System Message"
Exit Sub
Else
    For x = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(x, 1) = "0" Then
            MsgBox "Already Exists in the database", vbInformation, "System Message"
            Exit For
            Exit Sub
        End If
    Next x
    
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyDesc] (" & val(GetLvlbyCode(Text1.Text)) + 1 & ",'0','Others')") = 1 Then
        MsgBox "Acocuntname is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Are you sure do you want to save the 0-Others Account?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        Subcode = 0
        lvl = GetLvlbyCode(Text1.Text)
        If lvl = 0 Then
        lvl = 2
        Else
        lvl = val(lvl) + 1
        End If
        opndbaseFMIS.Execute "Exec [Proc_CheckIfExistSub] @lvl = " & lvl & ",@childcode = 'Empty',@accountcode = '" & Text1.Text & "'," & _
        " @subcode =" & Subcode & ",@subdesc = 'Others'"
        '" & ExecFunction("SELECT [fmis].[dbo].[GetCOAIDbyDesc]  (" & GetLvlbyCode(Text1.Text) & ",'" & Trim(Text1.Text) & "')") & "
    End If
    
End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = &HFF&
End Sub

Private Sub MSHFlexGrid1_DblClick()
If MSHFlexGrid1.Rows > 1 Then
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
        If Right(Trim(Text1.Text), 1) = "-" Then
        Text1.Text = Text1.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
        Else
            If Len(Trim(Text1.Text)) < 3 Then
            Text1.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            Else
            Text1.Text = Trim(Text1.Text) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            End If
        End If
    End If
End If
txtfind.Text = ""
Text1.SetFocus
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = &H80000008
End Sub

Public Function LoadAccountsbySub(ByVal accountcode As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
Dim Code As Long
Dim childcode As String
Dim z
    xx = Split(accountcode, "-")
    str() = Split(accountcode, "-", -1, vbTextCompare)
    lvl = UBound(xx) + 1
    If lvl = 1 Then
        lvl = 0
    End If
    
    Select Case (lvl)
        Case 0
        If accountcode = "" Then
        Else
        childcode = str(0)
        End If
        Case 2
        childcode = str(0)
        Case 3
        childcode = str(0) & "-" & str(1)
        Case 4
        childcode = str(0) & "-" & str(1) & "-" & str(2)
        Case 5
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3)
        Case 6
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3) & "-" & str(4)
        Case 7
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3) & "-" & str(4) & "-" & str(5)
    End Select
    
    If Right(Trim(accountcode), 1) <> "-" Then
        accountcode = accountcode & "-"
        If lvl <> 0 Then
            lvl = lvl + 1
        Else
            lvl = lvl + 2
        End If
    End If
    
'    ListView2.ListItems.Clear
    ARec.Open ("Exec Proc_GetSubCode @find = '" & Trim(txtfind.Text) & "' , @lvl = " & lvl & ",@childcode = '" & accountcode & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Cols = 3
        MSHFlexGrid1.Rows = 2
        If ARec.RecordCount > 0 Then
            Set MSHFlexGrid1.DataSource = ARec
        End If
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 6000
    ARec.Close
    Set ARec = Nothing
End Function

Private Function LoadAccountsByName(ByVal accntcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    ARec.Open "exec Proc_getNamebychildCode @childaccountcode = '" & accntcode & "', @Condition = '" & Condition & "'", opndbaseFMIS, adOpenStatic
        If ARec.RecordCount > 0 Then
            LoadAccountsByName = ARec!Accountfullname
        inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function
Public Function GetAccountNamebyorder(ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother where accountcode like '" & Text1.Text & "%' and accountname like '" & Trim(txtfind.Text) & "%' order by Accountname", opndbaseFMIS, adOpenStatic, adLockOptimistic
    'lst.ListItems.Clear
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
    If rec.RecordCount > 0 Then
    
'        For z = 1 To rec.RecordCount
'                    Set x = lst.ListItems.Add(, , rec.Fields!Accountcode)
'                    x.SubItems(1) = Trim(rec.Fields!Accountname)
'            rec.MoveNext
'        Next z
    
    Set MSHFlexGrid1.DataSource = rec
        MSHFlexGrid1.Cols = 4
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Formula"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 8000
        MSHFlexGrid1.ColWidth(3) = 0
        
        
    End If
'rec.Close
Set rec = Nothing
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
         If Len(Trim(Text1.Text)) >= 3 Then
        LoadAccountsbySub (Text1.Text)
        txtdetails.Text = LoadAccountsByName(Text1.Text, "Fullname")
        Else
        txtdetails.Text = ""
        Call GetAccountNamebyorder("Accountcode")
        End If
        txtfind.Text = ""
    End If

End Sub

Private Sub btnsave_Click()
Dim Subcode As Long
Dim lvl As Integer
If Trim(txtdetails.Text) = "" Then
MsgBox "Invalid Accountcode", vbInformation, "System Message"
Exit Sub
Else
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyDesc] (" & val(GetLvlbyCode(Text1.Text)) + 1 & ",'" & Trim(Text1.Text) & "','" & txtnme.Text & "')") = 1 Then
        MsgBox "Acocuntname is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("The new Account is Save to " & Trim(txtdetails.Text) & "~" & Trim(txtnme.Text) & vbNewLine & "Are you sure do want to save the Account?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        Subcode = ExecFunction("SELECT [fmis].[dbo].[GetCOAIDbyDesc]  (" & IIf((GetLvlbyCode(Text1.Text) = 0), 2, val(GetLvlbyCode(Text1.Text)) + 1) & ",'" & Trim(Text1.Text) & "')")
        lvl = GetLvlbyCode(Text1.Text)
        If lvl = 0 Then
        lvl = 2
        Else
        lvl = val(lvl) + 1
        End If
        opndbaseFMIS.Execute "Exec [Proc_CheckIfExistSub] @lvl = " & lvl & ",@childcode = 'Empty',@accountcode = '" & Text1.Text & "'," & _
        " @subcode =" & Subcode & ",@subdesc = '" & txtnme.Text & "'"
        '" & ExecFunction("SELECT [fmis].[dbo].[GetCOAIDbyDesc]  (" & GetLvlbyCode(Text1.Text) & ",'" & Trim(Text1.Text) & "')") & "
    End If
    
End If
End Sub

Private Sub Form_Load()
txtnme.Text = nme
Text1.Text = accntcode
If Len(Text1.Text) = 0 Then
Call GetAccountNamebyorder("Accountcode")
Else
Call LoadAccountsbySub(accntcode)
End If
End Sub

Private Sub lvButtons_H6_Click()
'MsgBox "SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Text1.Text & "'," & col & ")"
'If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Text1.Text & "'," & GetLvlbyCode(Text1.Text) & ")") > 1 Then
'    MsgBox "Sorry, the Accountcode is not allow to assign if have Sub Accountcode", vbInformation, "System Message"
'    frm.IsOK = False
'    Exit Sub
'End If
msg = MsgBox("Save Change?", vbYesNoCancel, "System Confirmation")
If msg = vbYes Then
        If frm.name = "frm_COAQueryGenerator" Then
        frm.txtaccountcode.Text = Trim(Text1.Text)
        frm.txtdescription.Text = txtdetails.Text
        ElseIf frm.name = "frm_COAQueryGenerator" Then
        frm.txtaccountcode.Text = Trim(Text1.Text)
        frm.txtdescription.Text = txtdetails.Text
        ElseIf frm.name = "frm_COAQueryGenerator" Then
            frm.LstAccountcode.SelectedItem.SubItems(2) = Trim(Text1.Text)
        ElseIf frm.name = "frm_RRRSourceName" Then
            frm.LstAccountcode.SelectedItem.SubItems(2) = Trim(Text1.Text)
        Else
        frm.MSFlexGrid1.TextMatrix(frm.MSFlexGrid1.Row, 3) = Trim(Text1.Text)
        End If
        
        frm.IsOK = True
        Unload Me
ElseIf msg = vbNo Then
    frm.IsOK = False
    Unload Me
End If
End Sub
'Private Sub text1_Change()
'    If Len(Trim(Text1.Text)) >= 3 Then
'        LoadAccountsbySub (Text1.Text)
'        txtdetails.Text = LoadAccountsByName(Text1.Text, "Fullname")
'    Else
'    Call GetAccountNamebyorder("Accountcode")
'    End If
'    txtfind.Text = ""
'End Sub
Private Sub txtfind_Change()
 If Len(Trim(Text1.Text)) >= 3 Then
        LoadAccountsbySub (Text1.Text)
        txtdetails.Text = LoadAccountsByName(Text1.Text, "Fullname")
    Else
    Call GetAccountNamebyorder("Accountcode")
    End If
End Sub
