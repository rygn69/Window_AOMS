VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmOtherClass 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   Icon            =   "frmOtherClass.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmOtherClass.frx":076A
   ScaleHeight     =   9645
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbfund 
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
      ItemData        =   "frmOtherClass.frx":AE19
      Left            =   3960
      List            =   "frmOtherClass.frx":AE1B
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmOtherClass.frx":AE1D
      Left            =   120
      List            =   "frmOtherClass.frx":AE1F
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   3735
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   9720
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "&Close"
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
      Image           =   "frmOtherClass.frx":AE21
      cBack           =   16777215
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6240
      Top             =   8880
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   5160
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   10080
      Width           =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   615
      Left            =   13200
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      Image           =   "frmOtherClass.frx":E92B
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7905
      ScaleWidth      =   10545
      TabIndex        =   1
      Top             =   1560
      Width           =   10575
      Begin lvButton.lvButtons_H lvlbrowse 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         Image           =   "frmOtherClass.frx":12435
         cBack           =   16777215
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   7935
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   13996
         _Version        =   393216
         ForeColorSel    =   65535
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
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   615
      Left            =   7440
      TabIndex        =   10
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "&Load"
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
      Image           =   "frmOtherClass.frx":1258F
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   615
      Left            =   8520
      TabIndex        =   12
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Query Settings"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "frmOtherClass.frx":16099
      cBack           =   16777215
   End
   Begin VB.Frame Frame1 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton Option3 
         Caption         =   "All"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "W/o Accountcode"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Accountcode"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundtype"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please identify the list below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   2430
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion Table for Chart of Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Table:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmOtherClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public reff, datetimeentered, UserID, Accountname, CName As String
Public Damount, Camount, Gamount As Currency
Public isEdit, inRec, ifCMB, IfEdit, IfNew, insert, delete, isOK As Boolean
Public Transtype, cmbtext As Integer

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    'Name
    MSFlexGrid1.TextMatrix(0, 0) = "ID"
    MSFlexGrid1.TextMatrix(0, 1) = "AccountCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 700
    MSFlexGrid1.ColWidth(2) = 8000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColAlignment(1) = 1
    
    
End Sub

Private Sub Combo1_Change()
Call Combo1_Click
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "PTO" Then
Label3.Visible = True
cmbfund.Visible = True
'lvButtons_H3.Left = 5760
Call LoadFund
ElseIf Combo1.Text = "Cashbook" Then
Call LoadFund
Label3.Visible = True
cmbfund.Visible = True
'lvButtons_H3.Left = 5760
Else
Label3.Visible = False
cmbfund.Visible = False
End If
End Sub

Private Sub Form_Load()
   ' LoadDetails (reff)
    Call Loadcombo("Select trnno as field1,queryname as field2 from tblAMIS_RelatedTableforCOA")
End Sub
Private Sub LoadFund()
Dim Frec As New ADODB.Recordset
Dim x As Integer

cmbfund.Clear

Frec.Open ("Select * from tblRefBMS_Funds Order By FundMedium"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    For x = 1 To Frec.RecordCount
        cmbfund.AddItem Frec!FundName
        cmbfund.ItemData(cmbfund.NewIndex) = Frec!fundcode
        Frec.MoveNext
    Next x
End If
Frec.Close
Set Frec = Nothing

End Sub
Public Function Loadformload()
Call Form_Load
Combo1.Text = Trim(cmbtext)
End Function
Public Function Loadcombo(ByVal sql As String)
Dim rec As New ADODB.Recordset
Dim x As Integer
Combo1.Clear
Combo1.AddItem ""
rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            Combo1.AddItem Trim(rec!Field2)
            Combo1.ItemData(Combo1.NewIndex) = CInt(rec!Field1)
            rec.MoveNext
            DoEvents
        Next x
    End If
rec.Close
End Function
'Private Function LoadDetails(ByVal reff As String)
'Dim DRec As New ADODB.Recordset
'Dim Sql As String
'
'Sql = "Select trnno ,ChildAccountcode, Debit,credit,actioncode,datetimeentered,userid From tblAMIS_AccoutingEntries Where [reffno]='" & reff & "' And (ActionCode=1)"
'
'DRec.Open Sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
'    Call SetGrid
'    If DRec.RecordCount > 0 Then
'        UserID = DRec!UserID
'        For x = 1 To DRec.RecordCount
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
'            MSFlexGrid1.TextMatrix(x, 1) = DRec!childaccountcode
'            MSFlexGrid1.TextMatrix(x, 2) = LoadAccountsByName(DRec!childaccountcode, "Summary")
'            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!Credit, "#,##0.00") = "0.00"), "", Format(DRec!Credit, "#,##0.00"))
'            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!Debit, "#,##0.00") = "0.00"), "", Format(DRec!Debit, "#,##0.00"))
'            'MSFlexGrid1.TextMatrix(x, 1) = DRec!ActionCode
'            DRec.MoveNext
'        Next x
'        Call GetSum
'    Else
'    MSFlexGrid1.TextMatrix(1, 2) = "TOTAL"
'    End If
'    DRec.Close
'    Set DRec = Nothing
'
'End Function

Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H3_Click()
On Error GoTo bad
Dim sql As String
Dim rec As New ADODB.Recordset
Dim str() As String
Dim splt As String
sql = ExecFunction("select [fmis].[dbo].GetqueryfromRtable (" & Combo1.ItemData(Combo1.ListIndex) & ")")

If Trim(Combo1.Text) = "PTO" Then
str() = Split(sql, "order", -1, vbTextCompare)
sql = str(0) & " where fundtype = '" & cmbfund.Text & "' order" & str(1)
'MsgBox sql
ElseIf Trim(Combo1.Text) = "Cashbook" Then
str() = Split(sql, "order", -1, vbTextCompare)
sql = str(0) & " where fundtype = '" & cmbfund.Text & "' order" & str(1)
End If

Set rec = opndbaseFMIS.Execute(sql)
If rec.RecordCount > 0 Then
Call SetGrid
Set MSFlexGrid1.DataSource = rec
End If
Exit Sub
bad:
MsgBox err.Description
End Sub

Private Sub lvButtons_H5_Click()
frm_relatedtableForCOA.Show
End Sub


Private Sub lvlbrowse_Click()
Dim rec As New ADODB.Recordset
isOK = False
With frmforCOA
.nme = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
.accntcode = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
.Trnno = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
.Show 1
End With
If isOK = True Then
      rec.Open "select tables,columns,conditions from tblAMIS_RelatedTableforCOA where trnno = " & Combo1.ItemData(Combo1.ListIndex) & "", opndbaseFMIS, adOpenStatic
      If rec.RecordCount > 0 Then
        Call UpdateExtractor(IIf(IsNull(rec!Tables), "", rec!Tables), IIf(IsNull(rec!columns), "", rec!columns), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), IIf(IsNull(rec!Conditions), "", rec!Conditions), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)))
      End If
      rec.Close
End If
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
        childcode = str(0)
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
End Function

Private Function LoadAccountsByName(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    ARec.Open "exec Proc_getNamebychildCode @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'", opndbaseFMIS, adOpenStatic
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
rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother where accountcode like '" & cmbEntry.Text & "%' and accountname like '" & Trim(txtfind.Text) & "%' order by Accountname", opndbaseFMIS, adOpenStatic, adLockOptimistic
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

Private Sub MSFlexGrid1_DblClick()
On Error GoTo bad

    Select Case MSFlexGrid1.col
    Case 3 'Debit/Credit
        
            Dim rec As New ADODB.Recordset
            isOK = False
            With frmforCOA
            .nme = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
            .accntcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
            .Trnno = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
            Set .frm = Me
            .Show 1
            End With
            If isOK = True Then
                  rec.Open "select tables,columns,conditions from tblAMIS_RelatedTableforCOA where trnno = " & Combo1.ItemData(Combo1.ListIndex) & "", opndbaseFMIS, adOpenStatic
                  If rec.RecordCount > 0 Then
                    Call UpdateExtractor(IIf(IsNull(rec!Tables), "", rec!Tables), IIf(IsNull(rec!columns), "", rec!columns), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), IIf(IsNull(rec!Conditions), "", rec!Conditions), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)))
                  End If
                  rec.Close
            End If
        
    End Select
Exit Sub
bad:
MsgBox err.Description
End Sub
Private Sub txt_entry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       ' If txt_entry.ListIndex <> -1 Then
            inRec = False
            Accountname = LoadAccountsByName(txt_entry.Text, "Summary")
            If Trim(txt_entry.Text) <> "" Then
                If inRec = False Then
                    If txt_entry.Text <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) Then
                    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                    End If
                End If
                ifCMB = True
                If Chckentry = False Then
                Exit Sub
                End If
                ifCMB = False
            End If
            
            
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = txt_entry.Text
            
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                    
            Else
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                End If
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
            
            
        If txt_entry.Text = "" Then
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
               
                    MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) <> "TOTAL" Then
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
                    
                End If
            End If
        Else
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                    End If
            Else
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
        End If
        txt_entry.Visible = False
        ListView2.Visible = False
        Picture2.Visible = False
        'Call GetSum
        MSFlexGrid1.SetFocus
        isEdit = True
    Else
       ' KeyAscii = AutoFind(txt_entry, KeyAscii, True)
        ListView2.Visible = True
        ListView2.Move MSFlexGrid1.CellLeft + txt_entry.Width
        Picture2.Visible = True
       Picture2.Move MSFlexGrid1.CellLeft + txt_entry.Width
    End If

End Sub
