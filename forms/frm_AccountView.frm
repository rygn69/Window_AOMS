VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AccountView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart of Accounts Viewer"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "frm_AccountView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_AccountView.frx":3AFA
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
         Left            =   12720
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   480
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
         Top             =   240
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
         Top             =   675
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
         Height          =   6375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   11245
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
         Top             =   720
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
         Image           =   "frm_AccountView.frx":75F4
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnsave 
         Height          =   375
         Left            =   19800
         TabIndex        =   12
         ToolTipText     =   "Save the Entry as New"
         Top             =   480
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
         Image           =   "frm_AccountView.frx":774E
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
         Top             =   120
         Visible         =   0   'False
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
         Left            =   11400
         TabIndex        =   11
         Top             =   525
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
         Top             =   285
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
         Top             =   720
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
Attribute VB_Name = "frm_AccountView"
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
MsgBox err.Description
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
Call Text1_KeyPress(13)
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
frm.cmb_Accountcode.Text = Text1.Text
Unload Me
End Sub

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
 
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(Text1.Text)) >= 3 Then
        LoadAccountsbySub (Text1.Text)
        txtdetails.Text = LoadAccountsByName(Text1.Text, "Fullname")
    Else
    Call GetAccountNamebyorder("Accountcode")
    End If
End If
End Sub
