VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AccountSpecialEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Conversion"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "frm_AccountSpecialEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_AccountSpecialEntry.frx":3AFA
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Auto Close"
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
         Left            =   7680
         TabIndex        =   15
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtcode 
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
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   7920
         Width           =   1575
      End
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
         Left            =   1800
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   7920
         Width           =   6615
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
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
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
         Top             =   555
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
         Left            =   1800
         TabIndex        =   1
         Top             =   7080
         Width           =   6615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   9975
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
         ToolTipText     =   "Close form"
         Top             =   600
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
         Image           =   "frm_AccountSpecialEntry.frx":75F4
         ImgSize         =   24
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnsave 
         Height          =   375
         Left            =   8520
         TabIndex        =   11
         ToolTipText     =   "Save the Entry as New"
         Top             =   7920
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
         ImgAlign        =   4
         Image           =   "frm_AccountSpecialEntry.frx":B0FE
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   8520
         TabIndex        =   16
         ToolTipText     =   "Save the Entry as New"
         Top             =   7080
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
         ImgAlign        =   4
         Image           =   "frm_AccountSpecialEntry.frx":EC08
         cBack           =   16777215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Code:"
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
         Left            =   -600
         TabIndex        =   14
         Top             =   7560
         Width           =   1215
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
         Left            =   1200
         TabIndex        =   12
         Top             =   7560
         Width           =   1215
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
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
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
         TabIndex        =   7
         Top             =   165
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
         TabIndex        =   6
         Top             =   600
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
         Left            =   480
         TabIndex        =   5
         Top             =   7125
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_AccountSpecialEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nme, accntcode, field, Searchname As String
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
    
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyDesc] (" & Val(GetLvlbyCode(Text1.Text)) + 1 & ",'0','Others')") = 1 Then
        MsgBox "Acocuntname is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Are you sure do you want to save the 0-Others Account?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        Subcode = 0
        lvl = GetLvlbyCode(Text1.Text)
        If lvl = 0 Then
        lvl = 2
        Else
        lvl = Val(lvl) + 1
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

Private Sub lvButtons_H1_Click()
medll.centerme frmCDClaimantRegistry
frmCDClaimantRegistry.Show 1
End Sub

Private Sub lvButtons_H6_Click()
Unload Me
End Sub

Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.Row <> 0 Then
txtcode.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
txtnme.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
End If
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

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub MSHFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Call MSHFlexGrid1_Click
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = &H80000008
End Sub
Public Function GetAccountNamebyorder(ByVal Accountname As String)
On Error Resume Next
Dim rec As New ADODB.Recordset
Dim SRec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim tmpquery As String
 
rec.Open "Select Query from tblAMIS_queryLinkToAccountname where Whatname = '" & Replace(Me.Caption, "'", "''") & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
    If rec.RecordCount > 0 Then
    tmpquery = Replace(rec!query, "@field", "'" & txtfind.Text & "'")
    'MsgBox Tmpquery
    Set SRec = opndbaseFMIS.Execute(tmpquery)
    'MsgBox Tmpquery
        If SRec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = SRec
        Else
            MSHFlexGrid1.Cols = 4
        End If
            MSHFlexGrid1.TextMatrix(0, 1) = "Code"
            MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
    
            MSHFlexGrid1.ColWidth(0) = 0
            MSHFlexGrid1.ColWidth(1) = 700
            MSHFlexGrid1.ColWidth(2) = 8000
    Else
    MsgBox "Opps..! No link found", vbInformation, "System Message"
    End If
'rec.Close
Set rec = Nothing
Exit Function
bad:
MsgBox "Note: " & err.Description, vbCritical, "System Message"
End Function
Private Sub btnsave_Click()
Dim Subcode As Long
Dim lvl As Integer
If Trim(txtcode.Text) = "" And txtnme.Text = "" Then
MsgBox "Invalid Accountcode", vbInformation, "System Message"
Exit Sub
Else
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyDesc] (" & IIf(((Val(GetLvlbyCode(Text1.Text)) + 1) = 1), 2, Val(GetLvlbyCode(Text1.Text)) + 1) & ",'" & Trim(Text1.Text) & "','" & txtnme.Text & "')") = 1 Then
        MsgBox "Acocuntname is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    'MsgBox Trim(Text1.Text) & "-" & Trim(txtcode.Text)
    If ExecFunction("SELECT [fmis].[dbo].[MPfunc_ChkIfAlreadyInCOAbyCODE] (" & IIf(((Val(GetLvlbyCode(Text1.Text)) + 1) = 1), 2, Val(GetLvlbyCode(Text1.Text)) + 1) & ",'" & Trim(Text1.Text) & "-" & Trim(txtcode.Text) & "')") = 1 Then
        MsgBox "Code is Already Exist in the database", vbInformation, "System Message"
        Exit Sub
    End If
    
   ' If MsgBox("The new Account is Save to " & Trim(txtdetails.Text) & "~" & Trim(txtnme.Text) & vbNewLine & "Are you sure do want to save the Account?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        Subcode = txtcode.Text
        lvl = GetLvlbyCode(Text1.Text)
        If lvl = 0 Then
        lvl = 2
        Else
        lvl = Val(lvl) + 1
        End If
        opndbaseFMIS.Execute "Exec [Proc_CheckIfExistSub] @lvl = " & lvl & ",@childcode = 'Empty',@accountcode = '" & Text1.Text & "'," & _
        " @subcode =" & Subcode & ",@subdesc = '" & txtnme.Text & "'"
        '" & ExecFunction("SELECT [fmis].[dbo].[GetCOAIDbyDesc]  (" & GetLvlbyCode(Text1.Text) & ",'" & Trim(Text1.Text) & "')") & "
    'End If
    Unload Me
End If
End Sub

Private Sub Form_Load()
txtnme.Text = nme
txtfind.Text = Searchname
Text1.Text = accntcode
Me.Caption = field
Call GetAccountNamebyorder(field)
End Sub
Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call GetAccountNamebyorder(field)
End If
End Sub
