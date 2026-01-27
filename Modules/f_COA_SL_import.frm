VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form f_COA_SL_import 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "f_COA_SL_import.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbsheets 
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
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   5040
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Optin 
      Caption         =   "Include Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "opndbasePMIS"
      Top             =   8400
      Width           =   1815
   End
   Begin VB.OptionButton Optex 
      Caption         =   "System Generate Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      ToolTipText     =   "opndbaseFMIS"
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connect To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox cmbconn 
         Enabled         =   0   'False
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Other Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FMIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "opndbaseFMIS"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PMIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "opndbasePMIS"
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   2040
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Press F5 to Execute Query"
      Top             =   825
      Width           =   7215
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   8280
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
      Image           =   "f_COA_SL_import.frx":076A
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
      Left            =   8880
      TabIndex        =   3
      Top             =   9000
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "f_COA_SL_import.frx":4274
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   735
      Left            =   9120
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "&Save Query"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_import.frx":7D7E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   735
      Left            =   9120
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "&New Query"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_import.frx":80D0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   735
      Left            =   9120
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "&Delete Query"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_import.frx":9122
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   735
      Left            =   9120
      TabIndex        =   20
      ToolTipText     =   "Browse Excel File"
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "&Browse"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_import.frx":CC2C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H7 
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      ToolTipText     =   "Browse Excel File"
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Load"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "f_COA_SL_import.frx":DEAE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H8 
      Height          =   1095
      Left            =   9120
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Caption         =   "&Save with Sub"
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
      Image           =   "f_COA_SL_import.frx":119B8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "Sheet Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblname 
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
      TabIndex        =   19
      Top             =   9000
      Width           =   4455
   End
   Begin VB.Label lblcount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   8400
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Query Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10080
      Y1              =   8880
      Y2              =   8880
   End
End
Attribute VB_Name = "f_COA_SL_import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Accountname, code, childcode As String
Public ChartAccountChildID, levelno As Long
Public IfEdited As Boolean
Public nme, path As String
Private Function loadQuery()
Dim conn As String
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
If Option1.Value = True Then
   rec.Open txtaddress.Text, opndbaseFMIS, adOpenStatic, adLockOptimistic
ElseIf Option2.Value = True Then
rec.Open txtaddress.Text, opndbaseFMIS, adOpenStatic, adLockOptimistic
ElseIf Option4.Value = True Then
GetNewConnection (cmbconn.ItemData(cmbconn.ListIndex))
rec.Open txtaddress.Text, NewConnection, adOpenStatic, adLockOptimistic
End If
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Cols = 3
        MSHFlexGrid1.Rows = 2
    If rec.RecordCount > 0 Then
    lblcount.Caption = rec.RecordCount & " Records Found"
    Set MSHFlexGrid1.DataSource = rec
    
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 6500
    End If
Set rec = Nothing
Exit Function
rec.Close
bad:
MsgBox "Error:" & err.Number & vbNewLine & err.description, vbCritical, "System Message"
End Function

Private Function loadsqlQuery(ByVal Trnno As Long) As String
Dim qrec As New ADODB.Recordset
On Error Resume Next
 loadsqlQuery = ""
qrec.Open "select sqlquery,connection from tblAMIS_SqlQuery where trnno = " & List1.ItemData(List1.ListIndex) & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If qrec.RecordCount > 0 Then
        loadsqlQuery = qrec!sqlquery
        If Trim(qrec!Connection) = "FMIS" Then
        Option2.Value = True
        cmbconn.Enabled = False
        cmbconn.ListIndex = 0
        ElseIf Trim(qrec!Connection) = "PMIS" Then
        Option1.Value = True
        cmbconn.ListIndex = 0
        cmbconn.Enabled = False
        Else
        Option4.Value = True
        cmbconn.Text = Trim(qrec!Connection)
        End If
    Else
        Option4.Value = True
        cmbconn.ListIndex = 0
    End If
qrec.Close
End Function

Private Sub list1_Change()
txtaddress.Text = loadsqlQuery(List1.ItemData(List1.ListIndex))
End Sub

Private Sub cmbconn_Click()
GetNewConnection (cmbconn.ItemData(cmbconn.ListIndex))
End Sub

Private Sub List1_Click()
txtaddress.Text = Replace(loadsqlQuery(List1.ItemData(List1.ListIndex)), "@brgyID", code)
loadQuery

End Sub

Private Sub Form_Load()
Call LoadQueryDesc
Call loadConectionName
Me.Caption = childcode
End Sub
Public Function LoadQueryDesc()
Dim rec As New ADODB.Recordset
Dim x As Integer
List1.Clear
rec.Open "Select trnno,sqldesc from tblAMIS_SqlQuery order by sqldesc ", opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
List1.AddItem ""
List1.Clear
    If rec.RecordCount > 0 Then
        List1.AddItem "<Add New>"
        For x = 1 To rec.RecordCount
            List1.AddItem Trim(rec!sqldesc)
            List1.ItemData(List1.NewIndex) = rec!Trnno
            rec.MoveNext
        Next x
    End If
rec.Close
Set rec = Nothing
End Function
Private Sub lvButtons_H1_Click()
Dim Connection, desc As String
Dim rec As New ADODB.Recordset
On Error GoTo bad
Connection = ""
If Option1.Value = True Then: Connection = Option1.Caption
If Option2.Value = True Then: Connection = Option2.Caption
If Option4.Value = True Then: Connection = Trim(cmbconn.Text)
rec.Open "Select * from tblAMIS_SqlQuery where trnno = " & IIf((List1.ListIndex = -1), 0, List1.ItemData(List1.ListIndex)) & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    If MsgBox("Are you sure do you want to Update the Query?", vbCritical + vbYesNo, "System Message") = vbYes Then
    opndbaseFMIS.Execute "Update tblAMIS_SqlQuery set sqlquery = '" & Replace(Trim(txtaddress.Text), "'", "''") & "',userid = '" & Trim(ActiveUserID) & "',datetimeentered =  '" & Now & "' where trnno = " & IIf((List1.ListIndex = -1), 0, List1.ItemData(List1.ListIndex)) & ""
    End If
Else
    desc = InputBox("Description:", "SQL Query Field")
    If MsgBox("Are you sure do you save the Query?", vbCritical + vbYesNo, "System Message") = vbYes Then
    opndbaseFMIS.Execute "insert into tblAMIS_SqlQuery (sqlquery,sqldesc,datetimeentered,userid,connection) values ('" & Replace(txtaddress.Text, "'", "''") & "','" & Replace(desc, "'", "''") & "','" & Now & "','" & ActiveUserID & "','" & Connection & "')"
    MsgBox "Successfully Save...", vbInformation, "System Message"
    Call LoadQueryDesc
    List1.Text = Trim(desc)
    End If
End If
rec.Close
Set rec = Nothing

Exit Sub
bad:
MsgBox err.description
End Sub
Private Sub lvButtons_H2_Click()
txtaddress.Text = ""
List1.ListIndex = 0
Option4.Value = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.TextMatrix(0, 1) = "Code"
MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 700
MSHFlexGrid1.ColWidth(2) = 6500
End Sub

Private Sub lvButtons_H3_Click()
On Error GoTo bad
Dim x As Integer
Dim Acode As String
Dim Aname As String
Dim rec_sub As New ADODB.Recordset
If MsgBox("Are you sure do want to import data?", vbInformation + vbYesNo, "System Message") = vbYes Then
progStat.Max = MSHFlexGrid1.Rows - 1
progStat.Visible = True

    For x = 1 To MSHFlexGrid1.Rows - 1
Resumes:
        If Trim(MSHFlexGrid1.TextMatrix(x, 1)) <> "" Or Trim(MSHFlexGrid1.TextMatrix(x, 2)) <> "" Then
        Aname = Trim(MSHFlexGrid1.TextMatrix(x, 2))
            If Optin.Value = True Then
                Acode = Trim(MSHFlexGrid1.TextMatrix(x, 1))
            ElseIf Optex.Value = True Then
                Acode = Get_MaxValue_COA(ChartAccountChildID)
            Else
                MsgBox "Please Select either Include code or System Generate Code", vbInformation, "System Message"
                Exit For
                Exit Sub
            End If
            
            If IsNumeric(Acode) = False Then
                MsgBox "Code is none Numeric,Please Check the details", vbInformation, "System Message"
                Exit Sub
            End If
            
            If ExecFunction("Select Accounting.fn_check_IfExistCode_COA  (" & ChartAccountChildID & ",'" & Acode & "')") = 1 Then
                'If MsgBox("Code " & Code & " is already exist in the database. Do you want to ignore and continue?", vbCritical + vbYesNo, "System Confirmation") = vbNo Then
                   ' Exit For
                'End If
            Else
                If ExecFunction("Select Accounting.fn_check_IfExistName_COA   (" & ChartAccountChildID & ",'" & Replace(Aname, "'", "''") & "')") = 1 Then
                    'If MsgBox("Name '" & Trim(MSHFlexGrid1.TextMatrix(x, 2)) & "' is already exist in the database. Do you want to ignore and continue?", vbCritical + vbYesNo, "System Confirmation") = vbNo Then
                       ' Exit For
                    'End If
                Else
                    If Aname <> "" And Acode <> "" Then
                        opndbaseFMIS.Execute "exec Accounting.usp_Save_ChartOfAccountsChild @code = '" & Acode & "',@AccountChildParentID = '" & ChartAccountChildID & "' ,@AccountChildName= '" & Replace(Aname, "'", "''") & "',@parentchildcode = '" & childcode & "' ,@parentLevelno = '" & levelno & "',@ModifiedByID = '" & ActiveUserID & "'"
                    End If
                End If
            End If
            
            
            
        End If
        progStat.Value = x
        DoEvents
        lblname.Caption = UCase(Trim(MSHFlexGrid1.TextMatrix(x, 2)))
            
    Next x
'MsgBox "Save Successfully..!", vbInformation, "System Information"
lblname.Caption = ""
progStat.Visible = False
Unload Me
End If
Exit Sub
bad:
If err.Number = -2147467259 Then
   MDIFrm_MAIN.tmeConnChck.Enabled = False
    frmConnCheck.Show 1
    MDIFrm_MAIN.tmeConnChck.Enabled = True
    GoTo Resumes
Else
MsgBox "Error: " & err.Number & " " & err.description & vbNewLine & "Please Contact System Administrator"
End If
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
Private Sub lvButtons_H4_Click()
Unload Me
End Sub

Private Sub lvButtons_H5_Click()
If MsgBox("Are you sure do you want to delete the Query?...", vbCritical + vbYesNo, "System Information") = vbYes Then
    opndbaseFMIS.Execute "Delete from tblAMIS_SqlQuery where trnno = '" & List1.ItemData(List1.ListIndex) & "'"
    MsgBox "Successfully Deleted", vbInformation, "System Messgae"
    List1.ListIndex = 0
    Call Form_Load
End If
End Sub

Private Sub lvButtons_H6_Click()
On Error GoTo bad
Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim var As Variant
    
    With Com
        .DialogTitle = "Load Excel File"
        .Filter = "EXCEL 2007 (*.xlsx) | *.xlsx" & "|" & "EXCEL 2003 (*.xls) | *.xls"
        .ShowOpen
        nme = .FileName
    End With
    txtaddress.Text = nme
   Set xlApp = New Excel.Application

   Set wb = xlApp.Workbooks.Open(nme)
   
       cmbsheets.Clear
        For x = 1 To xlApp.Worksheets.Count
        cmbsheets.AddItem wb.Worksheets.Item(x).name '  Item(x).name
        DoEvents
        Next x
        
   wb.Close

   xlApp.Quit

   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub lvButtons_H7_Click()
On Error GoTo bad
Dim xlApp As Excel.Application
   Dim wb As Workbook
   Dim ws As Worksheet
   
   Dim var As Variant
    
    
    
   Set xlApp = New Excel.Application

   Set wb = xlApp.Workbooks.Open(nme)

   Set ws = wb.Worksheets(cmbsheets.Text) 'Specify your worksheet name
   'var = ws.Range("A1").Value

   'or
   
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Cols = 3
        MSHFlexGrid1.Rows = ws.UsedRange.Rows.Count + 2
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 6500
        progStat.Max = ws.UsedRange.Rows.Count
    progStat.Visible = True
    For x = 1 To ws.UsedRange.Rows.Count + 1
        MSHFlexGrid1.TextMatrix(x, 1) = ws.Cells(x, 1).Value
        MSHFlexGrid1.TextMatrix(x, 2) = ws.Cells(x, 2).Value
        progStat.Value = x
        DoEvents
    Next x
    
        progStat.Visible = False
   wb.Close

   xlApp.Quit

   Set ws = Nothing
   Set wb = Nothing
   Set xlApp = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub MSHFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    With MSHFlexGrid1
        .RemoveItem (.Row)
    End With
End If
End Sub

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

Private Sub Option4_Click()
If Option4.Value = True Then
cmbconn.Enabled = True
Call loadConectionName
End If
End Sub

Private Sub txtaddress_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
    loadQuery
End If
End Sub
Public Function loadConectionName()
Dim rec As New ADODB.Recordset
Dim x As Integer
rec.Open "Select *  from tblreff_ManageConnection", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmbconn.Clear
    cmbconn.AddItem " "
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            cmbconn.AddItem Trim(rec!name)
            cmbconn.ItemData(cmbconn.NewIndex) = rec!Trnno
            rec.MoveNext
        Next x
    End If
rec.Close
Set rec = Nothing
End Function
