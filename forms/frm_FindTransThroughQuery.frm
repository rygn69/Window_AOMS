VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_FindTransThroughQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Transaction Through Query"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   Icon            =   "frm_FindTransThroughQuery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_print 
      Caption         =   "Print Preview"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox cmbdescription 
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
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   11775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11456
      _Version        =   393216
      ScrollTrack     =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Find"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   6480
      TabIndex        =   7
      Top             =   1395
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Select Criteria"
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
      TabIndex        =   5
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Criteria:"
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
      TabIndex        =   4
      Top             =   1400
      Width           =   735
   End
End
Attribute VB_Name = "frm_FindTransThroughQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_print_Click()
Dim frm As New frm_RPTQueryViewer
Report9 = "PTOAccnt111"
    With frm
    .accnt = "Execute [dbo].[MPproc_transThruQery] '" & cmbdescription.ItemData(cmbdescription.ListIndex) & "','" & txtsearch.Text & "'"
    .Show
    End With
End Sub

Private Sub Command1_Click()
Dim rec As New ADODB.Recordset
Dim x As Long
On Error GoTo bad
Set rec = opndbaseFMIS.Execute("execute fmis.dbo.MPproc_transThruQery @trnno = " & cmbdescription.ItemData(cmbdescription.ListIndex) & ",@field = '" & txtsearch.Text & "'")
Set MSHFlexGrid1.DataSource = rec
    If cmbdescription.ItemData(cmbdescription.ListIndex) = 1 Then
        MSHFlexGrid1.TextMatrix(0, 1) = "FMIS number"
        MSHFlexGrid1.TextMatrix(0, 2) = "Dvno"
        MSHFlexGrid1.TextMatrix(0, 3) = "Claimant"
        MSHFlexGrid1.TextMatrix(0, 4) = "Particular"
        MSHFlexGrid1.TextMatrix(0, 5) = "Disbursing Officer"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 1500
        MSHFlexGrid1.ColWidth(2) = 1500
        MSHFlexGrid1.ColWidth(3) = 3500
        MSHFlexGrid1.ColWidth(4) = 4000
    ElseIf cmbdescription.ItemData(cmbdescription.ListIndex) = 12 Then
        MSHFlexGrid1.TextMatrix(0, 1) = "JEVno"
        MSHFlexGrid1.TextMatrix(0, 2) = "JEVdate"
        MSHFlexGrid1.TextMatrix(0, 3) = "Fundtype"
        MSHFlexGrid1.TextMatrix(0, 4) = "Checkno"
        MSHFlexGrid1.TextMatrix(0, 5) = "PTO Amount"
        MSHFlexGrid1.TextMatrix(0, 6) = "JEVCreditEntry"
        MSHFlexGrid1.TextMatrix(0, 7) = "PTO Accountname"
        MSHFlexGrid1.TextMatrix(0, 8) = "Accnt. AccountName"
        MSHFlexGrid1.TextMatrix(0, 9) = "PTO Classification"
        MSHFlexGrid1.TextMatrix(0, 10) = "Accnt. Accountcode"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 1500
        MSHFlexGrid1.ColWidth(2) = 1000
        MSHFlexGrid1.ColWidth(3) = 2000
        MSHFlexGrid1.ColWidth(4) = 1300
        MSHFlexGrid1.ColWidth(5) = 1500
'        MSHFlexGrid1.ColAlignment(5) = 3
'        MSHFlexGrid1.ColAlignment(6) = 3
        MSHFlexGrid1.ColWidth(6) = 1500
        MSHFlexGrid1.ColWidth(7) = 3000
        MSHFlexGrid1.ColWidth(8) = 8000
        MSHFlexGrid1.ColWidth(9) = 1500
        MSHFlexGrid1.ColWidth(10) = 1500
        
        For x = 1 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.TextMatrix(x, 5) = Format(MSHFlexGrid1.TextMatrix(x, 5), "#,##0.00")
            MSHFlexGrid1.TextMatrix(x, 6) = Format(MSHFlexGrid1.TextMatrix(x, 6), "#,##0.00")
            DoEvents
        Next x
    End If
    
    
    Label3.Caption = MSHFlexGrid1.Rows - 1 & " Record(s) Found"
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub Form_Load()
Dim rec As New ADODB.Recordset
Dim x As Long
rec.Open "Select * from tblAMIS_SearchThruQuery", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
For x = 1 To rec.RecordCount
    With cmbdescription
        .AddItem rec!description
        .ItemData(.NewIndex) = rec!Trnno
        rec.MoveNext
    End With
Next x
End If
rec.Close
End Sub

Private Sub Form_Resize()
On Error Resume Next
MSHFlexGrid1.Height = ScaleHeight - (ScaleHeight / 4.5)
MSHFlexGrid1.Width = ScaleWidth - 0.5
cmbdescription.Width = ScaleWidth - 0.5
End Sub

