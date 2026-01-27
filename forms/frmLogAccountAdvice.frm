VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLogAccountAdvice 
   Caption         =   "Log Printed Account Advice"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmLogAccountAdvice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14420
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "Account's Advice Number:"
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
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogAccountAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
loadLog
End Sub

Private Sub Form_Load()
loadLog
End Sub
Public Function loadLog()
Dim rec As New ADODB.Recordset
rec.Open "Select sum(logprinted) as [Log Printed],adviceno as [Advice No.],date_time as [Date Printed],rtrim(ltrim(lastname)) + ', ' + rtrim(ltrim(firstname)) + ' ' + rtrim(ltrim(left(lastname,1))) + '.' as [User Printed] from tblAMIS_logPrintedAccntAdvice as a inner join pmis.dbo.Employee on SwipEmployeeID = userid  where adviceno like '" & Text1.Text & "%' group by adviceno,date_time,userid,rtrim(ltrim(lastname)) + ', ' + rtrim(ltrim(firstname)) + ' ' + rtrim(ltrim(left(lastname,1))) + '.'", opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
If rec.RecordCount <> 0 Then
    Set MSHFlexGrid1.DataSource = rec
    Call SetGrid
End If
rec.Close
End Function

Private Sub SetGrid()
Dim cc As Integer

'MSFlexGrid1.Clear
'MSFlexGrid1.Cols = 7
'MSFlexGrid1.Rows = 2

'    MSFlexGrid1.TextMatrix(0, 0) = ""
'    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
'    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
'    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
'    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
'    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
'    MSFlexGrid1.TextMatrix(0, 6) = "Actioncode"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 3000
'MSFlexGrid1.ColWidth(4) = 2000
'MSFlexGrid1.ColWidth(5) = 2000
'MSFlexGrid1.ColWidth(6) = 1500

For cc = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Row = 0
    MSHFlexGrid1.col = cc
    MSHFlexGrid1.CellAlignment = 4
Next cc
End Sub
