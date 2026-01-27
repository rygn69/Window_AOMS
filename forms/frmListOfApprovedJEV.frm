VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmListOfApprovedJEV 
   Caption         =   "List of Approved JEV"
   ClientHeight    =   7710
   ClientLeft      =   2910
   ClientTop       =   2250
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   13890
   Begin VB.CommandButton Command3 
      Caption         =   "Show"
      Height          =   1110
      Left            =   7365
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1200
      Left            =   3915
      TabIndex        =   10
      Top             =   150
      Width           =   3285
      Begin VB.ComboBox cmb_Fund 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   11
         Top             =   585
         Width           =   2595
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   555
      Left            =   11940
      TabIndex        =   9
      Top             =   915
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   555
      Left            =   11940
      TabIndex        =   8
      Top             =   300
      Width           =   1500
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5310
      Left            =   225
      TabIndex        =   6
      Top             =   1935
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   9366
      _Version        =   393216
      FixedCols       =   0
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
   Begin VB.Frame Frame1 
      Caption         =   "Approval Scope"
      Height          =   1200
      Left            =   195
      TabIndex        =   0
      Top             =   150
      Width           =   3585
      Begin VB.ComboBox cmb_trnYear 
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
         Left            =   2175
         TabIndex        =   4
         Top             =   585
         Width           =   1155
      End
      Begin VB.ComboBox cmb_month 
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
         Left            =   240
         TabIndex        =   3
         Top             =   585
         Width           =   1635
      End
      Begin VB.OptionButton opn_quarterly 
         Caption         =   "Quarter"
         Height          =   195
         Left            =   1050
         TabIndex        =   2
         Top             =   300
         Width           =   900
      End
      Begin VB.OptionButton opn_monthly 
         Caption         =   "Month"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2175
         TabIndex        =   5
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approved JEV"
      Height          =   195
      Left            =   225
      TabIndex        =   7
      Top             =   1665
      Width           =   1020
   End
End
Attribute VB_Name = "frmListOfApprovedJEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim opnJEVs As New ADODB.Recordset
Dim sql As String

If opn_quarterly.Value = True Then 'by quarter
    sql = "SELECT  tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.* " & _
            " FROM  tblAMIS_IncomingDVTrns INNER JOIN " & _
            "       tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
            " WHERE  tblAMIS_IncomingDVTrns.FundType = '" & cmb_Fund.Text & "' and year(tblAMIS_JournalEntry.DateTimeApproved)=" & cmb_trnYear.Text & " and " & _
            "       month(tblAMIS_JournalEntry.DateTimeApproved) in (" & GetQuarterMonthsInAyear(cmb_month.ItemData(cmb_month.ListIndex)) & ") and " & _
            "       (dbo.tblAMIS_JournalEntry.Actioncode = 1) AND (dbo.tblAMIS_IncomingDVTrns.Actioncode = 1)"

ElseIf opn_monthly.Value = True Then 'by month
    sql = "SELECT  tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.* " & _
            " FROM  tblAMIS_IncomingDVTrns INNER JOIN " & _
            "       tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
            " WHERE  tblAMIS_IncomingDVTrns.FundType = '" & cmb_Fund.Text & "' and year(tblAMIS_JournalEntry.DateTimeApproved)=" & cmb_trnYear.Text & " and " & _
            "       month(tblAMIS_JournalEntry.DateTimeApproved)=" & cmb_month.ItemData(cmb_month.ListIndex) & " and  " & _
            "       (dbo.tblAMIS_JournalEntry.Actioncode = 1) AND (dbo.tblAMIS_IncomingDVTrns.Actioncode = 1)"
End If
'Debug.Print sql

opnJEVs.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnJEVs.RecordCount <> 0 Then
    Set MSHFlexGrid1.DataSource = opnJEVs
Else
    Set MSHFlexGrid1.DataSource = opnJEVs
End If
opnJEVs.Close
Set opnJEVs = Nothing


End Sub
Public Function GetQuarterMonthsInAyear(ByVal Quarter As Integer) As String
Select Case Quarter
    Case 1:
        GetQuarterMonthsInAyear = "1,2,3"
    Case 2:
        GetQuarterMonthsInAyear = "4,5,6"
    Case 3:
        GetQuarterMonthsInAyear = "7,8,9"
    Case 4:
        GetQuarterMonthsInAyear = "10,11,12"
End Select

End Function
Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Call LoadTrnYear(cmb_trnYear)
Call SelectionScope
Call LoadFundType

End Sub
Private Sub LoadFundType()
Dim opnfund As New ADODB.Recordset
Dim cc As Integer

opnfund.Open "Select FundType from tblAMIS_IncomingDVTrns group by FundType order by FundType", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    cmb_Fund.Clear
    Do Until opnfund.EOF
        cmb_Fund.AddItem (opnfund!FundType)
        'cmb_Fund.ItemData(cc) = opnfund!Fundcode
        'cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_Fund.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub
Private Sub SelectionScope()
If opn_monthly.Value = True Then
    Call LoadTrnMonth
ElseIf opn_quarterly.Value = True Then
    Call LoadQuarterMonths
End If
End Sub
Private Sub LoadQuarterMonths()

cmb_month.Clear
cmb_month.AddItem ("1st")
cmb_month.ItemData(0) = 1

cmb_month.AddItem ("2nd")
cmb_month.ItemData(1) = 2

cmb_month.AddItem ("3rd")
cmb_month.ItemData(2) = 3

cmb_month.AddItem ("4th")
cmb_month.ItemData(3) = 4

cmb_month.ListIndex = GetIndex(cmb_month, GetQuarterAffiliation(Month(Date)))

End Sub
Private Sub LoadTrnMonth()
Dim cc As Integer
Dim xx As Integer

cmb_month.Clear
For cc = 1 To 12
    cmb_month.AddItem (GetMonthMedium(cc))
    cmb_month.ItemData(xx) = cc
    xx = xx + 1
Next cc
cmb_month.ListIndex = GetIndex(cmb_month, GetMonthMedium(Month(Date)))

End Sub

Private Sub opn_monthly_Click()
Call SelectionScope
End Sub

Private Sub opn_quarterly_Click()
Call SelectionScope
End Sub
