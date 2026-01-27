VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_LocateTrans 
   Caption         =   "DV Search"
   ClientHeight    =   8805
   ClientLeft      =   1440
   ClientTop       =   1905
   ClientWidth     =   13905
   Icon            =   "frm_Locatetrans.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbRC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      TabIndex        =   8
      Text            =   "cmbRC"
      Top             =   1680
      Visible         =   0   'False
      Width           =   13455
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   13455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResult 
      Height          =   5895
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   13380
         Begin VB.OptionButton opn_Employee 
            Caption         =   "Capitol Employee"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            TabIndex        =   18
            Top             =   330
            Width           =   1830
         End
         Begin VB.OptionButton opn_Individual 
            Caption         =   "Other Individual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   17
            Top             =   330
            Width           =   1710
         End
         Begin VB.OptionButton opn_company 
            Caption         =   "Company"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3840
            TabIndex        =   16
            Top             =   330
            Width           =   1140
         End
         Begin VB.OptionButton opn_offices 
            Caption         =   "Provincial Offices"
            Height          =   195
            Left            =   5745
            TabIndex        =   15
            Top             =   330
            Width           =   1680
         End
         Begin VB.OptionButton opn_National 
            Caption         =   "National Offices"
            Height          =   195
            Left            =   7725
            TabIndex        =   14
            Top             =   330
            Width           =   1455
         End
         Begin VB.OptionButton opn_BT 
            Caption         =   "Barangay Treasurers"
            Height          =   195
            Left            =   9495
            TabIndex        =   13
            Top             =   330
            Width           =   1845
         End
         Begin VB.OptionButton opn_MT 
            Caption         =   "Municipal Treasurers"
            Height          =   195
            Left            =   11400
            TabIndex        =   12
            Top             =   330
            Width           =   1845
         End
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Claimant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Responsibility Center"
         Height          =   375
         Index           =   4
         Left            =   11475
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Amount"
         Height          =   375
         Index           =   3
         Left            =   9630
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Details"
         Height          =   375
         Index           =   2
         Left            =   8010
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "ObR Number"
         Height          =   375
         Index           =   1
         Left            =   5925
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "DV Number"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbClaimant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   13455
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   8400
      Width           =   6615
   End
   Begin VB.Menu men 
      Caption         =   ""
      Begin VB.Menu x 
         Caption         =   " "
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frm_LocateTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClaimant_Click()
    Call txtSearch_KeyPress(13)
End Sub

Private Sub cmbClaimant_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoFind(cmbClaimant, KeyAscii, True)
End Sub

Private Sub cmbRC_Click()
    Call txtSearch_KeyPress(13)
End Sub

Private Sub cmbrc_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoFind(cmbRC, KeyAscii, True)
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call LoadOffice
    optSearch(5).Value = True
End Sub

Private Sub LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmbRC.Clear

OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmbRC.AddItem OREc![OfficeMedium]
        cmbRC.ItemData(cmbRC.NewIndex) = OREc!fmisofficeid
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing

End Sub

Private Sub Form_Resize()
On Error Resume Next
grdResult.Width = Me.ScaleWidth - 350
  grdResult.Height = Me.ScaleHeight - grdResult.Top - 500
End Sub

Private Sub opn_BT_Click()
    Call LoadClaimant("BT")
End Sub

Private Sub opn_company_Click()
    Call LoadClaimant("CO")
End Sub

Private Sub opn_Employee_Click()
    Call LoadClaimant("CE")
End Sub


Private Sub LoadClaimant(ByVal ClaimantType As String)
Dim ERec As New ADODB.Recordset
Dim x As Integer

    cmbClaimant.Clear
    grdResult.Clear
    lblMessage.Caption = ""
    Select Case ClaimantType
    Case "CE": 'Capitol Employee
                ERec.Open ("Select * From pmis.dbo.Employee where len(SwipEmployeeID)>0 Order by Lastname, Firstname, MI, suffix"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname & ", " & ERec!Firstname & " " & ERec!mi & " " & ERec!Suffix
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = ERec![swipemployeeid]
                        ERec.MoveNext
                        DoEvents
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "OI": 'Other Individual
                Set ERec = opndbaseFMIS.Execute("Select lastname,firstname,mi,suffix,ClaimantCode from tblCMS_CDClaimantDetails where left(ClaimantCode,1) = 'O' order by lastname, firstname, mi, suffix")
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname & ", " & ERec!Firstname & " " & ERec!mi & " " & ERec!Suffix
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = Val(Mid(ERec![ClaimantCode], 2, 4))
                        ERec.MoveNext
                        DoEvents
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "PO": 'Provincial Offices
                ERec.Open "Select * from tblREF_AIS_Offices order by OfficeMedium", opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount <> 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem (ERec!OfficeMedium)
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = ERec!fmisofficeid
                        ERec.MoveNext
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "NO": 'National Offices
                ERec.Open ("Select * from tblCMS_CDClaimantDetails where ClaimantCode like 'N%' order by lastname"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = Val(Mid(ERec![ClaimantCode], 2, 4))
                        ERec.MoveNext
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "CO": 'Company
                ERec.Open ("Select * from tblCMS_CDClaimantDetails where ClaimantCode like 'C%' order by lastname"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = Val(Mid(ERec![ClaimantCode], 2, 4))
                        ERec.MoveNext
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "BT": 'Barangay Treasurers
                ERec.Open ("Select * from tblCMS_CDClaimantDetails where ClaimantCode like 'BT%' order by lastname"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = Val(Mid(ERec![ClaimantCode], 3, 4))
                        ERec.MoveNext
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    Case "MT": 'Municipal Treasurers
                ERec.Open ("Select * from tblCMS_CDClaimantDetails where ClaimantCode like 'MT%' order by lastname"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If ERec.RecordCount > 0 Then
                    For x = 1 To ERec.RecordCount
                        cmbClaimant.AddItem ERec!lastname
                        cmbClaimant.ItemData(cmbClaimant.NewIndex) = Val(Mid(ERec![ClaimantCode], 3, 4))
                        ERec.MoveNext
                    Next x
                End If
                ERec.Close
                Set ERec = Nothing
    End Select
'    cmbClaimant.SetFocus
    
    
End Sub

Private Sub opn_Individual_Click()
    Call LoadClaimant("OI")
End Sub

Private Sub opn_MT_Click()
    Call LoadClaimant("MT")
End Sub

Private Sub opn_National_Click()
    Call LoadClaimant("NO")
End Sub

Private Sub opn_offices_Click()
    Call LoadClaimant("PO")
End Sub

Private Sub optSearch_Click(Index As Integer)
    grdResult.Clear
    lblMessage.Caption = ""
    If Index = 4 Then
        txtSearch.Text = ""
        txtSearch.Visible = False
        Frame2.Visible = False
        cmbRC.Visible = True
        cmbClaimant.Visible = False
        opn_Employee = True
        cmbRC.SetFocus
    ElseIf Index = 5 Then
        txtSearch.Text = ""
        txtSearch.Visible = False
        Frame2.Visible = True
        cmbClaimant.Visible = True
        cmbRC.Visible = False
        opn_Employee.Value = True
        Call opn_Employee_Click
    Else
        cmbRC.Visible = False
        cmbClaimant.Visible = False
        Frame2.Visible = False
        txtSearch.Visible = True
        txtSearch.SetFocus
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
Dim SRec As New ADODB.Recordset
Dim sql As String
Dim x As Integer

    If KeyAscii = 13 Then
        
        grdResult.Clear
        lblMessage.Caption = ""
        
        If optSearch(0).Value = True Then 'Search DV
            sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
            ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
            "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where DVNo='" & Trim(Replace(txtSearch.Text, "'", "''")) & "' and actioncode=1 order by trnno desc"
        ElseIf optSearch(1).Value = True Then 'Search ObR
            sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
            ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
            "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where ObrNo='" & Trim(Replace(txtSearch.Text, "'", "''")) & "' and actioncode=1 order by trnno desc"
        ElseIf optSearch(2).Value = True Then 'Search Details
            sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
            ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
            "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where Particular like '%" & Trim(Replace(txtSearch.Text, "'", "''")) & "%' and actioncode=1 order by trnno desc"
        ElseIf optSearch(3).Value = True Then 'Search Amount
            If IsNumeric(txtSearch.Text) = False Then
                txtSearch.Text = 0
            End If
            sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
            ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
            "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where GAmount=" & Format(CCur(txtSearch.Text), "####.00") & " and actioncode=1 order by trnno desc"
        ElseIf optSearch(4).Value = True Then 'Search Responsibility Center
            If cmbRC.ListIndex <> -1 Then
                sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [RCenter]=" & cmbRC.ItemData(cmbRC.ListIndex) & " and actioncode=1 order by trnno desc"
            End If
        ElseIf optSearch(5).Value = True Then 'Search Claimant
            If cmbClaimant.ListIndex <> -1 Then
                If opn_Employee.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                ElseIf opn_Individual.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='O" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                ElseIf opn_offices.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='" & cmbClaimant.ItemData(cmbClaimant.ListIndex) & "' and actioncode=1 order by trnno desc"
                ElseIf opn_National.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='N" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                ElseIf opn_company.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='C" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                ElseIf opn_BT.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='BT" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                ElseIf opn_MT.Value = True Then
                    sql = "Select [DVNo] as 'DV No',[ObrNo] as 'ObR No',[FundType] as 'Fund Type',(Select OfficeMedium from tblREF_AIS_Offices where FMISOfficeID=[RCenter]) as 'Responsibility Center',[ClaimantCode] as 'Claimant',[Particular],[GAmount] as 'Amount',[TransactionDate] as 'In',[PAoutDate] as 'PREAUDIT (OUT)',[PADesc] as 'Remarks'" & _
                    ", (Select top 1 [DateTimeApproved] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Prepared', (Select top 1 [LogOutDateTime] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'Approve Out', (Select top 1 [LogOutRemark] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'REMARKS', " & _
                    "(Select top 1 [JEVDate] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV Entry', (Select top 1 [JEVSeriesNo] from [tblAMIS_JournalEntry] Where DVNo=tblAMIS_IncomingDVTrns.DVNo and ActionCode=1) as 'JEV No' From tblAMIS_IncomingDVTrns where [ClaimantCode]='MT" & Format(cmbClaimant.ItemData(cmbClaimant.ListIndex), "000#") & "' and actioncode=1 order by trnno desc"
                End If
            End If
        End If

        
        SRec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If SRec.RecordCount > 0 Then
            Set grdResult.Recordset = SRec
            Set grdResult.Recordset = Nothing
            
            grdResult.ColWidth(0) = 0
            grdResult.ColWidth(1) = 1300
            grdResult.ColWidth(2) = 1700
            grdResult.ColWidth(3) = 0
            grdResult.ColWidth(4) = 1500
            grdResult.ColWidth(5) = 2300
            grdResult.ColWidth(6) = 5000
            grdResult.ColWidth(7) = 1300
            grdResult.ColWidth(8) = 0
            grdResult.ColWidth(9) = 1900
            grdResult.ColWidth(10) = 0
            grdResult.ColWidth(11) = 1600
            grdResult.ColWidth(12) = 1800
            grdResult.ColWidth(13) = 700
            grdResult.ColWidth(14) = 0
            grdResult.ColWidth(15) = 0
            
            
            
            lblMessage.Caption = SRec.RecordCount & " Records Found!"
            
            grdResult.Row = 0
            For x = 1 To grdResult.Cols - 1
                grdResult.col = x
                grdResult.CellAlignment = 4
            Next x
            
            For x = 1 To grdResult.Rows - 1
                grdResult.TextMatrix(x, 5) = getClaimant(grdResult.TextMatrix(x, 5))
                grdResult.TextMatrix(x, 7) = Format(grdResult.TextMatrix(x, 7), "#,###.00")
            Next x
        Else
            lblMessage.Caption = SRec.RecordCount & " Records Found!"
        End If
        SRec.Close
        Set SRec = Nothing
        
    End If

End Sub
Private Sub x_Click()
Unload Me
End Sub
