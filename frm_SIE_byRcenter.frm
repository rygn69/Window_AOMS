VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SIE_byRcenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statement of Income and Expense"
   ClientHeight    =   4965
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   3900
   Icon            =   "frm_SIE_byRcenter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_SIE_byRcenter.frx":076A
   ScaleHeight     =   4965
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6465
      ScaleWidth      =   8610
      TabIndex        =   13
      Top             =   9840
      Visible         =   0   'False
      Width           =   8635
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
         TabIndex        =   16
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtdetails 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   75
         Width           =   7215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Many"
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
         Left            =   120
         TabIndex        =   14
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5175
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   9128
         _Version        =   393216
         BackColor       =   16777215
         BackColorSel    =   8454143
         ForeColorSel    =   0
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
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   8160
         TabIndex        =   18
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frm_SIE_byRcenter.frx":AE19
         cBack           =   16777215
      End
      Begin VB.Label Label8 
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
         TabIndex        =   21
         Top             =   5925
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         TabIndex        =   20
         Top             =   5835
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Details:"
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
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
      Begin VB.OptionButton Option3 
         Caption         =   "Current Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   1680
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Continuing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkcomparative 
         Caption         =   "With Comparative Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkRC 
         Caption         =   "By Responsibility Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox cmbRC 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_SIE_byRcenter.frx":E923
         Left            =   120
         List            =   "frm_SIE_byRcenter.frx":E930
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2520
         Width           =   3420
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   360
         Width           =   1530
      End
      Begin VB.ComboBox cmb_Accountcode 
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
         ItemData        =   "frm_SIE_byRcenter.frx":E966
         Left            =   0
         List            =   "frm_SIE_byRcenter.frx":E968
         TabIndex        =   5
         Top             =   3480
         Width           =   3465
      End
      Begin VB.ComboBox cmb_FundType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_SIE_byRcenter.frx":E96A
         Left            =   120
         List            =   "frm_SIE_byRcenter.frx":E977
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3420
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   360
         TabIndex        =   8
         Top             =   3840
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   297205761
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM dd,  yyyy"
         Format          =   297205763
         UpDown          =   -1  'True
         CurrentDate     =   40574
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Accountcode"
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
         Left            =   0
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fund type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton dsf 
      Caption         =   "&View Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frm_SIE_byRcenter.frx":E9AD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&View"
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
      Image           =   "frm_SIE_byRcenter.frx":F117
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
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
      Image           =   "frm_SIE_byRcenter.frx":FB11
      cBack           =   16777215
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statement of Income and Expense"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Define criteria prior and click view button to display the details."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   3675
   End
End
Attribute VB_Name = "frm_SIE_byRcenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim IEE, IEE100, IEE200, IEE300 As Currency
Private Sub LoadSavedReport()
Dim frm As New frma_RPTfinancialViewer
On Error GoTo bad
Dim sql As String
Dim RC As String
Dim Continuing As Integer
Report9 = "SIE"


If Option1.Value = True Then
    Continuing = 1
End If

If Option2.Value = True Then
    Continuing = 3
End If

If Option3.Value = True Then
    Continuing = 0
End If

If chkRC.Value = 1 Or Check2.Value = 1 Then
    If Check2.Value = 1 Then
        RC = ""
    Else
        RC = cmbRC.ItemData(cmbRC.ListIndex)
    End If
    With frm
    .accnt = "EXECUTE Accounting.[ufn_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_SUB',@Continuing = " & Continuing & ""
    .query = "EXEC [fmis].[dbo].[MPproc_RPT_Financials_byRCenter] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundtype ='" & cmb_FundType.Text & "',@reports = 'SIE',@RCenter = '" & Trim(RC) & "'"
    .dated = "For the month ended " & Format(DTPicker1.Value, "MMMM dd, yyyy")
    If chkConsolidated.Value = 1 Then
        .FundType = Trim(cmb_FundType.Text) & " Consolidated (" & cmbRC.Text & ")"
    Else
        .FundType = Trim(cmb_FundType.Text) & " (" & cmbRC.Text & ")"
    End If
    .ComparativeDate = DTPicker1.Value
    .Show
    End With
Else
    With frm
    If cmb_FundType.ItemData(cmb_FundType.ListIndex) = 0 Then
        .AllFunds = True
    Else
        .AllFunds = False
    End If
    
    .query = "EXECUTE Accounting.[ufn_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE',@Continuing = " & Continuing & ""
    .accnt = "EXECUTE Accounting.[ufn_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_SUB',@Continuing = " & Continuing & ""
    If chkcomparative.Value = 1 Then
        Call GetTotalSIE_Comparative
        .query = "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_Comparative'"
        .accnt = "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_SUB_Comparative'"
        .Comparative = True
    Else
        Call GetTotalSIE
        .query = "EXECUTE Accounting.[ufn_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE' ,@Continuing = " & Continuing & ""
        .accnt = "EXECUTE Accounting.[ufn_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_SUB',@Continuing = " & Continuing & ""
        .Comparative = False
    End If
    .IEE = IEE
    .IEE100 = IEE100
    .IEE200 = IEE200
    .IEE300 = IEE300
    '.query = "EXEC [fmis].[dbo].[MPproc_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundtype ='" & cmb_FundType.Text & "',@reports = 'SIE'"
    If DTPicker1.Month = 12 Then
        .dated = "For The Year Ended " & Format(DTPicker1.Value, "MMMM dd, yyyy")
    Else
        .dated = "For The Month Ended " & Format(DTPicker1.Value, "MMMM dd, yyyy")
    End If
    If chkConsolidated.Value = 1 Then
        .FundType = Trim(cmb_FundType.Text) & " Consolidated"
    Else
        .FundType = Trim(cmb_FundType.Text)
    End If
    .ComparativeDate = DTPicker1.Value
    .Show
    End With
End If
Set frm = Nothing
Exit Sub
bad:
    If err.Number = 364 Then
    MsgBox "No Record Found..", vbInformation, "System Message"
    Else
    MsgBox err.description
    End If
End Sub


Public Sub LoadFund()
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
                
opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 
If opnfund.RecordCount <> 0 Then
    cmb_FundType.Clear
    Do Until opnfund.EOF
        cmb_FundType.AddItem (opnfund!FundName)
        cmb_FundType.ItemData(cc) = opnfund!fundcode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_FundType.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub
Public Sub GetTotalSIE()
Dim rec As New ADODB.Recordset
Dim cc As Integer
rec.Open "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_Total'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount <> 0 Then
    IEE = IIf(IsNull(rec!amount), 0, rec!amount)
    IEE100 = IIf(IsNull(rec!amount100), 0, rec!amount100)
    IEE200 = IIf(IsNull(rec!amount200), 0, rec!amount200)
    IEE300 = IIf(IsNull(rec!amount300), 0, rec!amount300)
End If
rec.Close
Set rec = Nothing
End Sub
Public Sub GetTotalSIE_Comparative()
Dim rec As New ADODB.Recordset
Dim cc As Integer
rec.Open "EXECUTE [fmis].[dbo].[MPproc_new_RPT_Financials] @from ='" & DTPicker2.Value & "',@to='" & DTPicker1.Value & "',@Accountcode ='',@Fundcode ='" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@reports = 'SIE_Total_Comparative'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount <> 0 Then
    IEE = IIf(IsNull(rec!amount), 0, rec!amount)
    IEE100 = IIf(IsNull(rec!beginamount), 0, rec!beginamount)
End If
rec.Close
Set rec = Nothing
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
chkRC.Value = 0
cmbRC.Enabled = False
cmbRC.Clear
End If
End Sub

Private Sub chkConsolidated_Click()
If chkConsolidated.Value = 1 Then
Call LoadMotherFund(cmb_FundType)
Else
Call LoadFundType(cmb_FundType)
End If
End Sub

Private Sub chkRC_Click()
If chkRC.Value = 1 Then
    Check2.Value = 0
    cmbRC.Enabled = True
    Call LoadRC(cmbRC)
Else
    cmbRC.Enabled = False
    cmbRC.Clear
End If
End Sub

Private Sub cmb_Accountcode_KeyPress(KeyAscii As Integer)
KeyAscii = AutoFind(cmb_Accountcode, KeyAscii, True)
End Sub

Private Sub cmb_FundType_Change()
'Call loadChildAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub Command1_Click()
Call PlayAVI(Me.Animation1, "Refresh.avi")
  Call LoadSavedReport
Call StopAvi(Me.Animation1)
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
DTPicker1.Value = Now
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
