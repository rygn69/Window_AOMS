VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_SE_Option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5865
   ClientLeft      =   5730
   ClientTop       =   3840
   ClientWidth     =   9165
   Icon            =   "frm_SE_Option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_SE_Option.frx":076A
   ScaleHeight     =   5865
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   2040
      Width           =   5055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6465
      ScaleWidth      =   8610
      TabIndex        =   8
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
         TabIndex        =   11
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtdetails 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5175
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   13
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
         Image           =   "frm_SE_Option.frx":AE19
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   5280
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
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
         ItemData        =   "frm_SE_Option.frx":E923
         Left            =   240
         List            =   "frm_SE_Option.frx":E930
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
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
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   3300
         _ExtentX        =   5821
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
         Format          =   121241601
         UpDown          =   -1  'True
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   3315
         _ExtentX        =   5847
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
         Format          =   121241601
         UpDown          =   -1  'True
         CurrentDate     =   40574
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "To"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2040
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
         Left            =   240
         TabIndex        =   4
         Top             =   1155
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   5160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
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
      Image           =   "frm_SE_Option.frx":E966
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H10 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Select top 34"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Image           =   "frm_SE_Option.frx":F360
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H9 
      Height          =   375
      Left            =   1725
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Deselect all"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Image           =   "frm_SE_Option.frx":FFB2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   8040
      TabIndex        =   23
      Top             =   5160
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
      Image           =   "frm_SE_Option.frx":10304
      cBack           =   16777215
   End
   Begin VB.Label lblException 
      BackStyle       =   0  'Transparent
      Caption         =   "Status of Expenses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: maximum selection of offices allowed is only 34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8835
   End
End
Attribute VB_Name = "frm_SE_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Private Sub LoadSavedReport()
Dim frm As New frm_RPTQueryViewer
On Error GoTo bad
Dim sql As String
    Report9 = "SE"
    With frm
    .accnt = "EXECUTE [fmis].[dbo].[MPproc_New_RPT_Query] @from = '" & Format(DTPicker2.Value, "mm/dd/yyyy") & "',@to = '" & Format(DTPicker1.Value, "mm/dd/yyyy") & "',@reports = 'SE',@fundcode = '" & cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "'"
    .dated = "From " & Format(DTPicker2.Value, "MMMM dd, yyyy") & " To " & Format(DTPicker1.Value, "MMMM dd, yyyy")
    .Show
    End With
Exit Sub
bad:
    If err.Number = 364 Then
    MsgBox "No Record Found..", vbInformation, "System Message"
    Else
    MsgBox err.description
    End If
End Sub
Private Sub LoadOffices()
Dim OREc As New ADODB.Recordset
Dim x As Integer

List1.Clear
OREc.Open ("SELECT cast(RCenter as int) as Rcenter,OfficeMedium FROM [fmis].[dbo].[tblAMIS_FinalJEV] inner join dbo.tblREF_AIS_Offices as b on RCenter = b.FMISOfficeID where RCenter != '' and year(Jevdate ) = 2013  group by cast(RCenter as int),OfficeMedium order by OfficeMedium"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        List1.AddItem OREc![OfficeMedium]
        List1.ItemData(List1.NewIndex) = OREc!RCenter
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing
End Sub

Private Sub chkConsolidated_Click()
If chkConsolidated.Value = 1 Then
Call LoadMotherFund(cmb_fundtype)
Else
Call LoadFundType(cmb_fundtype)
End If
End Sub

Private Sub Command1_Click()
'    If CheckIfOverSelected = True Then
'    MsgBox "Maximum selection of offices allowed is only 34"
'    Exit Sub
'    End If
    
    Call InsertRcenter
    Call LoadSavedReport
    
End Sub
Private Sub InsertRcenter()
Dim x, y As Integer
opndbaseFMIS.Execute "delete from tblAMIS_SelectedRcenter"
For x = 0 To List1.ListCount - 1

    If List1.Selected(x) = True Then
        opndbaseFMIS.Execute "insert into tblAMIS_SelectedRcenter(rcenter) values (" & List1.ItemData(x) & ")"
    End If
    
Next x
opndbaseFMIS.Execute "Execute MPproc_NeedToExecute @type = 2"
End Sub
Private Function CheckIfOverSelected() As Boolean
Dim x, y As Integer
CheckIfOverSelected = False
For x = 0 To List1.ListCount - 1
    If List1.Selected(x) = True Then
        y = y + 1
        If y > 34 Then
            CheckIfOverSelected = True
            Exit For
        End If
    End If
Next x
End Function
'Private Sub CheckList(ByVal typ As Integer)
'Dim x, y As Integer
'CheckIfOverSelected = False
'For x = 0 To List1.ListCount - 1
'    If List1.Selected(x) = True Then
'        y = y + 1
'        If y > 34 Then
'            CheckIfOverSelected = True
'            Exit For
'        End If
'    End If
'Next x
'End Sub
Private Sub Form_Load()
Call LoadFundType(cmb_fundtype)
Call LoadOffices
DTPicker1.Value = Now
DTPicker2.Value = Now
End Sub

Private Sub List1_Click()
'MsgBox List1.ItemData(List1.ListIndex)
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H10_Click()
Call CheckUncheckList(1)
End Sub
Private Sub CheckUncheckList(ByVal typ As Integer)
Dim x As Integer
For x = 0 To List1.ListCount - 1
    If typ = 1 Then
        List1.Selected(x) = True
    Else
        List1.Selected(x) = False
    End If
    DoEvents
Next x
End Sub

Private Sub lvButtons_H9_Click()
Call CheckUncheckList(2)
End Sub
