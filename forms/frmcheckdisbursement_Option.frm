VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmcheckdisbursement_Option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Disbursements Journal"
   ClientHeight    =   5325
   ClientLeft      =   6345
   ClientTop       =   4380
   ClientWidth     =   4110
   Icon            =   "frmcheckdisbursement_Option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmcheckdisbursement_Option.frx":09EA
   ScaleHeight     =   3.698
   ScaleMode       =   5  'Inch
   ScaleWidth      =   2.854
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
      Begin VB.CheckBox chkRecap 
         Caption         =   "With Recap"
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
         TabIndex        =   17
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Recap Criteria"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   3375
         Begin VB.OptionButton OptDetailed 
            Caption         =   "Detailed"
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
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton OptConso 
            Caption         =   "Consolidated"
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
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.ComboBox cmb_FundType 
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
         ItemData        =   "frmcheckdisbursement_Option.frx":B099
         Left            =   240
         List            =   "frmcheckdisbursement_Option.frx":B0A6
         TabIndex        =   8
         Text            =   "cmb"
         Top             =   600
         Width           =   3420
      End
      Begin VB.CheckBox chkConsolidated 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1530
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Group by Bank"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   9
         Top             =   1365
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
         Format          =   169345027
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fund type"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Account Number:"
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   8640
      TabIndex        =   0
      Top             =   4320
      Width           =   3780
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
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   3435
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin lvButton.lvButtons_H Command1 
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "View"
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
      Image           =   "frmcheckdisbursement_Option.frx":B0DC
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   4680
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
      Image           =   "frmcheckdisbursement_Option.frx":BAD6
      cBack           =   16777215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Select Special Accouts and set the period that you want to print"
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
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Disbursement Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -360
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmcheckdisbursement_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub LoadSavedReport(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal FTYPE As String)
Dim opnRCINo As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
Dim frm As New frm_RPTviewer

With frm
If Check1.Value = 0 Then
.NoGroup = True
Else
.NoGroup = False
End If

 .Jquery = "Exec MPproc_RPTJournals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 2"

If chkRecap.Value = 1 Then
.WRecap = True
    If OptConso.Value = True Then
        .Rquery = "Exec MPproc_RPTRecap_Journals_Conso @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 2"
    Else
        .Rquery = "Exec MPproc_RPTRecap_Journals @month = " & trnMonth & ",@year = " & TrnYear & ",@fundtype = '" & FTYPE & "',@Transtype = 2"
    End If
Else
.WRecap = False
End If

If chkConsolidated.Value = 1 Then
   .Ifconso = True
Else
   .Ifconso = False
End If

.mnth = "Month: " & Format(DTPicker2.Value, "mmmm") & " " & DTPicker2.Year
.fund = Trim(cmb_fundtype.Text)
.TrnsType = 2
Set .frm = Me
.Show
End With
Exit Sub
bad:
    If err.Number = 364 Then
    Else
    Call LoadErr(err.Number, Me.name, err.description)
    End If
End Sub



Private Sub cmb_Fund_click()
'''Call LoadBankAccntNo(Combo1)
End Sub

Private Sub chkConsolidated_Click()
If chkConsolidated.Value = 1 Then
    Call LoadMotherFund(cmb_fundtype)
Else
    Call Form_Load
End If
End Sub

Private Sub chkRecap_Click()
If chkRecap.Value = 1 Then
Frame2.Enabled = True
Else
Frame2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Call PlayAVI(Me.Animation1, "Refresh.avi")
If chkConsolidated.Value = 1 Then
    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_fundtype.ItemData(cmb_fundtype.ListIndex))
Else
    Call LoadSavedReport(DTPicker2.Month, DTPicker2.Year, cmb_fundtype.ItemData(cmb_fundtype.ListIndex))
End If
Call StopAvi(Me.Animation1)
End Sub

Private Sub Form_Load()
  Call LoadFundType(cmb_fundtype)
  Call chkRecap_Click
  DTPicker2.Value = Now
End Sub


Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub Option2_Click()

End Sub
