VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Audited 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre-Audit Assigment"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   1455
   ClientWidth     =   14505
   ForeColor       =   &H00000000&
   Icon            =   "frm_Audited.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   14505
   Begin VB.ComboBox cmb_trans_class 
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
      ItemData        =   "frm_Audited.frx":076A
      Left            =   7455
      List            =   "frm_Audited.frx":0777
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2160
      Width           =   4470
   End
   Begin VB.ListBox lst_DVNO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   12120
      TabIndex        =   35
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cmb_preAudit 
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
      Left            =   7500
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3120
      Width           =   4470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Audit IN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7560
      TabIndex        =   31
      Top             =   360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkAudit 
      Caption         =   "Audit OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9120
      TabIndex        =   30
      Top             =   360
      Width           =   1455
   End
   Begin VB.CheckBox chkApprove 
      BackColor       =   &H00000000&
      Caption         =   "Approve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   12240
      TabIndex        =   29
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CheckBox chkReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbapprove 
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
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   8160
      Width           =   4470
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7500
      TabIndex        =   20
      Top             =   1215
      Width           =   4470
   End
   Begin VB.ComboBox cmbaudit 
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
      Left            =   7500
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3960
      Width           =   4470
   End
   Begin VB.TextBox txt_Remark 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7485
      MaxLength       =   50
      TabIndex        =   18
      Top             =   4905
      Width           =   4500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
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
      Height          =   5805
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7155
      Begin VB.TextBox txtFund 
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4500
         Width           =   4500
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5205
         Width           =   4500
      End
      Begin VB.TextBox txtParticular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   2355
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2385
         Width           =   4500
      End
      Begin VB.TextBox txtAlobs 
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   300
         Width           =   4500
      End
      Begin VB.TextBox txtClaimant 
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   4500
      End
      Begin VB.TextBox txtRC 
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1620
         Width           =   4500
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
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
         Height          =   240
         Left            =   1110
         TabIndex        =   15
         Top             =   4590
         Width           =   1020
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross):"
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
         Height          =   240
         Left            =   705
         TabIndex        =   13
         Top             =   5370
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
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
         Height          =   240
         Left            =   1305
         TabIndex        =   11
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alobs/OBR No:"
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
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   390
         Width           =   1950
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant:"
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
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center:"
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
         Height          =   240
         Left            =   300
         TabIndex        =   6
         Top             =   1710
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   45
      ScaleHeight     =   4200
      ScaleWidth      =   14100
      TabIndex        =   0
      Top             =   8850
      Width           =   14130
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   7408
         _Version        =   393216
         FixedCols       =   0
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
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   735
      Left            =   10920
      TabIndex        =   33
      Top             =   6000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      Caption         =   "&Save Changes"
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
      Image           =   "frm_Audited.frx":0796
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   735
      Left            =   13080
      TabIndex        =   34
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
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
      Image           =   "frm_Audited.frx":1190
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   12120
      TabIndex        =   36
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "&Refresh"
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
      Image           =   "frm_Audited.frx":4C9A
      cBack           =   16777215
   End
   Begin VB.Label Label3 
      Caption         =   "Claims Classification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   38
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Approved By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter DV Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7485
      TabIndex        =   25
      Top             =   840
      Width           =   1950
   End
   Begin VB.Label Label15 
      Caption         =   "Audited By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7485
      TabIndex        =   24
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7440
      TabIndex        =   23
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter DV Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7485
      TabIndex        =   22
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label Label10 
      Caption         =   "Asigned Auditor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7485
      TabIndex        =   21
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5670
      TabIndex        =   16
      Top             =   -315
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   8580
      Width           =   1335
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Disbursement Voucher No :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      Top             =   -480
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frm_Audited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Edited As Boolean
Dim xDebit As Currency
Dim xCredit As Currency
Dim xObR As String
Dim xNAcode As String
Private Sub btnPrtJEV_Click()
Unload Me
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
        cmb_preAudit.Enabled = True
        cmbaudit.Enabled = False
        chkAudit.Value = 0
    End If
End Sub

Private Sub chkAudit_Click()
If chkAudit.Value = 1 Then
        cmb_preAudit.Enabled = False
        cmbaudit.Enabled = True
        Check1.Value = 0
    End If
End Sub

Private Sub chkReturn_Click()
If chkReturn.Value = 1 Then
        'cmb_preAudit.Value = 0
        cmb_preAudit.Enabled = False
'        cmbapprove.ListIndex = 0
        cmbapprove.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub LoadDVNO()
Dim opnaccnt As New ADODB.Recordset
lst_DVNO.Clear
opnaccnt.Open "Select dvno from  dbo.ufn_getDVNOForAudit()", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnaccnt.RecordCount <> 0 Then
    Do Until opnaccnt.EOF
        lst_DVNO.AddItem (Trim(opnaccnt!dvno))
        opnaccnt.MoveNext
    Loop
End If
opnaccnt.Close
Set opnaccnt = Nothing
'Label11.Caption = lst_DVNO.ListCount & " Accnt. Advice/s Found"
End Sub

Private Sub LoadJEVDetails(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim x As Integer
xNAcode = ""
    Drec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & Text1.Text & "' and ActionCode=1 and paout = 0"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        txtClaimant.Text = getClaimant(Drec!ClaimantCode)
        txtRC.Text = GetOfficeName(Drec!RCenter, "OfficeMedium")
        txtParticular.Text = Drec!Particular
        txtFund.Text = Drec!FundType
        txtAmount.Text = Format(Drec![Gamount], "#,##0.00")
        cmb_trans_class.ListIndex = ExcuteScalar("select isnull(claimclass,0) as field from dbo.tblAMIS_LogApprovedAndAudit where dvno = '" & Text1.Text & "'")
        
        If Drec!NonAlobs = 1 Then
            xObR = GetNonAlobsName(Drec!obrno)
            xNAcode = Drec!obrno
        Else
            xObR = Drec!obrno
        End If
        txtAlobs.Text = xObR
    Else
        'MsgBox "Transaction already Log Out and Approve...", vbInformation, "System Information"
    End If
    Drec.Close
    Set Drec = Nothing
End Sub
Private Sub Form_Load()
Call GetSignatory(cmbapprove, "Approved 2")
Call GetSignatory(cmbaudit, "Audit by")
Call GetSignatory(cmb_preAudit, "Audit by")
Call LoadDVNO
End Sub

Private Sub lst_DVNO_Click()
Text1.Text = lst_DVNO.Text
End Sub

Private Sub lvButtons_H1_Click()
On Error GoTo bad
    Call Savedata(2)
Exit Sub
bad:
MsgBox "Error: " & err.description, vbCritical, "System Message"

End Sub
Private Function Savedata(ByVal TYP As Integer)
Dim rec As New ADODB.Recordset
Dim tmpAppID As String
Dim xObR As String
Dim status As Long

status = getStatusID(Text1.Text)
If chkAudit.Value = 0 And Check1.Value = 0 Then
    MsgBox "Please Select check box above to procceed the transaction", vbInformation, "System Message"
    Exit Function
End If

If chkAudit.Value = 1 Then
    If status = 3 Then
        MsgBox "Oops..! Transaction already log out in Pre-Audit division, for Approval of the Accountant", vbCritical, "System Messagse"
        Exit Function
    ElseIf status = 1 Or status = 2 Then
        MsgBox "This transaction is for Audit-IN...", vbCritical, "System Messagse"
        Exit Function
    End If
End If

If Check1.Value = 1 Then
    If status = 26 Then
        MsgBox "Oops..! Done Assigning the Auditor..", vbCritical, "System Messagse"
        Exit Function
    End If
    
    If status = 3 Then
        MsgBox "Oops..! Transaction already log out in Pre-Audit division, for Approval of the Accountant", vbCritical, "System Messagse"
        Exit Function
    End If
End If

    Select Case DVApproved(Text1.Text)
        Case 0 'For Approval
                If Text1.Text = "" Then
                    MsgBox "Oops..! Please Specify the Dvno", vbCritical, "System Messagse"
                    Exit Function
                End If
                
                If chkAudit.Value = 1 And cmbaudit.Text = "" Then
                    MsgBox "Oops..! Please Specify Who Audited the Transaction", vbCritical, "System Messagse"
                    Exit Function
                End If
                
                If cmb_trans_class.Text = "" Then
                    MsgBox "Oops..! Please Specify claim classification", vbCritical, "System Messagse"
                    Exit Function
                End If
                
                 If Check1.Value = 1 And cmb_preAudit.Text = "" Then
                    MsgBox "Oops..! Assign first the Auditor...", vbCritical, "System Messagse"
                    Exit Function
                End If
                
            If MsgBox("Are you sure want you to save this Transaction", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then

                If Len(ActiveUserID) > 0 Then
                    If xNAcode = "" Then
                    xObR = txtAlobs.Text
                    Else
                    xObR = xNAcode
                    End If
                    If Check1.Value = 1 Then
                        Call AuditSig("preAudit", cmb_preAudit.ItemData(cmb_preAudit.ListIndex))
                    End If
                    
                    If chkAudit.Value = 1 Then
                        Call AuditSig("Auditby", cmbaudit.ItemData(cmbaudit.ListIndex))
                    End If
                    
                    opndbaseFMIS.Execute "Update [fmis].[dbo].[tblAMIS_LogApprovedAndAudit] set [claimClass]='" & cmb_trans_class.ListIndex & "' Where DVNo='" & Text1.Text & "'"
                    
                    MsgBox "Transaction Successfully Audited!", vbInformation, "Sytem Information"
                    Text1.Text = ""
                Else
                    Exit Function
                End If
            End If
        Case 1
            MsgBox "Transaction already Log Out and Approve...", vbInformation, "System Information"
        Case 4 'Not Yet Assigned
            MsgBox "Specified DV No. was not yet Registered!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
    End Select
End Function
Private Sub ApprovedSig()
Dim rec As New ADODB.Recordset
'
'set rec = opndbaseFMIS.Execute( "Update tblamis_journalentry set actioncode = 3,datetimeentered = datetimeentered + ',' + '" & Now & "',userid = userId + ',' + '" & ActiveUserID & "' where dvno = '" & Text1.Text & "' and actioncode = 1"
opndbaseFMIS.Execute "Update tblamis_journalentry set actioncode = 3,datetimeentered = datetimeentered + ',' + '" & Now & "',userid = userId + ',' + '" & ActiveUserID & "' where dvno = '" & Text1.Text & "' and actioncode = 1"
opndbaseFMIS.Execute "Insert Into fmis.dbo.tblAMIS_JournalEntry (DVNo,ObrNo,TransDate,UserID,Actioncode,DateTimeEntered,Continuing,debitcredit,isnew,FmisAccntCode,ApprovedByID,DateTimeApproved,logoutby,logoutdatetime,LogOutRemark) values " & _
                    "('" & Trim(Replace(Text1.Text, "'", "''")) & "','" & xObR & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & 0 & ",0,1,0,'" & ActiveUserID & "','" & Now & "','" & ActiveUserID & "','" & Now & "','" & txt_Remark.Text & "')"
Call LogApprovedAndAudit(Text1.Text, "Approvedby", cmbapprove.ItemData(cmbapprove.ListIndex))
End Sub
'Private Sub ApprovedSig()
'Dim rec As New ADODB.Recordset
'
'Set rec = opndbaseFMIS.Execute("Select dvno from tblamis_journalentry where dvno = '" & Text1.Text & "' and actioncode = 1")
'If rec.RecordCount > 0 Then
'opndbaseFMIS.Execute "Update tblamis_journalentry where dvno = '" & Text1.Text & "' and actioncode = 1"
'Else
'opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (DVNo,ObrNo,TransDate,UserID,Actioncode,DateTimeEntered,Continuing,debitcredit,isnew,FmisAccntCode,ApprovedByID,DateTimeApproved,logoutby,logoutdatetime,LogOutRemark) values " & _
'                    "('" & Trim(Replace(Text1.Text, "'", "''")) & "','" & xObR & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & 0 & ",0,1,0,'" & ActiveUserID & "','" & Now & "','" & ActiveUserID & "','" & Now & "','" & txt_Remark.Text & "')"
'End If
'Call LogApprovedAndAudit(Text1.Text, "Approvedby", cmbapprove.ItemData(cmbapprove.ListIndex))
'End Sub
Private Sub AuditSig(ByVal auditType As String, auditID As String)
opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns set PADesc='" & Replace(Trim(txt_Remark.Text), "'", "''") & "',ReturnFlag=" & chkReturn.Value & " Where DVNo='" & Text1.Text & "' And ActionCode=1"
 
 
 Call LogApprovedAndAudit(Text1.Text, auditType, auditID)
 
 
 If chkReturn.Value = 1 Then
    Call LogTrans(Text1.Text, 2) 'audit and return
 Else
    If auditType = "Auditby" Then
       Call LogTrans(Text1.Text, 3) 'final Audit
    End If
 
    If auditType = "preAudit" Then
    Call LogTrans(Text1.Text, 26) ' partial Audit
    End If
 End If
End Sub

Private Function DVApproved(ByVal dvno As String) As Integer
Dim opnDV As New ADODB.Recordset
Dim rec As New ADODB.Recordset

        rec.Open "Select TRNNO,returnFlag,PAout from tblAMIS_IncomingDVTrns where dvno = '" & Text1.Text & "' and actioncode = 1 ", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            If rec!returnflag = 0 Then
                 If rec!PAout = 1 Then
                    opnDV.Open "Select logoutby from [tblAMIS_JournalEntry] where dvno = '" & Text1.Text & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
                    If opnDV.RecordCount > 0 Then
                        If Len(Trim(opnDV!Logoutby)) = 0 Then
                            DVApproved = 0
                        Else
                            DVApproved = 1 'Already Log out in journal Entry
                        End If
                    Else
                    DVApproved = 0 ' for Log Out
                    End If
                    opnDV.Close
                 Else
                    DVApproved = 0 ' for Log Out
                 End If
                Else
                    DVApproved = 0 ' returned to the claimant
                End If
        Else
            DVApproved = 4 ' Not Register
        End If
        rec.Close
Set opnDV = Nothing
End Function
Private Sub lvButtons_H2_Click()
Unload Me
End Sub
Private Function getStatusID(ByVal dvno As String) As Long
Dim rec As New ADODB.Recordset
getStatusID = 0
Set rec = opndbaseFMIS.Execute("SELECT  [status] FROM [fmis].[dbo].[tblAMIS_Logtrans] where dvno = '" & Text1.Text & "' and actioncode = 1")
If rec.RecordCount > 0 Then
    getStatusID = rec!status
End If
rec.Close
End Function

Private Sub lvButtons_H3_Click()
Call LoadDVNO
End Sub

Private Sub text1_Change()
Dim status As Long
If Len(Trim(Text1.Text)) = 14 Then
status = getStatusID(Text1.Text)
    If status = 1 Or status = 2 Then
        Check1.Value = 1
        cmb_preAudit.Enabled = True
        cmbaudit.Enabled = False
        chkAudit.Value = 0
    ElseIf status = 26 Then
        chkAudit.Value = 1
        cmb_preAudit.Enabled = False
        cmbaudit.Enabled = True
        Check1.Value = 0
    ElseIf status = 3 Then
        Check1.Value = 0
        chkAudit.Value = 0
        MsgBox "This transaction is for Approval of the Accountant", vbCritical, "System Message"
    Else
        MsgBox "Status: " & ExcuteScalar("SELECT status as field FROM [fmis].[dbo].[tblCMS_TransStatusMap] where code = " & status & ""), vbInformation, "System Message"
    End If
    
    Call LoadJEVDetails(Text1.Text)
Else
    txtClaimant.Text = ""
    txtRC.Text = ""
    txtParticular.Text = ""
    txtFund.Text = ""
    txtAmount.Text = ""
    txtAlobs.Text = ""
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
