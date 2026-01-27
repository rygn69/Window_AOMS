VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_DVManagement 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disbursement Voucher Management"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   1455
   ClientWidth     =   12315
   ForeColor       =   &H00000000&
   Icon            =   "frm_DVManagement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   12315
   Begin VB.CheckBox chkAudit 
      BackColor       =   &H00000000&
      Caption         =   "Audit"
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
      Left            =   7560
      TabIndex        =   33
      Top             =   360
      Width           =   1095
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
      Left            =   9240
      TabIndex        =   32
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkReturn 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   11040
      TabIndex        =   27
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
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   3000
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
      Left            =   7620
      TabIndex        =   20
      Top             =   1095
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
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2040
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
      Height          =   795
      Left            =   7605
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3945
      Width           =   4500
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H8000000E&
      Height          =   5805
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   7155
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   6480
         TabIndex        =   34
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "..."
         CapAlign        =   2
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
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_DVManagement.frx":076A
         cBack           =   -2147483633
      End
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
         Width           =   4020
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
         ForeColor       =   &H8000000E&
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
         ForeColor       =   &H8000000E&
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
         ForeColor       =   &H8000000E&
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
         ForeColor       =   &H8000000E&
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
         ForeColor       =   &H8000000E&
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
         ForeColor       =   &H8000000E&
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
      Height          =   975
      Left            =   7560
      TabIndex        =   30
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      Caption         =   "&Update"
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
      cFHover         =   255
      LockHover       =   3
      cGradient       =   4210752
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frm_DVManagement.frx":4274
      ImgSize         =   24
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   975
      Left            =   11040
      TabIndex        =   31
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
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
      cFHover         =   255
      LockHover       =   3
      cGradient       =   4210752
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frm_DVManagement.frx":45C6
      ImgSize         =   24
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   975
      Left            =   8760
      TabIndex        =   35
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      Caption         =   "&Update"
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
      cFHover         =   255
      LockHover       =   3
      cGradient       =   4210752
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frm_DVManagement.frx":80D0
      ImgSize         =   24
      cBack           =   16777215
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
      Left            =   7560
      TabIndex        =   29
      Top             =   2640
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7605
      TabIndex        =   26
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7605
      TabIndex        =   25
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   7560
      TabIndex        =   24
      Top             =   3600
      Width           =   960
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
      Left            =   7605
      TabIndex        =   23
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label Label10 
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
      Left            =   7605
      TabIndex        =   22
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   7560
      TabIndex        =   21
      Top             =   3600
      Width           =   960
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
Attribute VB_Name = "frm_DVManagement"
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
Private Sub LoadStat()
    If chkAudit.Value = 1 Then
        chkApprove.Enabled = True
        cmbaudit.Enabled = True
    Else
        chkApprove.Enabled = False
        chkApprove.Value = 0
        cmbaudit.Enabled = False
        cmbaudit.ListIndex = 0
    End If

    If chkApprove.Value = 1 Then
        chkApprove.Value = 1
        cmbapprove.Enabled = True
    Else
        cmbapprove.ListIndex = 0
        cmbapprove.Enabled = False
    End If
    If chkReturn.Value = 1 Then
        chkApprove.Value = 0
        chkApprove.Enabled = False
'        cmbapprove.ListIndex = 0
        cmbapprove.Enabled = False
    Else

    End If
End Sub

Private Sub chkApprove_Click()
Call LoadStat
End Sub

Private Sub chkAudit_Click()
Call LoadStat
End Sub

Private Sub chkReturn_Click()
Call LoadStat
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub LoadJEVDetails(ByVal dvno As String)
On Error GoTo bad
Dim Drec As New ADODB.Recordset
Dim x As Integer
xNAcode = ""
    Drec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & Text1.Text & "' and ActionCode=1 "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        txtClaimant.Text = getClaimant(Drec!ClaimantCode)
        txtRC.Text = GetOfficeName(Drec!RCenter, "OfficeMedium")
        txtParticular.Text = Drec!Particular
        txtFund.Text = Drec!FundType
        txtAmount.Text = Format(Drec![Gamount], "#,##0.00")
        If Drec!NonAlobs = 1 Then
            xObR = GetNonAlobsName(Drec!obrno)
            xNAcode = Drec!obrno
        Else
            xObR = Drec!obrno
        End If
        txtAlobs.Text = xObR
        txt_Remark.Text = Drec!padesc
        Call getAuditSIG
    End If
    Drec.Close
    Set Drec = Nothing
Exit Sub
bad:
MsgBox err.description
Resume Next
End Sub
Public Sub getAuditSIG()
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("select fullname from tblReff_Signatory where cast(id as bigint) = (Select top 1 auditby from tblAMIS_LogApprovedAndAudit where dvno = '" & Text1.Text & "')")
'cmbaudit.ListIndex = 0
If rec.RecordCount > 0 Then
    cmbaudit.Text = rec!FullName
Else

End If
End Sub

Private Sub Form_Load()
Call GetSignatory(cmbapprove, "Approved 2")
Call GetSignatory(cmbaudit, "Audit by")
End Sub

Private Sub lvButtons_H1_Click()
On Error GoTo bad

If MsgBox("Please Select Either Approved or Returned before you Log Out the transaction...!", vbInformation, "System Message") = vbYes Then

End If
Exit Sub
bad:
MsgBox "Error: " & err.description, vbCritical, "System Message"
End Sub
Private Function Savedata(ByVal TYP As Integer)
Dim rec As New ADODB.Recordset
Dim tmpAppID As String
Dim xObR As String
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
                
                If chkAudit.Value = 1 And cmbaudit.Text <> "" And chkApprove.Value = 1 And cmbapprove.Text = "" Then
                    MsgBox "Oops..! Please Specify Who Approved the Transaction", vbCritical, "System Messagse"
                    Exit Function
                End If
            
                If txt_Remark.Text = "" Then
                    MsgBox "Remarks Empty, Please Specify the Remark..!", vbCritical, "System Messagse"
                    txt_Remark.SetFocus
                    Exit Function
                End If
                
            If MsgBox("Are you sure want to Log Out this JEV", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then

                If Len(ActiveUserID) > 0 Then
                    If xNAcode = "" Then
                    xObR = txtAlobs.Text
                    Else
                    xObR = xNAcode
                    End If
                    
                    If chkAudit.Value = 1 Then
                        Call AuditSig
                    End If
                    
                    If chkApprove.Value = 1 Then
                        Call ApprovedSig
                    End If
                    
                    MsgBox "Transaction LogOut!", vbInformation, "Sytem Information"
                    If chkReturn.Value = 1 Then
                    Call LogTrans(Text1.Text, 5) 'Approve and Log Out
                    Else
                    Call LogTrans(Text1.Text, 4)
                    End If
                    Text1.Text = ""
                    cmbapprove.ListIndex = 0
                    cmbaudit.ListIndex = 0
                Else
                    Exit Function
                End If
            End If
        Case 1 'Approved
            MsgBox "Specified DV No. was Already Approved and Log Out!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
         Case 3 'returned to the claimant
            MsgBox "Specified DV No. was Returned to the Claimant!" & Chr(13) & Chr(13) & "Please In first to the Pre-Audit to Proceed the Transaction", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
        Case 4 'Not Yet Assigned
            MsgBox "Specified DV No. was not yet Registered!" & Chr(13) & Chr(13) & "Please Enter a New DVNo.", vbInformation, "System Information"
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Text1.SetFocus
    End Select
End Function
Private Sub ApprovedSig()
opndbaseFMIS.Execute "Update tblamis_journalentry set actioncode = 3,datetimeentered = datetimeentered + ',' + '" & Now & "',userid = userId + ',' + '" & ActiveUserID & "' where dvno = '" & Text1.Text & "' and actioncode = 1"
opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (DVNo,ObrNo,TransDate,UserID,Actioncode,DateTimeEntered,Continuing,debitcredit,isnew,FmisAccntCode,ApprovedByID,DateTimeApproved,logoutby,logoutdatetime,LogOutRemark) values " & _
                    "('" & Trim(Replace(Text1.Text, "'", "''")) & "','" & xNAcode & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & 0 & ",0,1,0,'" & ActiveUserID & "','" & Now & "','" & ActiveUserID & "','" & Now & "','" & txt_Remark.Text & "')"
Call LogApprovedAndAudit(Text1.Text, "Approvedby", cmbapprove.ItemData(cmbapprove.ListIndex))
End Sub
Private Sub AuditSig()
 opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns Set PAout=1, PAoutDate='" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "', PADesc='" & Replace(Trim(txt_Remark.Text), "'", "''") & "', OutBy='" & ActiveUserID & "',ReturnFlag=" & chkReturn.Value & " Where DVNo='" & Text1.Text & "' And ActionCode=1"
 Call LogApprovedAndAudit(Text1.Text, "Auditby", cmbaudit.ItemData(cmbaudit.ListIndex))
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
                    DVApproved = 3
3                     ' returned to the claimant
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

Private Sub text1_Change()
If Len(Trim(Text1.Text)) = 14 Then
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
