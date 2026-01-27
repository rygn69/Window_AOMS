VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmJEVPreparationforAjustment_new 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjusting Entry"
   ClientHeight    =   11775
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVPreparationforAjustment_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11775
   ScaleWidth      =   11535
   Visible         =   0   'False
   Begin lvButton.lvButtons_H btn_generate 
      Height          =   495
      Left            =   9720
      TabIndex        =   24
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Generate JEV No."
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtJEV2 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8760
      TabIndex        =   23
      Top             =   6120
      Width           =   825
   End
   Begin VB.TextBox txtJEV1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   6120
      Width           =   2025
   End
   Begin VB.CheckBox chkSC 
      Caption         =   "Single Click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6000
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   180
      TabIndex        =   10
      Top             =   1920
      Width           =   11115
      Begin VB.CheckBox Check1 
         Caption         =   "Continuing"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox txtFund 
         Appearance      =   0  'Flat
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   4935
      End
      Begin VB.ComboBox cmbrc 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   14
         Top             =   3360
         Width           =   1785
      End
      Begin VB.TextBox txtParticular 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2160
         Width           =   8970
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1635
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   165
      ScaleHeight     =   4665
      ScaleWidth      =   11160
      TabIndex        =   2
      Top             =   6840
      Width           =   11190
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4680
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8255
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   165
      TabIndex        =   1
      Top             =   1335
      Width           =   4845
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1508
      ButtonWidth     =   1032
      ButtonHeight    =   1455
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   7560
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":076A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":20FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":3A8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":5420
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":6DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":8744
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":A0D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":BA68
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":D3FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":ED8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":FA6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":1034A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":11026
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":11D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":129DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":136BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforAjustment_New.frx":14396
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   600
      TabIndex        =   5
      Top             =   6960
      Width           =   7830
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6405
         TabIndex        =   9
         Tag             =   "04"
         Top             =   300
         Width           =   1230
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4245
         TabIndex        =   8
         Tag             =   "03"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCheck 
         Caption         =   "Check Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   7
         Tag             =   "02"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCollection 
         Caption         =   "Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Tag             =   "01"
         Top             =   285
         Value           =   -1  'True
         Width           =   1260
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JEV No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      TabIndex        =   25
      Top             =   6240
      Width           =   1125
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference No.(JEVNO):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   3
      Top             =   6480
      Width           =   2010
   End
End
Attribute VB_Name = "frmJEVPreparationforAjustment_new"
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
Dim CUFlag As Boolean           'Claimant Update Flag
Dim XFlag As Boolean
Public isfrom_jevNumbering As Boolean
Public EditCount, IsSaveAccntng As Boolean
Public ClaimantCode As String, RCenter As String

Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double


Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub


Private Sub btn_generate_Click()
If txtFund.Text = "" Then
    MsgBox "Please specify fund type..", vbInformation, "System Message"
    Exit Sub
End If
Dim rec As New ADODB.Recordset
Dim Lastno As Double
JevOk = False
frmPOstdate.Show 1
If JevOk = True Then
    rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = 4,@jevyeardate = '" & DatePost & "' ,@fundtype = '" & txtFund.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        Lastno = rec.Fields!MAXJEVSERIES
    rec.Close
    
    txtJEV1.Text = txtFund.ItemData(txtFund.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & "04" & "-"
    txtJEV2.Text = Format(Lastno, "0000")
End If
End Sub


Private Sub Form_Load()
    txtJEV2.MaxLength = 4
    Call SetGrid
    ActiveUserID = Trim(ActiveUserID)
    Call LoadFundType(txtFund)
    Call LoadOffice
End Sub
Private Sub SetGrid()
Dim cc As Integer
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 50
    MSFlexGrid1.Cols = 7
    
    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    MSFlexGrid1.TextMatrix(0, 6) = "Formula"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 2500
    MSFlexGrid1.ColWidth(2) = 5000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColWidth(6) = 0
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Function jevno() As String
jevno = txtJEV1.Text & txtJEV2.Text
End Function
Private Sub lvButtons_H1_Click()

End Sub

Private Sub MSFlexGrid1_Click()
If jevno = "" Then
    MsgBox "Please Enter JEVNO to Proceed the Entry..", vbInformation, "System Message"
    Exit Sub
End If
If chkSC.Value = 1 Then
    Call MSFlexGrid1_DblClick
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error GoTo bad
If jevno = "" Then
    MsgBox "Please Enter JEVNO to Proceed the Entry..", vbInformation, "System Message"
    Exit Sub
End If
If IsNumeric(txtAmount.Text) = False Then
    MsgBox "Please enter valid gross amount", vbInformation, "System Message"
    txtAmount.SetFocus
    Exit Sub
End If
If Trim(txtparticular.Text) <> "" Then
    With frmSub3
        .isPOSTED = False
        .REFF = jevno
        .Gamount = IIf(IsNumeric(txtAmount.Text), txtAmount.Text, 0)
        .CName = "N/A"
        .isEdit = True
       Set .frm = Me
        Call LoadAcctngEntries(jevno)
        .Show 1
        Call GetAccntngEntries
    End With
End If
Exit Sub
bad:
MsgBox err.description
End Sub
Public Sub GetAccntngEntries()
Dim Drec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set Drec = opndbaseFMIS.Execute("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & jevno & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)")
        If Drec.RecordCount > 0 Then
            For x = 1 To Drec.RecordCount
    '            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
                MSFlexGrid1.TextMatrix(x, 1) = Drec!childcode
                MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
                MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
                MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
              MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
               ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
                Drec.MoveNext
            Next x
        Else
        IsSaveAccntng = True
        Call GetAccntngEntries
        End If
    Else
        Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & jevno & "' group by Dvno,left(Accountcode,3)")
    If Drec.RecordCount > 0 Then
        For x = 1 To Drec.RecordCount
            'MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid1.TextMatrix(x, 1) = Drec!childcode
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            Drec.MoveNext
        Next x
    End If
    End If
    Call GetSum
    Drec.Close
    Set Drec = Nothing
End Sub

Public Function LoadAcctngEntries(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer

    Drec.Open ("Select ChildAccountcode,Debit ,Credit From tblAMIS_AccoutingEntries Where [reffno]='" & dvno & "' And (ActionCode=1) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        If EditCount = False Then
        EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & dvno & "'"
                        For x = 1 To Drec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(Drec!childaccountcode) & "'," & Drec!Debit & "," & Drec!Credit & ")"
                            Drec.MoveNext
                        Next x
                    End If
            Else
            For x = 1 To Drec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(Drec!childaccountcode) & "'," & Drec!Debit & "," & Drec!Credit & ")"
                            Drec.MoveNext
                        Next x
            End If
            rec.Close
        End If
    Else
         If EditCount = False Then
            EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & dvno & "'"
                        For x = 1 To Drec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(Drec!childaccountcode) & "'," & Drec!Debit & "," & Drec!Credit & ")"
                            Drec.MoveNext
                        Next x
                    End If
            End If
            rec.Close
        End If
    End If
    Drec.Close
    Set Drec = Nothing
End Function

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Integer
Dim xType As Integer, coloraly_signal As Integer

    Select Case Button:
    Case "New":
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
                
                txtDVNo.Text = ""
                

                
                txtparticular.Text = ""
                'txtFund.Text = ""
                txtAmount.Text = ""
                
                
                optCollection.Value = True
                
                Call SetGrid
    Case "Save":
                If ChkEntry = True Then
                
                    If cmbRC.Text = "" Then
                        MsgBox "Please select responsibility center", vbInformation, "System Message"
                        Exit Sub
                    End If
                
                
                    If CheckIfExistInFinalJEV(jevno) = True Then
                        If MsgBox("JEV number already exist in the Database, Do you want System Generated JEV Number?", vbInformation + vbYesNo, "System Message") = vbYes Then
                            Call btn_generate_Click
                        End If
                    End If
                        If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                        
                            Call SaveAcctngEntries(jevno)
                            Call GEtCompleteJEVDetails_v1(jevno, "Reffno", DatePost, "", "" _
                                , Replace(txtparticular.Text, "'", "''"), jevno, "", "", txtAmount.Text, "0", "0", 4, "", "", "", txtFund.Text, cmbRC.ItemData(cmbRC.ListIndex), "", "", txtDVNo.Text, ExtractJEVSNo(jevno), DatePost, "", Check1.Value, 1)
                            MsgBox "Successfully Save", vbInformation, "System Message"
                        End If
                    Else
debit_credit_error:
                        MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
                
                End If
    Case "Delete":
               
    Case "Close":
            Unload Me
    End Select
    
End Sub
Public Function SaveAcctngEntries(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim xType As Integer
If optCollection.Value = True Then xType = CInt(optCollection.Tag)
If optOther.Value = True Then xType = CInt(optOther.Tag)
    Drec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & dvno & "'group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_AccoutingEntries set actioncode =2 where reffno = '" & dvno & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = UserID + '," & Trim(ActiveUserID) & "'
        For x = 1 To Drec.RecordCount
            opndbaseFMIS.Execute "Insert into tblAMIS_AccoutingEntries (reffNo,ChildAccountcode,debit,credit,actioncode,datetimeentered,transtype,userid) values " & _
            "('" & Trim(dvno) & "','" & Trim(Drec!accountcode) & "'," & Drec!Debit & "," & Drec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & xType & ",'" & Trim(ActiveUserID) & "')"
            Drec.MoveNext
            DoEvents
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & dvno & "'"
    End If
    Drec.Close
    Set Drec = Nothing
End Function
Private Function coloraly() As Boolean
Dim x As Integer
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(x, 5) <> "" Then
                If MSFlexGrid1.TextMatrix(x, 5) = "5" Then
                    coloraly = True
                    Exit Function
                End If
            End If
        Else
            Exit For
        End If
    Next x
End Function


Private Function ChkEntry() As Boolean

    ChkEntry = False
    If Trim(txtDVNo.Text) <> "" And txtparticular.Text <> "" And txtFund.Text <> "" And txtAmount.Text <> "" And jevno <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
        If coloraly = True Then GoTo coloraly_jmp 'coloraly consideration - set chkentry to true even if not balance
            If Format(xDebit, "###,##0.00") = Format(txtAmount.Text, "###,##0.00") Then
coloraly_jmp:
                ChkEntry = True
            End If
        End If
    End If
    
End Function
'
'Private Sub LoadExcessDetails(ByVal ObR As String)
'Dim OREc As New ADODB.Recordset
'Dim x As Integer
'Dim y As Integer
'
'    Call SetGrid
'    OREc.Open ("Select * from [tblBMS_ExcessControl] where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If OREc.RecordCount > 0 Then
'        For x = 1 To OREc.RecordCount
'            For y = 0 To cmbEntry.ListCount - 1
'                If cmbEntry.List(y) = "401" Then
'                    cmbEntry.ListIndex = y
'                    Exit For
'                Else
'                    If y = cmbEntry.ListCount - 1 Then
'                        cmbEntry.ListIndex = -1
'                    End If
'                End If
'            Next y
'            MSFlexGrid1.TextMatrix(x, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
'            MSFlexGrid1.TextMatrix(x, 1) = "401"
'            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
'            MSFlexGrid1.TextMatrix(x, 4) = OREc!amount
'            OREc.MoveNext
'        Next x
'        Call GetSum
'    End If
'    OREc.Close
'    Set OREc = Nothing
'
'End Sub


Private Sub LoadObRDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim x As Integer
    
    Call SetGrid
    OREc.Open ("Select * from tblBMS_SubsidiaryLedger where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For x = 1 To OREc.RecordCount
            MSFlexGrid1.TextMatrix(x, 0) = OREc!FmisAccountcode
            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(OREc!FmisAccountcode)
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(OREc!FmisAccountcode)
            MSFlexGrid1.TextMatrix(x, 4) = OREc!amount
            OREc.MoveNext
        Next x
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub
'
'Public Sub LoadAccountsByFund(ByVal fundmedium As String)
'Dim ARec As New ADODB.Recordset
'Dim x As Integer
'Dim FundName As String
'
'    cmbEntry.Clear
'    cmbEntry.Visible = False
'    If Left(fundmedium, 3) = "Eco" Then
'    FundName = "Economic Enterprises"
'    Else
'    FundName = fundmedium
'    End If
'    ARec.Open ("Select distinct * from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & FundName & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If ARec.RecordCount > 0 Then
'        For x = 1 To ARec.RecordCount
'            cmbEntry.AddItem ARec![childaccountcode]
'            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
'            ARec.MoveNext
'        Next x
'    End If
'    ARec.Close
'    Set ARec = Nothing
'
'End Sub


Private Function sumAmount(ByVal amnt As String) As String
On Error GoTo sum
Dim x As Integer
Dim y As String
Dim str() As String
    If Left(amnt, 1) = "+" Then
    Else
    amnt = "+" & amnt
    End If
 
 str = Split(Trim(amnt), "+", -1, vbTextCompare)
 y = 0

 For x = 1 To 1000
y = val(y) + val(str(x))
 Next x
 Exit Function
sum:
 If err.Number = 9 Then
 sumAmount = y
Else
MsgBox "Incorrect Format", vbInformation, "System Message"
End If
End Function

Private Sub GetSum()
On Error GoTo bad
Dim x As Integer
    not_coloraly_total_debit = 0
    not_coloraly_total_credit = 0
     coloraly_total_credit = 0
     coloraly_total_debit = 0
      
    xDebit = 0
    xCredit = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
        Else
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = Format(xDebit, "#,##0.00")
            MSFlexGrid1.TextMatrix(x, 4) = Format(xCredit, "#,##0.00")
            Exit For
        End If
    Next x
Exit Sub
bad:
MsgBox err.description
End Sub

Private Function ChkIfAlreadyJEV(ByVal dvno As String) As String
Dim Jrec As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    Jrec.Open ("Select * from tblAMIS_COllectionDepositt where PTVNO='" & dvno & "' and (Actioncode=1 or Actioncode=5) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
'        If Not IsNull(JREc!ApprovedByID) Then
'            ChkIfAlreadyJEV = "Approved" & "-" & JREc!JEVNo
'        Else
            ChkIfAlreadyJEV = dvno
       ' End If
    End If
    Jrec.Close
    Set Jrec = Nothing
    
End Function


Public Sub LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer
cmbRC.Clear
OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
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


Private Sub txtAmount_LostFocus()
Call Format_Number(txtAmount)
End Sub

Private Sub txtFund_Change()
txtJEV1.Text = ""
txtJEV2.Text = ""
End Sub

Private Sub txtFund_Click()
txtJEV1.Text = ""
txtJEV2.Text = ""
End Sub

Private Sub txtJEV2_LostFocus()
If txtJEV2.Text = "" Then
    Exit Sub
End If
If IsNumeric(txtJEV2.Text) = True Then
   txtJEV2.Text = Format(txtJEV2.Text, "0000")
Else
    MsgBox "Invalid JEV No. Format..", vbCritical + vbInformation, "System Message"
    txtJEV2.SetFocus
End If
End Sub
