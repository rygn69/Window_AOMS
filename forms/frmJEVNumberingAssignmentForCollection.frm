VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJEVNumberingAssignmentForCollection_new 
   Caption         =   "JEV Numbering Assignment for Collection and General Journal"
   ClientHeight    =   8925
   ClientLeft      =   450
   ClientTop       =   1155
   ClientWidth     =   12240
   Icon            =   "frmJEVNumberingAssignmentForCollection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   12240
   Begin VB.TextBox txtDate 
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
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   345
      Width           =   2565
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   29
      ToolTipText     =   "Saves to Journal Entry"
      Top             =   240
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingAssignmentForCollection.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingAssignmentForCollection.frx":0BBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign No. to JEV"
      Height          =   720
      Left            =   11040
      TabIndex        =   24
      Top             =   3600
      Width           =   990
   End
   Begin VB.TextBox txt_JEVNo 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7155
      MaxLength       =   18
      TabIndex        =   21
      Top             =   3840
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   210
      ScaleHeight     =   4200
      ScaleWidth      =   11910
      TabIndex        =   18
      Top             =   4560
      Width           =   11940
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cmbEntry 
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Text            =   "cmbEntry"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   7408
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   225
      TabIndex        =   13
      Top             =   3495
      Width           =   2550
      Begin VB.OptionButton opn_Coll 
         Caption         =   "Collection"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Tag             =   "01"
         Top             =   405
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton opn_CheckDisb 
         Caption         =   "Check Disbursement"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   16
         Tag             =   "02"
         Top             =   1020
         Width           =   2100
      End
      Begin VB.OptionButton opn_CashDisb 
         Caption         =   "Cash Disbursement"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4245
         TabIndex        =   15
         Tag             =   "03"
         Top             =   1020
         Width           =   2100
      End
      Begin VB.OptionButton opn_Other 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6405
         TabIndex        =   14
         Tag             =   "04"
         Top             =   1020
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   225
      TabIndex        =   0
      Top             =   1200
      Width           =   11835
      Begin VB.ComboBox cmbrc 
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
         Left            =   5160
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2580
         Visible         =   0   'False
         Width           =   4335
      End
      Begin MSComctlLib.ImageCombo cmbfundtype 
         Height          =   330
         Left            =   11280
         TabIndex        =   27
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
      Begin VB.TextBox txt_RCenter 
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
         Left            =   5160
         TabIndex        =   6
         Top             =   2460
         Width           =   4290
      End
      Begin VB.TextBox txt_Claimant 
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
         Left            =   315
         TabIndex        =   5
         Top             =   2385
         Width           =   4260
      End
      Begin VB.TextBox txt_AlobsNo 
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
         Left            =   1995
         TabIndex        =   4
         Top             =   420
         Width           =   4260
      End
      Begin VB.TextBox txt_particular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   1995
         TabIndex        =   3
         Top             =   840
         Width           =   4290
      End
      Begin VB.TextBox txt_Amount 
         Alignment       =   2  'Center
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
         Left            =   7680
         TabIndex        =   2
         Top             =   840
         Width           =   3540
      End
      Begin VB.TextBox txt_FundType 
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
         Left            =   7680
         TabIndex        =   1
         Top             =   420
         Width           =   3540
      End
      Begin MSComctlLib.ImageCombo cmbOffice 
         Height          =   330
         Left            =   6960
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
         Height          =   195
         Left            =   5100
         TabIndex        =   12
         Top             =   2310
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2130
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Number:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   810
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   8
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6420
         TabIndex        =   7
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_DVNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   22
      Top             =   480
      Width           =   3885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update Accountcode"
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5760
      TabIndex        =   32
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PTV Number:"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   180
      TabIndex        =   23
      Top             =   75
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned JEV No."
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7185
      TabIndex        =   20
      Top             =   3495
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   990
      Left            =   7005
      Top             =   3450
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1080
      Left            =   -60
      Top             =   -15
      Width           =   5820
   End
End
Attribute VB_Name = "frmJEVNumberingAssignmentForCollection_new"
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
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
Public ptv As String

Private Sub LoadBackDVDetails(ByVal dvno As String)
Dim opnDV As New ADODB.Recordset

End Sub
Public Sub LoadAccountsByFund(ByVal fundmedium As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer

    cmbEntry.Clear
    cmbEntry.Visible = False
    ARec.Open ("Select distinct * from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & txt_FundType.Text & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For x = 1 To ARec.RecordCount
            cmbEntry.AddItem ARec![childAccountcode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
            ARec.MoveNext
        Next x
    End If
    ARec.Close
    Set ARec = Nothing
    
End Sub

Private Function coloraly() As Boolean
Dim x As Integer
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(x, 6) <> "" Then
                If MSFlexGrid1.TextMatrix(x, 6) = "5" Then
                    coloraly = True
                    Exit Function
                End If
            End If
        Else
            Exit For
        End If
    Next x
End Function

Private Sub SetGrid()
Dim cc As Integer

MSFlexGrid1.Clear
MSFlexGrid1.Cols = 7
MSFlexGrid1.Rows = 2

    MSFlexGrid1.TextMatrix(0, 0) = "trnno"
    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
    MSFlexGrid1.TextMatrix(0, 6) = "Actioncode"

MSFlexGrid1.ColWidth(0) = 0
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 4000
MSFlexGrid1.ColWidth(3) = 5000
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(5) = 2000
MSFlexGrid1.ColWidth(6) = 600

For cc = 0 To MSFlexGrid1.Cols - 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.col = cc
    MSFlexGrid1.CellAlignment = 4
Next cc
End Sub
Private Sub LoadOtherDetails(ByVal DV As String)
Dim opnDV As New ADODB.Recordset

opnDV.Open "Select * from tblAMIS_COllectionDepositt where ptvno='" & DV & "' and (actioncode=1 or actioncode=5) ", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDV.RecordCount <> 0 Then

    txt_AlobsNo.Text = opnDV.Fields!ReportNo
    'txt_fundtype.Text = getfundmedium(opndv.Fields!
    
    Call SelectTrnType(opnDV!Transtype)
    
    
    Do Until opnDV.EOF
    
    MSFlexGrid1.TextMatrix(0, 0) = "trnno"
    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
    MSFlexGrid1.TextMatrix(0, 6) = "actioncode"
    
    
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = opnDV!Trnno
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = opnDV!FmisAccntCode
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = GetAccntDescription(opnDV!FmisAccntCode, "ACCT_CODE")
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = GetAccntDescription(opnDV!FmisAccntCode, "ACCT_ENTRIES")
 
    If opnDV!DebitCredit = 1 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = opnDV!Amount
    ElseIf opnDV!DebitCredit = 0 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = opnDV!Amount
    End If
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = opnDV!ActionCode
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    opnDV.MoveNext
    Loop
    Call GetSum
Else
    Call ClearAllOption
End If
opnDV.Close
Set opnDV = Nothing

End Sub

Private Function GetAccntDescription(ByVal FMISCode As Long, ByVal NeedFld As String) As String
Dim opnDesc As New ADODB.Recordset

opnDesc.Open "Select * from tblREF_AIS_ChartofAccounts where FMISAccountCode=" & FMISCode & " and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDesc.RecordCount <> 0 Then
    Select Case NeedFld
        Case "ACCT_ENTRIES"
            If opnDesc!Accountname = opnDesc!AccountNamefull Then
                GetAccntDescription = opnDesc!Accountname
            Else
                GetAccntDescription = opnDesc!Accountname & "-" & opnDesc!AccountNamefull
            End If
        
        Case "ACCT_CODE"
            GetAccntDescription = opnDesc!childAccountcode
    End Select
End If
opnDesc.Close

End Function
Private Sub ClearAllOption()
opn_Coll.Value = False
opn_CheckDisb.Value = False
opn_CashDisb.Value = False
opn_Other.Value = False
End Sub
Private Sub SelectTrnType(ByVal TransCode As String)
Select Case TransCode
    Case 1
        opn_Coll.Value = True
    Case 2
        opn_CheckDisb.Value = True
    Case 3
        opn_CashDisb.Value = True
    Case 4
        opn_Other.Value = True
End Select
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            ElseIf MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            ElseIf Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)) > 0 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) > 0 Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = GetFMISAccountCodeUSingchildaccountcode(cmbEntry.Text, txt_FundType.Text)
        Else
             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = ""
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
        End If
        cmbEntry.Visible = False
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        If cmbEntry.Text = "101" Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "5"
        Else
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "1"
        End If
        
        Call GetSum
        MSFlexGrid1.SetFocus
    Else
        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
    End If
Edited = True
End Sub
Private Sub GetSum1()
On Error GoTo bad
Dim x As Integer

    xDebit = 0
    xCredit = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
        Else
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = xDebit
            MSFlexGrid1.TextMatrix(x, 4) = xCredit
            Exit For
        End If
    Next x
Exit Sub
bad:
MsgBox err.Description
End Sub

Private Sub cmbRC_Click()

    If Trim(cmbrc.Text) <> "" Then
        txt_RCenter.Text = Trim(cmbrc.Text)
        txt_RCenter.Visible = True
        cmbrc.Visible = False
    End If
End Sub

Private Sub cmbrc_LostFocus()
cmbrc.Visible = False
txt_RCenter.Visible = True
End Sub

Private Sub Command1_Click()
If CheckIfExistInFinalJEV(txt_JEVNo.Text) = False Then
Select Case ActiveFormCaller
    Case "frmCDCashReceiptsJevNumbering"
        frmCDCashReceiptsJevNumbering.lstDetails.SelectedItem.SubItems(10) = txt_JEVNo.Text
        Unload Me
End Select
Else
    MsgBox "JEV number Already exist on the database..", vbInformation, "System Message"
End If
End Sub
Private Function ChkEntry() As Boolean

    ChkEntry = False
    If Trim(txt_DVNo.Text) <> "" And txt_AlobsNo.Text <> "" And txt_particular.Text <> "" And txt_FundType.Text <> "" And txt_Amount.Text <> "" Then
        
        
        If xDebit = xCredit And xDebit > 0 Then
        If coloraly = True Then GoTo coloraly_jmp 'coloraly consideration - set chkentry to true even if not balance
            If Format(xDebit, "###,##0.00") = Format(txt_Amount.Text, "###,##0.00") Then
coloraly_jmp:
                ChkEntry = True
            End If
        End If
        

'        If xDebit = xCredit And xDebit > 0 Then
'            If Format(xDebit, "###,##0.00") = Format(txt_Amount.Text, "###,##0.00") Then
'                ChkEntry = True
'            End If
'        End If
    End If
    
End Function


Private Sub cmdSave_Click()
Dim DRec As New ADODB.Recordset
Dim xType As Integer
Dim x As Integer
      If ChkEntry = True Then
        If MsgBox("Are you sure you want to update this transaction?", vbQuestion + vbYesNo) = vbYes Then
            If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
            If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
            If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
            If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
            
            'If Edited = True Then
                opndbaseFMIS.Execute "Update tblAMIS_COllectionDepositt set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where ptVNo='" & Me.txt_DVNo.Text & "' And ActionCode=1"
                opndbaseFMIS.Execute "Update tblAMIS_COllectionDepositt set ActionCode=6, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where ptVNo='" & Me.txt_DVNo.Text & "' And ActionCode=5"
            'End If
            
            
            If xNAcode <> "" Then
                xObR = xNAcode
            End If
            
            For x = 1 To MSFlexGrid1.Rows - 1
                If MSFlexGrid1.TextMatrix(x, 3) <> "TOTAL" Then
                    If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
                        If MSFlexGrid1.TextMatrix(x, 4) <> "" Or MSFlexGrid1.TextMatrix(x, 5) <> "" Then
                            opndbaseFMIS.Execute "Insert Into tblAMIS_COllectionDepositt (TransType,PTVno,reportno,FmisAccntCode,Amount,DebitCredit,TransDate,UserID,Actioncode,DateTimeEntered) values (" & xType & ",'" & Trim(Replace(txt_DVNo.Text, "'", "''")) & "','" & txt_AlobsNo.Text & "'," & CLng(MSFlexGrid1.TextMatrix(x, 1)) & "," & CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 4)), MSFlexGrid1.TextMatrix(x, 4), 0)) + CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 5)), MSFlexGrid1.TextMatrix(x, 5), 0)) & "," & IIf(Trim(MSFlexGrid1.TextMatrix(x, 4)) = "", 0, 1) & ",'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "','" & (MSFlexGrid1.TextMatrix(x, 6)) & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "')"
                        End If
                    End If
                Else
                    Exit For
                End If
            Next x
            'Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
        End If
    Else
        MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
    End If
End Sub
Private Sub LoadClaimantCODE(ByVal Claimant As Variant)
'Dim opnDetails As New ADODB.Recordsetn
'Dim opnDetails1 As New ADODB.Recordset
'
'Select Case Classification
'Case "Individual", "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
'opnDetails.Open "Select * from tblCMS_CDClaimantDetails where lastname = '" & Claimant & "'%", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnDetails.RecordCount <> 0 Then
'    LoadClaimantCODE = opnDetails!ClaimantCode
'Else
'    opnDetails1.Open "Select * from employee where firstname like '" & Claimant & "%'", opndbasePMIS, adOpenStatic, adLockOptimistic
'    If opnDetails.RecordCount <> 0 Then
'    LoadClaimantCODE = opnDetails1!SwipEmployeeID
'    End If
'End If
'opnDetails.Close
'Set opnDetails = Nothing
MsgBox "ERROR"
End Sub

'Private Sub Command2_Click()
'Dim xType As Integer
'        If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
'
'            If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
'            If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
'            If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
'            If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
'
'            For x = 1 To MSFlexGrid1.Rows - 1
'                If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
'                    If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
'
'                            opndbaseFMIS.Execute "update tblAMIS_JournalEntry set FmisAccntCode = '" & MSFlexGrid1.TextMatrix(x, 1) & "',transtype = " & xType & " where trnno = '" & MSFlexGrid1.TextMatrix(x, 0) & "'"
'
'                    End If
'                Else
'                    Exit For
'                End If
'            Next x
'        End If
'        MsgBox "Successfully Update", vbInformation, "System Message"
'    End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Call SetGrid
LoadOffice
End Sub


Private Sub GetSum()
Dim x As Integer
    not_coloraly_total_debit = 0
    not_coloraly_total_credit = 0
     coloraly_total_credit = 0
     coloraly_total_debit = 0
      
    xDebit = 0
    xCredit = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
                If Trim(MSFlexGrid1.TextMatrix(x, 6)) <> 5 Then
                    not_coloraly_total_debit = not_coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
                    not_coloraly_total_credit = not_coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
                Else
                    coloraly_total_debit = coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
                    coloraly_total_credit = coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
                End If
        Else
            MSFlexGrid1.TextMatrix(x, 3) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 4) = xDebit
            MSFlexGrid1.TextMatrix(x, 5) = xCredit
            Exit For
        End If
    Next x
    
End Sub
Public Sub LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmbrc.Clear

OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmbrc.AddItem OREc![OfficeMedium]
        cmbrc.ItemData(cmbrc.NewIndex) = OREc!FMISOfficeID
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing

End Sub



Private Sub MSFlexGrid1_Click()
On Error GoTo bad
    Select Case MSFlexGrid1.col
    Case 2 'AccntCode
        txt_entry.Visible = False
        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
        cmbEntry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            cmbEntry.Text = MSFlexGrid1.Text
        Else
            cmbEntry.ListIndex = -1
        End If
        cmbEntry.SetFocus
    Case 4 To 6 'Debit/Credit/actioncode
        cmbEntry.Visible = False
        txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
        txt_entry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            txt_entry.Text = MSFlexGrid1.Text
            txt_entry.SelStart = 0
            txt_entry.SelLength = Len(txt_entry.Text)
        Else
            txt_entry.Text = ""
        End If
        txt_entry.SetFocus
    Case Else
        txt_entry.Visible = False
        cmbEntry.Visible = False
    End Select
Exit Sub
bad:
    MsgBox err.Description
End Sub


Private Sub MSFlexGrid1_DblClick()
If Trim(txtAlobs.Text) <> "" Then
    With frmSub3
        .reff = txt_DVNo.Text
        .Gamount = txt_Amount.Text
        .CName = "N/A"
        .Show 1
        Call LoadAcctngEntries(Trim(txt_DVNo.Text))
    End With
End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Call MSFlexGrid1_Click
End Sub

Private Sub txt_DVNo_Change()
If Len(Trim(txt_JEVNo.Text)) = 0 Then
    txt_JEVNo.Text = SetNewJEVNo(txt_DVNo.Text, frmJEVNumberingThruRCI.DTPicker1.Year, frmJEVNumberingThruRCI.DTPicker1.Month)
End If
Call SetGrid
Call LoadOtherDetails(txt_DVNo.Text)
Call LoadAccountsByFund(Trim(Me.txt_FundType))
End Sub


Private Sub txt_entry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = txt_entry.Text
        If MSFlexGrid1.col = 4 Then
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
        Else
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
            End If
        End If
        txt_entry.Visible = False
        
        Call GetSum
        txt_entry.Text = ""
        MSFlexGrid1.SetFocus
    End If

End Sub

Private Sub txt_RCenter_Click()
cmbrc.Visible = True
End Sub
