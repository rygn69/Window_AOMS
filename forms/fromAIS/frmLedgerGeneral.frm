VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frm_LedgerGeneral_option 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5370
   ClientLeft      =   4365
   ClientTop       =   2610
   ClientWidth     =   4875
   Icon            =   "frmLedgerGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   4875
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Caption         =   "From"
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
      Height          =   900
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   2340
      Begin MSComCtl2.DTPicker DTPicker2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   195
         TabIndex        =   15
         Top             =   360
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   57344001
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "frmLedgerGeneral.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Special Accounts:"
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
      Height          =   900
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4620
      Begin VB.ComboBox cmb_FundType 
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
         ItemData        =   "frmLedgerGeneral.frx":0776
         Left            =   195
         List            =   "frmLedgerGeneral.frx":0783
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   4260
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "To"
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
      Height          =   900
      Left            =   2520
      TabIndex        =   9
      Top             =   3720
      Width           =   2220
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   57344001
         UpDown          =   -1  'True
         CurrentDate     =   40431
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Account Code"
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
      Height          =   900
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   4620
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
         ItemData        =   "frmLedgerGeneral.frx":07B9
         Left            =   240
         List            =   "frmLedgerGeneral.frx":07BB
         TabIndex        =   8
         Top             =   300
         Width           =   4260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   35
      Left            =   -135
      TabIndex        =   4
      Top             =   840
      Width           =   7335
   End
   Begin VB.CommandButton LaVolpeButton1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3810
      TabIndex        =   1
      Top             =   4680
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type of Ledger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   4815
      Begin VB.OptionButton Option1 
         Caption         =   "General Ledger"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   2
         Top             =   285
         Width           =   1740
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   540
      Top             =   5940
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   525
      Left            =   4185
      TabIndex        =   3
      Top             =   3030
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   926
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   39
      FullHeight      =   35
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview and set criteria for ledger (GL)."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   480
      Width           =   2850
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LEDGER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   165
      TabIndex        =   5
      Top             =   210
      Width           =   705
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   -75
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frm_LedgerGeneral_option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report

Private Sub LoadSavedReport()
Dim SQL As String
On Error GoTo bad
Report9 = "General"
frmSubsidiaryLedgerViewer.query = "Select left(childAccountcode,3) as Accntcode,fundtype,accountname,accountcode,dated,date_,pariculars,jevno,debit,credit from vw_MP_Final_LedgerSubSidiary where  fundtype = '" & cmb_FundType.Text & "' and  left(childaccountcode,3) = '" & cmb_Accountcode & "' and (dated  between '" & Format(DTPicker2.Value, "MM/dd/yyyy") & "' and '" & Format(DTPicker1.Value, "MM/dd/yyyy") & "') group by left(childaccountcode,3),fundtype,accountname,dated,date_,pariculars,jevno,debit,credit,accountcode   order by accountcode"
frmSubsidiaryLedgerViewer.accnt = "SELECT min(date_) as MinDated,fundtype,sum([Debit]) as SumDebit " & _
      ",sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance " & _
      ",left(ChildAccountCode,3) as Accountcode " & _
  "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where  fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3) = '" & cmb_Accountcode & "' and dated  < '" & Format(DTPicker2.Value, "MM/dd/yyyy") & "' group by fundtype,left(childaccountcode,3)  order by left(childaccountcode,3)"
frmSubsidiaryLedgerViewer.maxdated = "SELECT max(date_) as MinDated,fundtype,sum([Debit]) as SumDebit " & _
      ",sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance " & _
      ",left([ChildAccountCode],3) as Accountcode " & _
  "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where  fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3) = '" & cmb_Accountcode & "' and dated  > '" & Format(DTPicker1.Value, "MM/dd/yyyy") & "' group by fundtype,left(childaccountcode,3)  order by left(childaccountcode,3)"
frmSubsidiaryLedgerViewer.GrndTotal = "SELECT sum([Debit]) As SumDebit,sum([Credit]) as SumCredit,sum([Debit]) - sum([Credit]) as Balance" & _
      ",left([ChildAccountCode],3) as Accountcode,fundtype " & _
        "FROM [fmis].[dbo].[vw_MP_Final_LedgerSubSidiary] where fundtype = '" & cmb_FundType.Text & "' and left(childaccountcode,3)  = '" & cmb_Accountcode & "' group by left(ChildAccountCode,3),fundtype"
frmSubsidiaryLedgerViewer.Show
Exit Sub
bad:
     If err.Number = 364 Then
    MsgBox "No Record Found..", vbInformation, "System Message"
    Else
    MsgBox err.Description
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
        cmb_FundType.ItemData(cc) = opnfund!FundCode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb_FundType.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub cmb_Accountcode_KeyPress(KeyAscii As Integer)
KeyAscii = AutoFind(cmb_Accountcode, KeyAscii, True)
End Sub

Private Sub cmb_FundType_Change()
Call loadChildAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub cmb_FundType_Click()
Call loadAccountcode(cmb_FundType.Text, cmb_Accountcode)
End Sub

Private Sub Command1_Click()

  Call LoadSavedReport
End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
End Sub

Private Sub Label3_Click()

End Sub


