VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmAccountantsAdvice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accountant's Advice Generator"
   ClientHeight    =   9360
   ClientLeft      =   -915
   ClientTop       =   2145
   ClientWidth     =   14385
   Icon            =   "frmAccountantsAdvice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   14385
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8160
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   9360
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
            Picture         =   "frmAccountantsAdvice.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":20FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":3A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":5420
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":6DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":8744
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":A0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":BA68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":D3FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":ED8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":FA6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":1034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":11026
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":11D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":129DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":136BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountantsAdvice.frx":14396
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Log Printed"
      Height          =   525
      Left            =   12060
      TabIndex        =   30
      Top             =   8760
      Width           =   2115
   End
   Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
      Height          =   3225
      Left            =   195
      TabIndex        =   29
      Top             =   4425
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   5689
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
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
   Begin VB.Frame Frame2 
      Height          =   2445
      Left            =   165
      TabIndex        =   18
      Top             =   840
      Width           =   11550
      Begin VB.TextBox txt_Content 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   5025
         TabIndex        =   23
         Top             =   945
         Width           =   6360
      End
      Begin VB.ComboBox cmb_Bank 
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
         Left            =   870
         TabIndex        =   22
         Top             =   1335
         Width           =   3735
      End
      Begin VB.TextBox txt_AddressTo 
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
         Left            =   870
         TabIndex        =   21
         Top             =   825
         Width           =   3735
      End
      Begin VB.TextBox txt_BankAddress 
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
         Left            =   870
         TabIndex        =   20
         Top             =   1845
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   405
         Left            =   9285
         TabIndex        =   19
         Top             =   300
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   143851521
         CurrentDate     =   40354
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name:"
         Height          =   435
         Left            =   135
         TabIndex        =   28
         Top             =   1305
         Width           =   840
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address To:"
         Height          =   435
         Left            =   135
         TabIndex        =   27
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Address:"
         Height          =   435
         Left            =   135
         TabIndex        =   26
         Top             =   1815
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Containing Text:"
         Height          =   195
         Left            =   5010
         TabIndex        =   25
         Top             =   705
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Advice:"
         Height          =   435
         Left            =   8445
         TabIndex        =   24
         Top             =   285
         Width           =   795
      End
   End
   Begin VB.TextBox txt_TotalAmt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   8175
      Width           =   3930
   End
   Begin VB.TextBox txt_CheckNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      TabIndex        =   10
      Top             =   3720
      Width           =   3825
   End
   Begin VB.TextBox txt_Delivered 
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
      Left            =   1110
      TabIndex        =   9
      Top             =   8640
      Width           =   3735
   End
   Begin VB.TextBox txt_Certified 
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
      Left            =   1125
      TabIndex        =   6
      Top             =   8040
      Width           =   3705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Advice"
      Height          =   525
      Left            =   12060
      TabIndex        =   4
      Top             =   8220
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "For the Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11940
      TabIndex        =   1
      Top             =   1320
      Width           =   2490
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   105
         TabIndex        =   2
         Top             =   285
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   143851523
         UpDown          =   -1  'True
         CurrentDate     =   38240
      End
   End
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
      Height          =   4620
      Left            =   12075
      TabIndex        =   0
      Top             =   2370
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   1482
      ButtonWidth     =   1138
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Check No:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12090
      TabIndex        =   17
      Top             =   7200
      Width           =   1185
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7665
      TabIndex        =   16
      Top             =   7860
      Width           =   1380
   End
   Begin VB.Label lbl_OrigSN 
      Caption         =   "SN"
      Height          =   180
      Left            =   5970
      TabIndex        =   14
      Top             =   3105
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lbl_Origdate 
      Caption         =   "origDate"
      Height          =   210
      Left            =   8160
      TabIndex        =   13
      Top             =   3045
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label lbl_OrigUserid 
      Caption         =   "userid"
      Height          =   210
      Left            =   3750
      TabIndex        =   12
      Top             =   3030
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Check No:"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -135
      Top             =   3360
      Width           =   5340
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivered by:"
      Height          =   435
      Left            =   375
      TabIndex        =   8
      Top             =   8610
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Certified Correct:"
      Height          =   435
      Left            =   390
      TabIndex        =   7
      Top             =   8010
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Prepared Accnt's Advice"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12030
      TabIndex        =   3
      Top             =   870
      Width           =   2220
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   8880
      Left            =   11910
      Top             =   540
      Width           =   2490
   End
End
Attribute VB_Name = "frmAccountantsAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim saveflag As Integer
Private Sub cmb_Bank_Click()
txt_BankAddress.Text = GetBankDetail(cmb_bank.ItemData(cmb_bank.ListIndex), "Branch") & ", Agusan del Sur"
End Sub

Private Sub Command1_Click()
Dim sql As String
Dim rec As New ADODB.Recordset
rec.Open "Select * from tblAMIS_logPrintedAccntAdvice where adviceno = '" & List1.List(List1.ListIndex) & "'", opndbaseFMIS, adOpenStatic, adLockPessimistic
If rec.RecordCount <> 0 Then
If MsgBox("This Account Advice is Already Printed, Do you Want to Procced?", vbYesNo, "System Confirmation") = vbYes Then
    
        sql = "Select * from tblAMIS_AccountantAdvice where AdviceNo='" & List1.List(List1.ListIndex) & "' and actioncode=1 order by ChkBankAccntNo,chkno,trnno"
            
        ReportName = "AcctAdvice"
        rptAccntAdvice.Text19.SetText UCase(mydll.AmountToWords(txt_TotalAmt.Text))
        rptAccntAdvice.Database.SetDataSource opndbaseFMIS.Execute(sql)
        rptAccntAdvice.Database.Verify
        Call TransactionLogging("Print Preview", "Accountant Advice", Me.Caption, Winsock1.LocalIP)
        frmViewer.Show 1
        If MsgBox("Do You Want to Log this Account Advice as PRINTED?", vbQuestion + vbYesNo, "System Message") = vbYes Then
           opndbaseFMIS.Execute "Insert into tblAMIS_logPrintedAccntAdvice (adviceno,userid,date_time,logprinted) values ('" & List1.List(List1.ListIndex) & "','" & ActiveUserID & "','" & Now & "',1)"
        End If
    End If
Else
        sql = "Select * from tblAMIS_AccountantAdvice where AdviceNo='" & List1.List(List1.ListIndex) & "' and actioncode=1 order by ChkBankAccntNo,chkno,trnno"
            
        ReportName = "AcctAdvice"
        rptAccntAdvice.Text19.SetText UCase(mydll.AmountToWords(txt_TotalAmt.Text))
        rptAccntAdvice.Database.SetDataSource opndbaseFMIS.Execute(sql)
        rptAccntAdvice.Database.Verify
        Call TransactionLogging("Print Preview", "Accountant Advice", Me.Caption, Winsock1.LocalIP)
        frmViewer.Show 1
        If MsgBox("Do You Want to Log this Account Advice as PRINTED?", vbQuestion + vbYesNo, "System Message") = vbYes Then
           opndbaseFMIS.Execute "Insert into tblAMIS_logPrintedAccntAdvice (adviceno,userid,date_time,logprinted) values ('" & List1.List(List1.ListIndex) & "','" & ActiveUserID & "','" & Now & "',1)"
        End If
End If
End Sub



Private Sub Command2_Click()
frmLogAccountAdvice.Show 1
End Sub

Private Sub DTPicker1_Change()
Call LoadAllAccountantAdvice(DTPicker1.Month, DTPicker1.Year)

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DTPicker2.Value = Date
DTPicker1.Value = Month(Date) & "/1/" & Year(Date)
Call Loadbank(cmb_bank)
Call LoadAllAccountantAdvice(DTPicker1.Month, DTPicker1.Year)
Call SetGrid
Frame2.Enabled = False
Call VisibleToolbar(3)
End Sub
Private Function VerifyAllEntries() As Boolean
Dim cc As Integer
Dim tmpFirstBankID As String

For cc = 1 To MSHFlexGrid1.Rows - 1
    If cc = 1 Then
        tmpFirstBankID = MSHFlexGrid1.TextMatrix(cc, 5)
        VerifyAllEntries = True
    Else
        If tmpFirstBankID = MSHFlexGrid1.TextMatrix(cc, 5) Then
            VerifyAllEntries = True
        Else
            VerifyAllEntries = False
            Exit For
        End If
    End If
Next cc
End Function
Private Function GetLatestSN(ByVal MonthNo As Integer, ByVal YearOf As Integer) As Long
Dim opnSN As New ADODB.Recordset

opnSN.Open "Select SN from tblAMIS_AccountantAdvice where year(dateadvice)=" & YearOf & " and month(dateadvice)=" & MonthNo & " and actioncode=1 group by SN order by SN desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnSN.RecordCount <> 0 Then
    GetLatestSN = CLng(opnSN!sn) + 1
Else
    GetLatestSN = 1
End If
opnSN.Close
Set opnSN = Nothing
End Function
Private Sub LoadAllAccountantAdvice(ByVal MonthNo As Integer, ByVal YearOf As Integer)
Dim opnaccnt As New ADODB.Recordset

List1.Clear
opnaccnt.Open "Select AdviceNo from tblAMIS_AccountantAdvice where year(dateadvice)=" & YearOf & " and month(dateadvice)=" & MonthNo & " and actioncode=1 group by adviceno order by adviceno desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnaccnt.RecordCount <> 0 Then
    Do Until opnaccnt.EOF
        List1.AddItem (opnaccnt!adviceno)
        opnaccnt.MoveNext
    Loop
End If
opnaccnt.Close
Set opnaccnt = Nothing
Label11.Caption = List1.ListCount & " Accnt. Advice/s Found"
End Sub
Private Sub Save(ByVal saveflag As Integer)
Dim NewSN As Long
Dim NewAdviceNo As String
Dim cc As Integer
Select Case saveflag
    Case 1 'new
        NewSN = GetLatestSN(DTPicker2.Month, DTPicker2.Year)
        NewAdviceNo = DTPicker2.Year & "-" & Format(DTPicker2.Month, "00") & "-" & Format(NewSN, "00000")
        For cc = 1 To MSHFlexGrid1.Rows - 1
            opndbaseFMIS.Execute "Insert into tblAMIS_AccountantAdvice(SN,AdviceNo,DateAdvice,AddressTo,BankName,BankAddress,ContainingText,ChkBankAccntNo, " & _
                        " ChkNo,ChkDate,ChkPayee,ChkAmount,CertifiedCorrectBy,DeliveredBy,actioncode,UserID,DateTimeEntered,ManualEntry) " & _
                        " values(" & NewSN & ",'" & NewAdviceNo & "','" & DTPicker2.Value & "','" & txt_AddressTo.Text & "','" & cmb_bank.List(cmb_bank.ListIndex) & "', " & _
                        " '" & txt_BankAddress.Text & "','" & txt_Content.Text & "','" & MSHFlexGrid1.TextMatrix(cc, 0) & "','" & MSHFlexGrid1.TextMatrix(cc, 1) & "', " & _
                        " '" & MSHFlexGrid1.TextMatrix(cc, 2) & "','" & Replace(MSHFlexGrid1.TextMatrix(cc, 3), "'", "''") & "'," & CCur(MSHFlexGrid1.TextMatrix(cc, 4)) & ", " & _
                        " '" & txt_Certified.Text & "','" & txt_Delivered.Text & "',1,'" & ActiveUserID & "','" & Now & "'," & val(MSHFlexGrid1.TextMatrix(cc, 7)) & ")"
        Next cc
        MsgBox "Saving Accountant's Advice, Successful!", vbInformation, "System Information"
        Call LoadAllAccountantAdvice(DTPicker1.Month, DTPicker1.Year)
        Call Clear
        Call SetGrid
        saveflag = 0
    Case 2 'edit
        
        'Editing the Original Records
        opndbaseFMIS.Execute "Update tblAMIS_AccountantAdvice set actioncode=2,userid='" & Trim(lbl_OrigUserid.Caption) & "," & Trim(ActiveUserID) & "',datetimeentered='" & lbl_Origdate.Caption & "," & Now & "' where adviceno='" & List1.List(List1.ListIndex) & "'"
        
        For cc = 1 To MSHFlexGrid1.Rows - 1
            opndbaseFMIS.Execute "Insert into tblAMIS_AccountantAdvice(SN,AdviceNo,DateAdvice,AddressTo,BankName,BankAddress,ContainingText,ChkBankAccntNo, " & _
                        " ChkNo,ChkDate,ChkPayee,ChkAmount,CertifiedCorrectBy,DeliveredBy,actioncode,UserID,DateTimeEntered,ManualEntry) " & _
                        " values(" & lbl_OrigSN.Caption & ",'" & List1.List(List1.ListIndex) & "','" & DTPicker2.Value & "','" & txt_AddressTo.Text & "','" & cmb_bank.List(cmb_bank.ListIndex) & "', " & _
                        " '" & txt_BankAddress.Text & "','" & txt_Content.Text & "','" & MSHFlexGrid1.TextMatrix(cc, 0) & "','" & Trim(MSHFlexGrid1.TextMatrix(cc, 1)) & "', " & _
                        " '" & MSHFlexGrid1.TextMatrix(cc, 2) & "','" & MSHFlexGrid1.TextMatrix(cc, 3) & "'," & CCur(MSHFlexGrid1.TextMatrix(cc, 4)) & ", " & _
                        " '" & txt_Certified.Text & "','" & txt_Delivered.Text & "',1,'" & Trim(ActiveUserID) & "','" & Now & "'," & val(MSHFlexGrid1.TextMatrix(cc, 7)) & ")"
        Next cc
        
        
        MsgBox "Updating Accountant's Advice, Successful!", vbInformation, "System Information"
        Call LoadAllAccountantAdvice(DTPicker1.Month, DTPicker1.Year)
        Call Clear
        Call SetGrid
        saveflag = 0
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set frmAccountantsAdvice = Nothing
End Sub
Private Sub SetGrid()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8

MSHFlexGrid1.TextMatrix(0, 0) = "Bank Accnt. No."
MSHFlexGrid1.TextMatrix(0, 1) = "Check No."
MSHFlexGrid1.TextMatrix(0, 2) = "Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Payee"
MSHFlexGrid1.TextMatrix(0, 4) = "Amount"
MSHFlexGrid1.TextMatrix(0, 5) = "BankID"
MSHFlexGrid1.TextMatrix(0, 6) = "trnno"
MSHFlexGrid1.TextMatrix(0, 7) = "manual"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 5000
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 0
MSHFlexGrid1.ColWidth(6) = 0
MSHFlexGrid1.ColWidth(7) = 0

End Sub
Private Sub Clear()
lbl_OrigUserid.Caption = ""
lbl_Origdate.Caption = ""
lbl_OrigSN.Caption = ""

txt_AddressTo.Text = ""
txt_Content.Text = ""
txt_BankAddress.Text = ""
txt_Certified.Text = ""
txt_Delivered.Text = ""
txt_CheckNo.Text = ""
txt_TotalAmt.Text = ""
Call Loadbank(cmb_bank)
End Sub
Private Sub LoadBack(ByVal adviceno As String)
Dim opnChecks As New ADODB.Recordset

opnChecks.Open "Select * from tblAMIS_AccountantAdvice where adviceno='" & adviceno & "' and actioncode=1 order by trnno", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnChecks.RecordCount <> 0 Then
    Call SetGrid
    txt_CheckNo.Text = "" 'Clearing the CheckNo TextBox
    
    txt_AddressTo.Text = opnChecks!AddressTo
    txt_Content.Text = opnChecks!ContainingText
    txt_Certified.Text = opnChecks!CertifiedCorrectBy
    txt_Delivered.Text = opnChecks!DeliveredBy
    Call Loadbank(cmb_bank)
    cmb_bank.ListIndex = GetIndex(cmb_bank, opnChecks!BankName)
    txt_BankAddress.Text = opnChecks!BankAddress
    lbl_Origdate.Caption = opnChecks!datetimeentered
    lbl_OrigUserid.Caption = opnChecks!UserID
    lbl_OrigSN.Caption = opnChecks!sn
    DTPicker2.Value = opnChecks!DateAdvice
    Do Until opnChecks.EOF
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = opnChecks!ChkBankAccntNo
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = opnChecks!ChkNo
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 2) = opnChecks!ChkDate
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 3) = opnChecks!ChkPayee
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 4) = Format(opnChecks!ChkAmount, "###,##0.00")
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 5) = GetBankIDbyBankAccntNo(opnChecks!ChkBankAccntNo)
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 6) = opnChecks!Trnno
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 7) = IIf(IsNull(opnChecks!ManualEntry), 0, 1)
        
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    opnChecks.MoveNext
    Loop
    txt_TotalAmt.Text = Format(GetTotalEnteredAmtInGrid(MSHFlexGrid1, 4, 1), "###,##0.00")
    Toolbar1.Buttons(3).Visible = False
Else
    Call Clear
End If
opnChecks.Close
Set opnChecks = Nothing

End Sub

Private Sub List1_Click()
Call LoadBack(List1.List(List1.ListIndex))
saveflag = 0
Frame2.Enabled = False
Call VisibleToolbar(3)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cc As Integer

If Shift = 2 And KeyCode = vbKeyDelete Then
    If MsgBox("Are you sure want to Delete this Advice?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
        
        For cc = 1 To MSHFlexGrid1.Rows - 1
            opndbaseFMIS.Execute "Update tblAMIS_AccountantAdvice set actioncode=4,DateTimeEntered='" & lbl_Origdate.Caption & "," & Now & "',UserId='" & lbl_OrigUserid.Caption & "," & ActiveUserID & "' where ChkNo='" & MSHFlexGrid1.TextMatrix(cc, 1) & "'"
            opndbaseFMIS.Execute "Update tblAMIS_AcctntAdviceAuthorizedCheck set actioncode=4 where checkNo='" & MSHFlexGrid1.TextMatrix(cc, 1) & "'"
        Next cc
        
        MsgBox "Deleting Accountant Advice, Successful!", vbInformation, "System Information"
        Call LoadAllAccountantAdvice(DTPicker1.Month, DTPicker1.Year)
        Call Clear
        Call SetGrid

    End If
End If
End Sub



Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyDelete Then
    
    If saveflag > 0 Then
        If val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 6)) > 0 Then
            If MsgBox("Are you sure want to Remove this Check among others?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
                
                opndbaseFMIS.Execute "Update tblAMIS_AccountantAdvice set actioncode=4,DateTimeEntered='" & lbl_Origdate.Caption & "," & Now & "',UserId='" & lbl_OrigUserid.Caption & "," & ActiveUserID & "' where trnno=" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 6) & ""
                opndbaseFMIS.Execute "Update tblAMIS_AcctntAdviceAuthorizedCheck set actioncode=4 where checkNo='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & "'"
                MsgBox "Deleting Check Entry, Successful!", vbInformation, "System Information"
                Call LoadBack(List1.List(List1.ListIndex))
                saveflag = 0
                Frame2.Enabled = False
            
            End If
        
        Else 'This is For Unsaved or Newly Created Advice...
            If Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)) > 0 Then
                If MSHFlexGrid1.Rows = 1 Then
                    If MsgBox("are you sure want to remove this check?", vbQuestion + vbYesNo, "system confirmation") = vbYes Then
                        opndbaseFMIS.Execute "Update tblAMIS_AcctntAdviceAuthorizedCheck set actioncode=4 where checkNo='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & "'"
                        Call SetGrid
                    End If
                Else
                    If MsgBox("are you sure want to remove this check among others?", vbQuestion + vbYesNo, "system confirmation") = vbYes Then
                        opndbaseFMIS.Execute "Update tblAMIS_AcctntAdviceAuthorizedCheck set actioncode=4 where checkNo='" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) & "'"
                        Call MoveData(MSHFlexGrid1, MSHFlexGrid1.Row + 1, CopyGridDataDownWard(MSHFlexGrid1, MSHFlexGrid1.Row + 1), "Delete")
                    End If
                End If
            End If
        End If
    End If

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpcheck As Variant
Select Case Button.Index
    Case 1 'new
    Call VisibleToolbar(1)
        saveflag = 1
        Call Clear
        Call SetGrid
        Frame2.Enabled = True
        txt_AddressTo.Text = "The Bank Manager"
        txt_Content.Text = "Please be informed that the following checks were issued by this office:"
        txt_Certified.Text = "JOY T. DY-LAGAT, CPA"
        txt_Delivered.Text = "MERCEDITA Q. PALAD"
        DTPicker1.Value = Now
        DTPicker2.Value = Now
    Case 2 'save
        If VerifyAllEntries = True Then
            Call Save(saveflag)
            Frame2.Enabled = False
        Else
            MsgBox "One or More of the Check/s Entered were/was Inconsistent with the Bank Entity of the First Check!" & Chr(13) & "Please Verify your Entries...", vbInformation, "System Information"
        End If
        Call VisibleToolbar(1)
    Case 3 'edit
        saveflag = 2
        Frame2.Enabled = False
        Call VisibleToolbar(2)
    Case 4 'Search
        tmpcheck = InputBox("CheckNo :", "Verify Check Status")
        If Len(Trim(tmpcheck)) <> 0 Then
            If VerifyCheckNo(tmpcheck) = 1 Then
                MsgBox "Check No. was already Prepared with Accountants Advice!" & Chr(13) & Chr(13) & "Under Advice No.: " & GetAccntAdviceNo(tmpcheck), vbInformation, "System Information"
            ElseIf VerifyCheckNo(tmpcheck) <= 4 Then
                MsgBox "Check No. is Still For Signature!", vbInformation, "System Information"
            ElseIf VerifyCheckNo(tmpcheck) <= 6 Then
                MsgBox "Check No. is for Accountants Advice Preparation!", vbInformation, "System Information"
            Else
                MsgBox "Check No. was not yet Prepared by the Treasury Office!", vbInformation, "System Information"
            End If
        End If
        
    Case 5 'close
        Unload Me
    Case 6
    Call VisibleToolbar(3)
        saveflag = 0
        Call Clear
        Call SetGrid
        Frame2.Enabled = True
        txt_AddressTo.Text = "The Bank Manager"
        txt_Content.Text = "Please be informed that the following checks were issued by this office:"
        txt_Certified.Text = "JOY T. DY-LAGAT, CPA"
        txt_Delivered.Text = "MERCEDITA Q. PALAD"
        DTPicker1.Value = Now
        DTPicker2.Value = Now
        Toolbar1.Buttons(3).Visible = False
End Select
End Sub
Private Sub VisibleToolbar(ByVal TYP As Integer)
    If TYP = 1 Then 'new
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(2).Caption = "Save"
        txt_CheckNo.Enabled = True
    ElseIf TYP = 2 Then 'update
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(2).Caption = "Update"
        txt_CheckNo.Enabled = True
    Else
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = True
        Toolbar1.Buttons(4).Visible = True
        Toolbar1.Buttons(5).Visible = True
        Toolbar1.Buttons(6).Visible = False
        Toolbar1.Buttons(2).Caption = "Save"
        txt_CheckNo.Enabled = False
    End If
End Sub
Private Function IncludedInTheSelection(ByVal checkno As String) As Boolean
Dim cc As Integer

IncludedInTheSelection = False 'Initialized Value

For cc = 1 To MSHFlexGrid1.Rows - 1
    If Len(Trim(MSHFlexGrid1.TextMatrix(cc, 1))) <> 0 Then
        If checkno = MSHFlexGrid1.TextMatrix(cc, 1) Then
            IncludedInTheSelection = True
            Exit For
        End If
    End If
Next cc
End Function
Private Function VerifyBankOfChk(ByVal checkno As String, ByVal BankIDNo As Integer) As Boolean
Dim opnChkBank As New ADODB.Recordset

VerifyBankOfChk = False

opnChkBank.Open "SELECT  tblCMS_CDCheckRoutine.CheckNo, tblCMS_CDPreparedCheck.CompositionCode, vw_DepositoryBank.BankName, " & _
                " vw_DepositoryBank.BankID, vw_DepositoryBank.BankIDNo " & _
                " FROM tblCMS_CDCheckRoutine INNER JOIN " & _
                " tblCMS_CDPreparedCheck ON tblCMS_CDCheckRoutine.CheckNo = tblCMS_CDPreparedCheck.CheckNo INNER JOIN " & _
                " vw_DepositoryBank ON tblCMS_CDPreparedCheck.CompositionCode = vw_DepositoryBank.FMISAccountCode " & _
                " WHERE (tblCMS_CDCheckRoutine.Actioncode = 1) AND (tblCMS_CDPreparedCheck.actioncode = 1) AND " & _
                " (tblCMS_CDCheckRoutine.CheckNo = '" & checkno & "')", opndbaseFMIS, adOpenStatic, adLockOptimistic

If opnChkBank.RecordCount <> 0 Then
    If BankIDNo = opnChkBank!BankIDNo Then
        VerifyBankOfChk = True
    End If
End If
opnChkBank.Close
Set opnChkBank = Nothing

End Function
Private Function GetCheckDetail(ByVal checkno As String, ByVal FldName As String) As Variant
Dim opnChkDetails As New ADODB.Recordset

opnChkDetails.Open "Select * from tblCMS_CDPreparedCheck where checkno='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnChkDetails.RecordCount <> 0 Then
    Select Case FldName
        Case "BankAccntNo"
            GetCheckDetail = GetAccountNameCode(opnChkDetails!compositioncode, "AccCode")
        Case "CheckDate"
            GetCheckDetail = opnChkDetails!CheckDate
        Case "Payee"
            GetCheckDetail = opnChkDetails!claimantname
        Case "Amount"
            GetCheckDetail = opnChkDetails!NetAmount
        Case "BankID"
            GetCheckDetail = GetAccountNameCode(opnChkDetails!compositioncode, "BankID")
    End Select
End If
opnChkDetails.Close
Set opnChkDetails = Nothing
End Function
Private Function Verify() As Boolean
If Len(Trim(txt_AddressTo.Text)) <> 0 And Len(Trim(txt_BankAddress.Text)) <> 0 And Len(cmb_bank.Text) <> 0 Then
    Verify = True
Else
    Verify = False
End If
End Function
Private Sub txt_CheckNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If CheckIfExists("SELECT dvno FROM [fmis].[dbo].[tblAMIS_Logtrans] where Checkno = '" & txt_CheckNo.Text & "' and status in (13,22,14,15) and Actioncode = 1") = False Then
    MsgBox "Please Logged out fisrt in Admin/PGO to proceed Accountant's Advice", vbInformation + vbCritical, "System Warning"
    Exit Sub
End If

    If saveflag > 0 Then 'only New and Edit Mode will be allowed here.....
        If Verify = True Then
            If VerifyCheckNo(txt_CheckNo.Text) = 0 Then
                If MsgBox("Check No. was not yet Prepared by the Treasury Office!" & Chr(13) & "Would you like to Add this Check No in the Selection List?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
                    frmAccntAdviceSpecial.txt_CheckNo.Text = txt_CheckNo.Text
                    frmAccntAdviceSpecial.txt_BankName.Text = cmb_bank.Text
                    frmAccntAdviceSpecial.Show vbModal
                'In this line, code procedure for Check Authorization be included in the Preparation of Advice
                End If
            
            ElseIf VerifyCheckNo(txt_CheckNo.Text) = 1 Then
                MsgBox "Check No. was already Prepared with Accountants Advice!", vbInformation, "System Information"
'            ElseIf VerifyCheckNo(txt_CheckNo.Text) <= 3 Then
'                MsgBox "Check No. is Still For Signature!", vbInformation, "System Information"
            ElseIf VerifyCheckNo(txt_CheckNo.Text) <= 6 Then
                'Executing Final Verification From the List of Selected CheckNo
                If IncludedInTheSelection(txt_CheckNo.Text) = False Then
                    If VerifyBankOfChk(txt_CheckNo.Text, cmb_bank.ItemData(cmb_bank.ListIndex)) = True Then
                        'Include in the Grid and Get the Check Details..................
                        
                        'check if have JEV entries
                        'code here
                        
                            If Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0)) Then
                                MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1 'Adding New Row
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = GetCheckDetail(txt_CheckNo.Text, "BankAccntNo")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = txt_CheckNo.Text
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 2) = GetCheckDetail(txt_CheckNo.Text, "CheckDate")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 3) = GetCheckDetail(txt_CheckNo.Text, "Payee")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 4) = Format(GetCheckDetail(txt_CheckNo.Text, "Amount"), "###,##0.00")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 5) = GetCheckDetail(txt_CheckNo.Text, "BankID")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 7) = 0
                            Else
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = GetCheckDetail(txt_CheckNo.Text, "BankAccntNo")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = txt_CheckNo.Text
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 2) = GetCheckDetail(txt_CheckNo.Text, "CheckDate")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 3) = GetCheckDetail(txt_CheckNo.Text, "Payee")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 4) = Format(GetCheckDetail(txt_CheckNo.Text, "Amount"), "###,##0.00")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 5) = GetCheckDetail(txt_CheckNo.Text, "BankID")
                                MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 7) = 0
                            End If
                        
                        txt_TotalAmt.Text = Format(GetTotalEnteredAmtInGrid(MSHFlexGrid1, 4, 1), "###,##0.00")
                    Else
                        MsgBox "The Specied Check No. is not Belong to " & cmb_bank.Text, vbInformation, "System Information"
                    End If
                Else
                    MsgBox "Check No. was already included in the selected Item!", vbInformation, "System Information"
                End If
            End If
            txt_CheckNo.SelStart = 0
            txt_CheckNo.SelLength = Len(txt_CheckNo.Text)
            txt_CheckNo.SetFocus
        End If
    End If
End If
End Sub
