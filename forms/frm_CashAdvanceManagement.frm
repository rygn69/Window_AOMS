VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_CashAdvanceManagement 
   Caption         =   "Cash Advance Management"
   ClientHeight    =   9420
   ClientLeft      =   735
   ClientTop       =   1215
   ClientWidth     =   14415
   Icon            =   "frm_CashAdvanceManagement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14415
   Begin VB.CheckBox Check1 
      Caption         =   "Check Date"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3135
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   120
         TabIndex        =   10
         Top             =   915
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Year"
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
            Left            =   240
            TabIndex        =   12
            Tag             =   "1"
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Month-Year"
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
            Left            =   240
            TabIndex        =   11
            Tag             =   "3"
            Top             =   795
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPYear 
            Height          =   375
            Left            =   1080
            TabIndex        =   13
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy"
            Format          =   98762755
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
         Begin MSComCtl2.DTPicker DTpMY 
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   1035
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM yyyy"
            Format          =   98762755
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
      End
      Begin VB.ComboBox cmb_FundType 
         Appearance      =   0  'Flat
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
         Left            =   120
         TabIndex        =   8
         Top             =   555
         Width           =   2940
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Special Accounts:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_RecordID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      ToolTipText     =   "Type only CN No. then press Enter"
      Top             =   1200
      Width           =   3045
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3600
      Top             =   9000
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1365
      Top             =   9615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashAdvanceManagement.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashAdvanceManagement.frx":1EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashAdvanceManagement.frx":3FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashAdvanceManagement.frx":5280
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashAdvanceManagement.frx":810A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   1482
      ButtonWidth     =   2117
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   5520
         Top             =   0
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
               Picture         =   "frm_CashAdvanceManagement.frx":DE6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":F7FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":11190
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":12B22
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":144B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":15E46
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":177D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1916A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1AAFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1C490
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1D16C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1DA4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1E728
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":1F404
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":200E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":20DBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_CashAdvanceManagement.frx":21A98
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   450
         Left            =   11760
         TabIndex        =   4
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
         _Version        =   393216
         FullWidth       =   32
         FullHeight      =   30
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_details 
      Height          =   5610
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   9895
      _Version        =   393216
      FixedCols       =   0
      ForeColorFixed  =   128
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Load"
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
      Image           =   "frm_CashAdvanceManagement.frx":22374
      cBack           =   16777215
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   11595
      TabIndex        =   3
      Top             =   9135
      Width           =   2730
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9240
      TabIndex        =   2
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   9060
      Width           =   480
   End
End
Attribute VB_Name = "frm_CashAdvanceManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadLOCbySQL()
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select * from MPfunc_LoadliquidationOfCAbyDate('" & cmb_FundType.Text & "','" & DTPicker1.Month & "','" & DTPicker1.Year & "') order by cast(transactiondate as date)")
    Set grd_details.DataSource = rec
With grd_details
    .ColWidth(0) = 1000 'Date
    .ColWidth(1) = 1200 'Dvno
    .ColWidth(2) = 3000 'Name
    .ColWidth(3) = 8000 'Particular
    .ColWidth(4) = 1000 'Gross Amount
    .ColWidth(5) = 900  'Status
    .ColWidth(6) = 1500 'JEV number
End With
End Sub
Private Sub cmd_post_Click()
On Error GoTo bad
Dim x As Long
Dim tmp As Long
If MsgBox("Are you sure Do you want to POst?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    For x = 1 To grd_details.Rows - 1
        If IsFormatCorrect(grd_details.TextMatrix(x, 6)) = True Then
        tmp = Val(Right(grd_details.TextMatrix(x, 6), 5))
            Call GEtCompleteJEVDetails(Trim(grd_details.TextMatrix(x, 1)), whatfield, grd_details.TextMatrix(x, 0), "", "" _
            , grd_details.TextMatrix(x, 3), grd_details.TextMatrix(x, 6), "", 0, 0, 0, 0, 4, "", grd_details.TextMatrix(x, 1), "", "", "" _
            , "", "", "", tmp, DatePost, "")
            'Updating Accounting REcord...
            opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & grd_details.TextMatrix(x, 6) & "', " & _
                " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                " JEVDate='" & DatePost & "',transtype = 4 where DVNo='" & grd_details.TextMatrix(x, 1) & "'"
        End If
        DoEvents
    Next x
End If
Exit Sub
bad:
Call LoadErr(err.description, Me.name, err.description)
End Sub
Private Sub cmd_Mass_Click()
Dim rec As New ADODB.Recordset
Dim x As Long
Dim LastJEVSNno As Long

If grd_details.Rows = 1 Then
    MsgBox "No Record to Post", vbCritical, "System Message"
    Exit Sub
End If
    
    JevOk = False
    frmPOstdate.Show 1
    If JevOk = True Then
    Label13.Caption = "JEV Numbering..."
    Label13.Refresh
    Animation1.Visible = True
    Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
    Animation1.Play
        
        rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = '4',@jevyeardate = '" & DatePost & "' ,@fundtype = '" & cmb_FundType.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        LastJEVSNno = rec.Fields!MAXJEVSERIES
        rec.Close
        
        With grd_details
            For x = 1 To .Rows - 1
                If .TextMatrix(x, 5) = "Approved" Then
                    .TextMatrix(x, 6) = cmb_FundType.ItemData(cmb_FundType.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & "04" & "-" & Format(LastJEVSNno, "0000")
                    LastJEVSNno = LastJEVSNno + 1
                Else
                   .TextMatrix(x, 6) = "Not yet Ready"
                End If
            Next x
        End With
        
    Animation1.Stop
    Animation1.Close
    Animation1.Visible = False
    Label13.Caption = ""
    Else
    MsgBox "Cannot Generate the System JEV Number,If you cancel to Set the Date", vbInformation, "System Message"
    End If
    
End Sub



Private Sub Command1_Click()
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
    Call LoadLOCbySQL
Animation1.Stop
Animation1.Close
Animation1.Visible = False
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
DTPicker1.Value = Now
End Sub


Private Sub Form_Resize()
On Error Resume Next
grd_details.Width = Me.ScaleWidth - 350
  grd_details.Height = Me.ScaleHeight - grd_details.Top - 500
End Sub

Private Sub grd_details_Click()
ActiveFormCaller = Me.name
If grd_details.TextMatrix(grd_details.Row, 5) <> "Not Approved" Then
    ForTheGridRowNo = grd_details.Row
    'Load_Offices   '---loads offices to the OFFICE combo - RICHARD
    'Load_FundTypes '---loads functypes to the FUNDTYPE combo - RICHARD
   ' If Len(grd_details.TextMatrix(grd_details.Row, 14)) <> 0 Then 'Kung Naa nay JEV No
        frmJEVNumberingAssignment_New.IsSaveAccntng = False
        frmJEVNumberingAssignment_New.txt_JEVNO.Text = grd_details.TextMatrix(grd_details.Row, 6)
        frmJEVNumberingAssignment_New.txt_DVNo = grd_details.TextMatrix(grd_details.Row, 1)
        frmJEVNumberingAssignment_New.Show 1
Else
    MsgBox "Transaction must be Approved to Proceed JEV numbering", vbInformation, "System Message"
End If
Exit Sub
bad:
Call LoadErr(err.Number, Me.name, err.description)
End Sub

